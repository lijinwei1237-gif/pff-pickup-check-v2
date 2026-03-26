import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

st.set_page_config(page_title="PFF Pickup Check Tool", layout="wide")

# =========================
# 页面样式
# =========================
st.markdown("""
<style>
.block-container {
    padding-top: 1rem;
    padding-bottom: 1.5rem;
}
h1 {
    margin-bottom: 0.2rem !important;
}
.small-note {
    font-size: 12px;
    color: #666;
    margin-bottom: 1rem;
}
.metric-label {
    font-size: 12px;
    color: #666;
    margin-bottom: 4px;
}
.metric-value {
    font-size: 38px;
    font-weight: 700;
    line-height: 1.05;
}
.run-btn button {
    width: 100%;
    background-color: #ff4b4b !important;
    color: white !important;
    border: none !important;
    border-radius: 4px !important;
    height: 42px !important;
    font-weight: 600 !important;
}
.stDownloadButton button {
    width: 100%;
}
</style>
""", unsafe_allow_html=True)

PFF_ZONES = {"F3", "F4", "F5", "F11", "F12", "F13"}

# =========================
# 工具函数
# =========================
def clean_box_no(x):
    if pd.isna(x):
        return None
    s = str(x).strip().upper().replace(" ", "")
    m = re.search(r'B\d{13}', s)
    return m.group(0) if m else None

def standardize_columns(df):
    df = df.copy()
    df.columns = (
        pd.Index(df.columns)
        .astype(str)
        .str.strip()
        .str.replace("\n", "", regex=False)
        .str.replace("\r", "", regex=False)
    )
    return df

def normalize_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def normalize_upper_text(x):
    return normalize_text(x).upper()

def contains_unpicked(text):
    return "未领" in normalize_text(text)

def extract_f_zone(x):
    s = normalize_upper_text(x).replace(" ", "")
    m = re.search(r'F\d+', s)
    return m.group(0) if m else ""

def ratio_match(series, pattern_func, sample_size=200):
    s = series.dropna().astype(str).head(sample_size)
    if len(s) == 0:
        return 0
    return sum(pattern_func(v) for v in s) / len(s)

def is_box_value(x):
    s = str(x).strip().upper().replace(" ", "")
    return bool(re.fullmatch(r'B\d{13}', s))

def is_route_stop_value(x):
    s = str(x).strip().upper().replace(" ", "")
    return bool(re.fullmatch(r'F\d+-\d+', s))

def is_route_value(x):
    s = str(x).strip().upper().replace(" ", "")
    return bool(re.fullmatch(r'F\d+', s))

def is_numeric_value(x):
    s = str(x).strip()
    return bool(re.fullmatch(r'\d+', s))

def find_best_col_by_name(columns, candidates):
    cols = list(columns)

    def norm(s):
        return str(s).strip().lower().replace("_", "").replace("-", "").replace(" ", "")

    norm_map = {c: norm(c) for c in cols}

    for cand in candidates:
        cand_raw = str(cand).strip().lower()
        for c in cols:
            if str(c).strip().lower() == cand_raw:
                return c

    for cand in candidates:
        cand_norm = norm(cand)
        for c, c_norm in norm_map.items():
            if c_norm == cand_norm:
                return c

    for cand in candidates:
        cand_norm = norm(cand)
        for c, c_norm in norm_map.items():
            if cand_norm in c_norm:
                return c

    return None

def try_read_excel(uploaded_file, header_option):
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, header=header_option)

# =========================
# Gary 表识别
# =========================
def detect_gary_columns(df):
    cols = list(df.columns)

    box_col = find_best_col_by_name(cols, [
        "box no", "box", "bag", "bag no", "箱号", "袋号"
    ])

    route_stop_col = find_best_col_by_name(cols, [
        "route-stop", "route stop", "stop"
    ])

    address_col = find_best_col_by_name(cols, [
        "address", "地址"
    ])

    count_col = find_best_col_by_name(cols, [
        "count", "数量", "件数", "qty"
    ])

    route_col = find_best_col_by_name(cols, ["route", "路线"])

    if route_col is None:
        scores = []
        for c in cols:
            if c in [box_col, route_stop_col, address_col, count_col]:
                continue
            score = ratio_match(df[c], is_route_value)
            scores.append((c, score))
        best = max(scores, key=lambda x: x[1], default=(None, 0))
        if best[1] >= 0.3:
            route_col = best[0]

    if box_col is None:
        scores = [(c, ratio_match(df[c], is_box_value)) for c in cols]
        best = max(scores, key=lambda x: x[1], default=(None, 0))
        if best[1] >= 0.3:
            box_col = best[0]

    if route_stop_col is None:
        scores = [(c, ratio_match(df[c], is_route_stop_value)) for c in cols if c != box_col]
        best = max(scores, key=lambda x: x[1], default=(None, 0))
        if best[1] >= 0.3:
            route_stop_col = best[0]

    if count_col is None:
        scores = [
            (c, ratio_match(df[c], is_numeric_value))
            for c in cols if c not in [box_col, route_stop_col, route_col]
        ]
        best = max(scores, key=lambda x: x[1], default=(None, 0))
        if best[1] >= 0.5:
            count_col = best[0]

    if route_col == route_stop_col:
        route_col = None

    return {
        "box_col": box_col,
        "route_stop_col": route_stop_col,
        "route_col": route_col,
        "address_col": address_col,
        "count_col": count_col
    }

def load_onsite_file(uploaded_file):
    for h in [0, 1, 2, 3, 4, 5]:
        try:
            df = try_read_excel(uploaded_file, h)
            df = standardize_columns(df)
            detected = detect_gary_columns(df)

            if detected["box_col"] is not None:
                result = df.copy()
                result["box no"] = result[detected["box_col"]]
                result["box_no_clean"] = result["box no"].apply(clean_box_no)
                result = result[result["box_no_clean"].notna()].copy()

                if detected["route_stop_col"]:
                    result["route-stop"] = result[detected["route_stop_col"]].astype(str)
                else:
                    result["route-stop"] = ""

                if detected["route_col"]:
                    result["route"] = result[detected["route_col"]].astype(str)
                else:
                    result["route"] = result["route-stop"].apply(extract_f_zone)

                if detected["address_col"]:
                    result["address"] = result[detected["address_col"]]
                else:
                    result["address"] = ""

                if detected["count_col"]:
                    result["count"] = result[detected["count_col"]]
                else:
                    result["count"] = 0

                result = result.drop_duplicates(subset=["box_no_clean"])
                return result
        except Exception:
            continue

    raise ValueError("无法识别 Gary 现场箱号表。")

# =========================
# 任务表识别
# =========================
def looks_like_task_df(df):
    needed = {"运单号", "领取状态", "派送方", "快递员区域名称", "快递员路线", "箱号"}
    return len(needed.intersection(set(df.columns))) >= 4

def load_task_file(uploaded_file):
    for h in [0, 1, 2, 3, 4, 5]:
        try:
            df = try_read_excel(uploaded_file, h)
            df = standardize_columns(df)

            if looks_like_task_df(df):
                needed_cols = ["运单号", "领取状态", "派送方", "快递员区域名称", "快递员路线", "箱号"]
                missing = [c for c in needed_cols if c not in df.columns]
                if missing:
                    continue

                result = df.copy()
                result["box_no_clean"] = result["箱号"].apply(clean_box_no)
                result["vendor_clean"] = result["派送方"].apply(normalize_upper_text)
                result["pickup_status_clean"] = result["领取状态"].apply(normalize_text)
                result["zone_clean"] = result["快递员区域名称"].apply(extract_f_zone)
                result["route_clean"] = result["快递员路线"].apply(extract_f_zone)

                result = result[result["box_no_clean"].notna()].copy()
                result = result[result["vendor_clean"].str.contains("PFF", na=False)].copy()
                result = result[result["zone_clean"].isin(PFF_ZONES)].copy()
                result = result[result["pickup_status_clean"].apply(contains_unpicked)].copy()

                return result
        except Exception:
            continue

    raise ValueError("无法识别任务表。")

# =========================
# 比对逻辑
# =========================
def compare_files(onsite_df, task_df):
    keep_cols = [
        "运单号", "箱号", "box_no_clean", "派送方", "领取状态",
        "快递员区域名称", "快递员路线", "zone_clean", "route_clean"
    ]
    keep_cols = [c for c in keep_cols if c in task_df.columns]

    matched_df = onsite_df.merge(
        task_df[keep_cols].copy(),
        on="box_no_clean",
        how="inner"
    )

    matched_df["issue_flag"] = "现场有但任务表仍未领"
    display_df = matched_df.drop_duplicates(subset=["box_no_clean"]).copy()
    return matched_df, display_df

def build_route_summary(display_df):
    if display_df.empty:
        return pd.DataFrame(columns=["route", "issue_boxes", "affected_packages"])

    temp = display_df.copy()
    temp["count_num"] = pd.to_numeric(temp["count"], errors="coerce").fillna(0)

    route_summary = (
        temp.groupby("route", dropna=False)
        .agg(
            issue_boxes=("box_no_clean", "nunique"),
            affected_packages=("count_num", "sum")
        )
        .reset_index()
        .sort_values(by=["issue_boxes", "affected_packages"], ascending=False)
    )
    return route_summary

def build_summary_text(display_df):
    if display_df.empty:
        return "PFF现场领件检查结果：\n\n未发现现场到货箱号仍处于未领状态。"

    issue_boxes = display_df["box_no_clean"].nunique()
    affected_packages = pd.to_numeric(display_df["count"], errors="coerce").fillna(0).sum()

    return "\n".join([
        "PFF现场领件检查结果：",
        "",
        f"共发现 {issue_boxes} 个现场到货箱号仍处于未领状态，涉及 {int(affected_packages)} 件包裹。",
        "请站点经理和司机立即现场核查并完成补领。"
    ])

def build_excel_bytes(onsite_df, task_df, matched_df, display_df, route_summary, summary_text):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        onsite_df.to_excel(writer, sheet_name="onsite_boxes", index=False)
        task_df.to_excel(writer, sheet_name="task_unpicked_boxes", index=False)
        matched_df.to_excel(writer, sheet_name="matched_raw", index=False)
        display_df.to_excel(writer, sheet_name="issue_boxes", index=False)
        route_summary.to_excel(writer, sheet_name="route_summary", index=False)
        pd.DataFrame({"summary": summary_text.split("\n")}).to_excel(
            writer, sheet_name="summary", index=False
        )
    output.seek(0)
    return output.getvalue()

# =========================
# 页面
# =========================
st.title("PFF Pickup Check Tool")
st.markdown(
    '<div class="small-note">请上传 Gary 现场箱号表和任务表，系统将核查现场已到货但仍显示未领的箱号，PFF 仅检查 F3 / F4 / F5 / F11 / F12 / F13。</div>',
    unsafe_allow_html=True
)

col1, col2 = st.columns(2)

with col1:
    onsite_file = st.file_uploader(
        "Upload Gary onsite box file",
        type=["xlsx", "xls"],
        key="onsite_file"
    )

with col2:
    task_file = st.file_uploader(
        "Upload DMS unpicked task file",
        type=["xlsx", "xls"],
        key="task_file"
    )

st.markdown('<div class="run-btn">', unsafe_allow_html=True)
run_check = st.button("Run Pickup Check")
st.markdown('</div>', unsafe_allow_html=True)

if run_check:
    if onsite_file is None or task_file is None:
        st.error("请先上传两个文件。")
    else:
        try:
            onsite_df = load_onsite_file(onsite_file)
            task_df = load_task_file(task_file)

            matched_df, display_df = compare_files(onsite_df, task_df)
            route_summary = build_route_summary(display_df)
            summary_text = build_summary_text(display_df)

            onsite_boxes = onsite_df["box_no_clean"].nunique()
            unpicked_bags = task_df["box_no_clean"].nunique()
            issue_boxes = display_df["box_no_clean"].nunique()
            affected_packages = pd.to_numeric(display_df["count"], errors="coerce").fillna(0).sum()

            st.success("检查完成。")

            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.markdown('<div class="metric-label">Onsite Boxes</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="metric-value">{onsite_boxes}</div>', unsafe_allow_html=True)
            with k2:
                st.markdown('<div class="metric-label">Unpicked Bags</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="metric-value">{unpicked_bags}</div>', unsafe_allow_html=True)
            with k3:
                st.markdown('<div class="metric-label">Issue Boxes</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="metric-value">{issue_boxes}</div>', unsafe_allow_html=True)
            with k4:
                st.markdown('<div class="metric-label">Affected Packages</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="metric-value">{int(affected_packages)}</div>', unsafe_allow_html=True)

            st.markdown("### Issue Box List")
            if display_df.empty:
                st.info("No onsite boxes are still marked as unpicked.")
            else:
                cols_to_show = [
                    "box no", "route-stop", "route", "address", "count", "box_no_clean"
                ]
                cols_to_show = [c for c in cols_to_show if c in display_df.columns]
                st.dataframe(
                    display_df[cols_to_show],
                    use_container_width=True,
                    height=320
                )

            st.markdown("### Route Summary")
            st.dataframe(
                route_summary,
                use_container_width=True,
                height=240
            )

            st.markdown("### Summary")
            st.text_area(
                "Generated message",
                summary_text,
                height=110,
                label_visibility="collapsed"
            )

            excel_bytes = build_excel_bytes(
                onsite_df, task_df, matched_df, display_df, route_summary, summary_text
            )
            today_str = datetime.today().strftime("%Y%m%d")
            st.download_button(
                "Download Excel Report",
                data=excel_bytes,
                file_name=f"PFF_Pickup_Check_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"运行失败：{e}")
