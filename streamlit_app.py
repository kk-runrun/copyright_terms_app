import io
import re

import pandas as pd
import streamlit as st
import xlsxwriter


st.set_page_config(page_title="查询文案侵权词", layout="wide")
st.title("查询文案侵权词")
st.caption("上传 Excel，输入关键词（逗号分隔），生成标红结果报告")


def run_check(uploaded_file, keywords, header_row_num: int) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, header=header_row_num - 1)

    # 为了兼容原脚本逻辑，读取第2行作为父级列名（处理合并单元格）
    uploaded_file.seek(0)
    parent_cols = (
        pd.read_excel(uploaded_file, header=None, skiprows=1, nrows=1)
        .iloc[0]
        .ffill()
        .fillna("")
        .tolist()
    )

    search_pattern = "|".join(re.escape(word) for word in keywords)
    results = []

    sku_col_name = df.columns[0]
    site_col_name = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    lang_col_name = df.columns[2] if len(df.columns) > 2 else df.columns[0]

    for _, row in df.iterrows():
        sku = str(row[sku_col_name]) if pd.notnull(row[sku_col_name]) else "未知SKU"
        site = str(row[site_col_name]) if pd.notnull(row[site_col_name]) else "未知站点"
        lang = str(row[lang_col_name]) if pd.notnull(row[lang_col_name]) else "未知语种"

        for i in range(3, len(df.columns)):
            col_name = df.columns[i]
            cell_content = str(row[col_name]) if pd.notnull(row[col_name]) else ""
            if not cell_content:
                continue
            if re.search(search_pattern, cell_content, re.IGNORECASE):
                clean_col_name = str(col_name).split(".")[0]
                parent = str(parent_cols[i]) if i < len(parent_cols) else ""
                results.append(
                    {
                        "SKU": sku,
                        "站点": site,
                        "语种": lang,
                        "位置": f"{parent} {clean_col_name}".strip(),
                        "语句": cell_content,
                    }
                )
    return pd.DataFrame(results, columns=["SKU", "站点", "语种", "位置", "语句"])


def build_result_excel(result_df: pd.DataFrame, keywords) -> bytes:
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    worksheet = workbook.add_worksheet("查询结果")

    header_format = workbook.add_format({"bold": True, "bg_color": "#CFE2F3", "border": 1})
    red_format = workbook.add_format({"font_color": "red"})
    cell_format = workbook.add_format({"text_wrap": True, "valign": "vcenter", "border": 1})

    headers = ["SKU", "站点", "语种", "具体位置", "包含关键词的语句"]
    for col, text in enumerate(headers):
        worksheet.write(0, col, text, header_format)

    worksheet.set_column("A:A", 24)
    worksheet.set_column("B:C", 10)
    worksheet.set_column("D:D", 28)
    worksheet.set_column("E:E", 60)

    search_pattern = "|".join(re.escape(k) for k in keywords)
    for row_idx, item in enumerate(result_df.to_dict("records"), start=1):
        worksheet.write(row_idx, 0, item["SKU"], cell_format)
        worksheet.write(row_idx, 1, item["站点"], cell_format)
        worksheet.write(row_idx, 2, item["语种"], cell_format)
        worksheet.write(row_idx, 3, item["位置"], cell_format)
        text = item["语句"]
        parts = re.split(f"({search_pattern})", text, flags=re.IGNORECASE)
        rich_segments = []
        for part in parts:
            if not part:
                continue
            if any(part.lower() == k.lower() for k in keywords):
                rich_segments.extend([red_format, part])
            else:
                rich_segments.append(part)
        if len(rich_segments) > 1:
            try:
                worksheet.write_rich_string(row_idx, 4, *rich_segments, cell_format)
            except Exception:
                worksheet.write(row_idx, 4, text, cell_format)
        else:
            worksheet.write(row_idx, 4, text, cell_format)

    workbook.close()
    return output.getvalue()


uploaded = st.file_uploader("上传待检查 Excel", type=["xlsx", "xls"])
keywords_raw = st.text_area("关键词（逗号分隔）", value="Ratcheting cargo bar")
header_row_num = st.number_input("表头行号（从 1 开始）", min_value=1, max_value=20, value=3, step=1)

if st.button("开始检查", type="primary", disabled=uploaded is None):
    keywords = [k.strip() for k in keywords_raw.split(",") if k.strip()]
    if not keywords:
        st.error("请至少输入一个关键词")
    else:
        with st.spinner("检查中..."):
            try:
                result_df = run_check(uploaded, keywords, int(header_row_num))
            except Exception as exc:
                st.error(f"处理失败: {exc}")
                st.stop()

        if result_df.empty:
            st.warning("未发现匹配关键词。")
        else:
            st.success(f"完成，共发现 {len(result_df)} 处匹配")
            st.dataframe(result_df, use_container_width=True, height=420)

        excel_bytes = build_result_excel(result_df, keywords)
        st.download_button(
            "下载结果 Excel",
            data=excel_bytes,
            file_name="文案侵权词查询结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
