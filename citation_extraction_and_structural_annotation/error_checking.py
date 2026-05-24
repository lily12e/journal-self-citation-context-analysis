import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# ============================
# 异常Self-cited Article检测：检查多个年份
# ============================
def highlight_multiple_years_in_self_cited_article(self_cited_article: str) -> bool:
    """
    检测 Self-cited Article 是否匹配到多个年份，如果有，则返回 True
    """
    if not isinstance(self_cited_article, str):
        return False

    # 使用非捕获分组，findall 返回完整年份，如 2021 或 2019a
    years_found = re.findall(r'\b(?:19|20)\d{2}[a-z]?\b', self_cited_article)
    return len(set(years_found)) > 1


def highlight_abnormal_length_of_citation(content: str, min_len: int = 350) -> str:
    content = clean_text(content)
    print(f"Cleaned content length: {len(content)}")
    if len(content) < min_len:
        return "red"   # 过短标红
    return "none"


# ============================
# 解决异常字符问题
# ============================
def highlight_excessive_dashes(content: str) -> bool:
    """
    检测并标蓝过多的破折号（'-'、'–'、'—'等）
    """
    if not isinstance(content, str):
        return False
    dash_count = content.count('-') + content.count('–') + content.count('—')
    return dash_count > 10


# ============================
# 清洗 Citation Content
# ============================
def clean_text(content: str) -> str:
    """清洗文本中的常见异常字符和格式问题"""
    if not content or isinstance(content, float):
        return ""
    # 去除多余空白
    content = re.sub(r'\s+', ' ', content)
    # 去除连续破折号
    content = re.sub(r'[-\u2013\u2014]{2,}', ' ', content)
    return content.strip()


# ============================
# 工具函数：智能识别列名
# ============================
def detect_col_name(df: pd.DataFrame, candidates) -> str:
    """
    在 candidates 中按顺序寻找第一个存在于 df.columns 的列名（大小写精确匹配）。
    若均不存在，再做一次大小写不敏感匹配。
    """
    # 1) 直接匹配
    for c in candidates:
        if c in df.columns:
            return c
    # 2) 大小写不敏感匹配
    lower_map = {col.lower(): col for col in df.columns}
    for c in candidates:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    raise KeyError(f"未在表头中找到任一候选列: {candidates}. 实际列名为: {list(df.columns)}")


# ============================
# 清洗和标色
# ============================
def clean_and_highlight_citation_content(df: pd.DataFrame,
                                         citation_column: str,
                                         self_cited_column: str,
                                         ws) -> pd.DataFrame:
    """
    清洗 Citation Content，并标红、标蓝异常内容
    """
    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")   # 淡红
    blue_fill = PatternFill(start_color="99CCFF", end_color="99CCFF", fill_type="solid")  # 淡蓝

    cleaned_col = 'Citation Content (±3)_cleaned'
    if cleaned_col not in df.columns:
        df[cleaned_col] = df[citation_column].apply(clean_text)

    # 确保在工作表中也有新列的表头（Excel 第1行）
    cleaned_col_excel_idx = df.columns.get_loc(cleaned_col) + 1  # 1-based
    if ws.cell(row=1, column=cleaned_col_excel_idx).value != cleaned_col:
        ws.cell(row=1, column=cleaned_col_excel_idx, value=cleaned_col)

    for i, row in df.iterrows():
        citation_content = row.get(citation_column, "")
        self_cited_article = row.get(self_cited_column, "")

        # 多年份 -> 标蓝（Self-cited 列）
        if highlight_multiple_years_in_self_cited_article(self_cited_article):
            ws.cell(row=i + 2, column=df.columns.get_loc(self_cited_column) + 1).fill = blue_fill

        # 过短 -> 标红（Citation 列）
        length_status = highlight_abnormal_length_of_citation(citation_content)
        if length_status == "red":
            ws.cell(row=i + 2, column=df.columns.get_loc(citation_column) + 1).fill = red_fill
        # 破折号过多 -> 标蓝（Citation 列）
        elif highlight_excessive_dashes(citation_content):
            ws.cell(row=i + 2, column=df.columns.get_loc(citation_column) + 1).fill = blue_fill

        # 写入清洗后的文本
        ws.cell(row=i + 2, column=cleaned_col_excel_idx).value = row[cleaned_col]

    return df


def main(input_xlsx: str,
         output_xlsx: str,
         sheet_name: int = 0,
         citation_column_candidates=None,
         self_cited_column_candidates=None) -> None:
    if citation_column_candidates is None:
        # 兼容有/无空格两种
        citation_column_candidates = ["Citation Content(±3)", "Citation Content (±3)"]

    if self_cited_column_candidates is None:
        # 兼容两种写法 + 常见首字母大写变体
        self_cited_column_candidates = [
            "Self-cited Article",
            "Self-cited Article title",
            "Self-cited Article Title"
        ]

    # 读数据
    df = pd.read_excel(input_xlsx, sheet_name=sheet_name, engine="openpyxl")

    # 自动识别两类列名
    citation_column = detect_col_name(df, citation_column_candidates)
    self_cited_column = detect_col_name(df, self_cited_column_candidates)
    print(f"使用列：citation_column = '{citation_column}', self_cited_column = '{self_cited_column}'")

    # 打开工作簿与工作表
    wb = load_workbook(input_xlsx)
    sheet_names = wb.sheetnames
    print(f"工作簿中的工作表名称: {sheet_names}")
    ws = wb[sheet_names[sheet_name]]

    # 清洗 + 高亮
    df = clean_and_highlight_citation_content(df, citation_column, self_cited_column, ws)

    # 保存到新文件
    wb.save(output_xlsx)
    print(f"✅ 已完成处理，输出：{output_xlsx}")


if __name__ == "__main__":
    input_xlsx = "./input/your_annotation_file.xlsx"
    output_xlsx = "./output/your_annotation_file_checked.xlsx"
    main(input_xlsx, output_xlsx)
