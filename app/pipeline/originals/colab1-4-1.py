from __future__ import annotations

import os
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl import load_workbook


DEFAULT_FILE_PATH = "/content/CF付財務分析表（経営指標あり）_ReadingData_updated.xlsx"
DEFAULT_SHEET_NAME = "CF計算書"
DEFAULT_TITLE = "キャッシュ・フロー計算書"


def read_excel_as_dataframe(file_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Excelをdata_only=Trueで読み込み、数式は計算済み値を取得。
    """
    wb = load_workbook(file_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"シートが見つかりません: {sheet_name}")

    ws = wb[sheet_name]
    data = ws.values
    df = pd.DataFrame(data)
    return df


def build_html(file_path: str,
               sheet_name: str = DEFAULT_SHEET_NAME,
               title: str = DEFAULT_TITLE) -> str:
    """
    Colab版ロジックをそのまま使用しつつ、
    サーバ実行用HTML出力に対応した完全互換版。
    """

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Excelファイルが見つかりません: {file_path}")

    df_raw = read_excel_as_dataframe(file_path, sheet_name)

    # ヘッダー情報
    header_info_period = str(df_raw.iloc[4, 1]) if not pd.isna(df_raw.iloc[4, 1]) else ""
    header_info_unit = str(df_raw.iloc[5, 2]) if not pd.isna(df_raw.iloc[5, 2]) else ""

    # B列・C列抽出
    df = df_raw.iloc[6:51, [1, 2]].fillna('')

    custom_css = """
    <style>
        .report-container { font-family: "Meiryo", sans-serif; color: #000; }
        .report-title { font-size: 18px; font-weight: bold; margin: 20px 0 5px 0; border-left: 5px solid #000; padding-left: 10px; }
        .report-meta { font-size: 13px; margin-bottom: 5px; display: flex; justify-content: space-between; width: 580px; }
        .excel-table {
            border-collapse: collapse; width: fit-content; display: inline-table; font-size: 13px;
            table-layout: fixed; outline: 2px solid #000; box-shadow: 0 0 0 2px #000; background-color: white;
        }
        .excel-table td {
            border-left: 1px solid #999; border-bottom: 1px solid #999;
            padding: 4px 10px; overflow: hidden; white-space: nowrap; box-sizing: border-box; color: #000;
        }
        .excel-table tr td:last-child { border-right: 1px solid #999; }
        .col-subject { width: 400px; text-align: left; }
        .col-amt { width: 180px; text-align: right; font-family: "Consolas", monospace; }
        .header-row td { background-color: #004080 !important; color: #ffffff !important; font-weight: bold; }
        .total-row { background-color: #e6f3ff !important; font-weight: bold; }
        .grand-total { background-color: #d9ead3 !important; font-weight: bold; }
    </style>
    """

    html_output = '<div class="report-container">'
    html_output += f'<div class="report-title">{title}</div>'
    html_output += f'<div class="report-meta"><span>{header_info_period}</span><span>{header_info_unit}</span></div>'
    html_output += '<table class="excel-table"><tbody>'

    for i, row in df.iterrows():
        subject = str(row.iloc[0]).strip()
        amt = row.iloc[1]

        if not subject and (amt == "" or amt is None):
            continue

        row_class = ""
        if "キャッシュ・フロー" in subject and i < 20:
            row_class = "header-row"
        elif "計" in subject:
            row_class = "total-row"
        elif "現金及び現金同等物" in subject:
            row_class = "grand-total"

        display_amt = amt
        if isinstance(amt, (int, float)) and amt != "":
            try:
                display_amt = "{:,}".format(int(amt))
            except Exception:
                pass

        html_output += f'<tr class="{row_class}"><td class="col-subject">{row.iloc[0]}</td>'
        html_output += f'<td class="col-amt">{display_amt}</td></tr>'

    html_output += "</tbody></table></div>"

    return (
        '<!doctype html><html><head><meta charset="utf-8">'
        + custom_css +
        "</head><body>" +
        html_output +
        "</body></html>"
    )


def main() -> None:
    """
    実行方法：

    1) Colab:
       python colab1-4-1.py
       → HTMLをstdout出力

    2) サーバ:
       INPUT_XLSX=xxx.xlsx OUTPUT_HTML=out.html python colab1-4-1.py
    """

    file_path = os.environ.get("INPUT_XLSX", DEFAULT_FILE_PATH)
    sheet = os.environ.get("SHEET_NAME", DEFAULT_SHEET_NAME)
    title = os.environ.get("TITLE", DEFAULT_TITLE)
    out_html = os.environ.get("OUTPUT_HTML", "")

    html = build_html(file_path, sheet, title)

    if out_html:
        Path(out_html).write_text(html, encoding="utf-8")
        print(out_html)
    else:
        print(html)


if __name__ == "__main__":
    main()