"""
テンプレートExcelファイル（houkoku_template.xlsx）を作成するスクリプト
15シート用意（最初から全シートを用意しておき、使うものだけ記入）
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from pathlib import Path


def create_sheet_template(ws):
    """
    シートのテンプレート格式を設定する
    """
    # スタイル定義
    memo_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    memo_font = Font(bold=False, size=10)

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)

    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # 数値のカンマ区切り書式
    comma_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1  # #,##0

    # 1行目：メモ欄
    ws["A1"] = "[メモ]"
    for col in ["A", "B", "C", "D", "E"]:
        cell = ws[f"{col}1"]
        cell.fill = memo_fill
        cell.font = memo_font
        cell.border = thin_border
        cell.alignment = left_align

    # 2行目：ヘッダー行
    headers = ["アプリ名称", "フェーズ", "ベンダー/社員", "工数（人日）", "金額"]
    for col_idx, header in enumerate(headers, 1):
        col_letter = chr(64 + col_idx)
        cell = ws[f"{col_letter}2"]
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = center_align

    # 3行目以降：データエリア（50行分のダミー行を作成して枠線を用意）
    for row in range(3, 53):
        for col in range(1, 6):
            col_letter = chr(64 + col)
            cell = ws[f"{col_letter}{row}"]
            cell.border = thin_border
            if col == 4:  # D 列：工数（カンマ区切り）
                cell.alignment = right_align
                cell.number_format = comma_format
            elif col == 5:  # E 列：金額（カンマ区切り）
                cell.alignment = right_align
                cell.number_format = comma_format
            else:
                cell.alignment = left_align

    # 列幅を調整
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15


def create_template():
    """報告用Excelテンプレートを作成する（15シート用意）"""
    template_dir = Path("template")
    template_dir.mkdir(exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)

    # 15シート作成
    for sheet_num in range(1, 16):
        sheet_name = f"詳細_{sheet_num}"
        ws = wb.create_sheet(title=sheet_name)
        create_sheet_template(ws)

    # ファイルを保存
    template_path = template_dir / "houkoku_template.xlsx"
    wb.save(template_path)
    print(f"[OK] Template created: {template_path} (15 sheets)")


if __name__ == "__main__":
    create_template()
