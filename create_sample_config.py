"""
サンプル設定Excelを生成するスクリプト（新仕様版）

新しい構成：
- シート1: 「アプリ」（アプリID, アプリ名）
- シート2: 「フェーズ」（フェーズ名, 区分）
- シート3: 「社員単価」（単価ラベル, 単価）
- シート4: 「工数」（アプリID, ベンダー/社員, フェーズ名, 工数, 発注金額, -, 社員コスト, 出力シート指定）
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path

wb = Workbook()
wb.remove(wb.active)

# ============================================
# シート1: アプリマスタ
# ============================================
ws_apps = wb.create_sheet("アプリ")
ws_apps["A1"] = "アプリID"
ws_apps["B1"] = "アプリ名"

apps_data = [
    ("APP001", "ユーザー管理システム"),
    ("APP002", "決済システム"),
    ("APP003", "レポート生成システム"),
    ("APP004", "通知サービス"),
    ("APP005", "認証サービス"),
    ("APP006", "ログ管理システム"),
    ("APP007", "バックアップサービス"),
    ("APP008", "パフォーマンス監視"),
    ("APP009", "キャッシュシステム"),
    ("APP010", "データ同期エンジン"),
]

for idx, (app_id, app_name) in enumerate(apps_data, start=2):
    ws_apps[f"A{idx}"] = app_id
    ws_apps[f"B{idx}"] = app_name

ws_apps.column_dimensions["A"].width = 12
ws_apps.column_dimensions["B"].width = 20

# ============================================
# シート2: フェーズマスタ（新形式）
# ============================================
ws_phases = wb.create_sheet("フェーズ")
ws_phases["A1"] = "フェーズ名"
ws_phases["B1"] = "区分"

phases_data = [
    ("P001_要件定義", "資産"),
    ("P002_基本設計", "資産"),
    ("P003_詳細設計", "資産"),
    ("P004_実装", "費用"),
    ("P005_単体テスト", "費用"),
    ("P006_統合テスト", "費用"),
    ("P007_本番環境構築", "費用"),
]

for idx, (phase_name, division) in enumerate(phases_data, start=2):
    ws_phases[f"A{idx}"] = phase_name
    ws_phases[f"B{idx}"] = division

ws_phases.column_dimensions["A"].width = 25
ws_phases.column_dimensions["B"].width = 10

# ============================================
# シート3: 社員単価マスタ（新規）
# ============================================
ws_employee_cost = wb.create_sheet("社員単価")
ws_employee_cost["A1"] = "単価ラベル"
ws_employee_cost["B1"] = "単価"
ws_employee_cost["A2"] = "社員単価"
ws_employee_cost["B2"] = 100000  # ¥100,000/人日

ws_employee_cost.column_dimensions["A"].width = 15
ws_employee_cost.column_dimensions["B"].width = 15

# ============================================
# シート4: 工数データ（新仕様版）
# ============================================
ws_work = wb.create_sheet("工数")
ws_work["A1"] = "アプリID"
ws_work["B1"] = "ベンダー/社員"
ws_work["C1"] = "フェーズ名"
ws_work["D1"] = "工数（人日）"
ws_work["E1"] = "発注金額"
ws_work["F1"] = "社員コスト"
ws_work["G1"] = "出力シート指定"

# 新形式: (app_id, vendor_employee, phase, hours, vendor_amount, employee_cost, output_sheet)
# - vendor_employee: ベンダー名 または 「社員」
# - hours: 工数（人日）
# - vendor_amount: ベンダーの場合は金額、社員の場合は空
# - employee_cost: 社員の場合はコスト、ベンダーの場合は空
# - output_sheet: 出力先シート指定
work_data = [
    ("APP001", "ベンダーA", "P001_要件定義", 5, 150000, None, "詳細_1"),
    ("APP001", "ベンダーA", "P002_基本設計", 8, 240000, None, "詳細_1"),
    ("APP001", "社員", "P001_要件定義", 2, None, 200000, "詳細_1"),
    ("APP001", "ベンダーB", "P004_実装", 15, 450000, None, "詳細_2"),
    ("APP001", "社員", "P004_実装", 3, None, 300000, "詳細_2"),

    ("APP002", "ベンダーC", "P001_要件定義", 10, 300000, None, "詳細_3"),
    ("APP002", "ベンダーC", "P002_基本設計", 12, 360000, None, "詳細_3"),
    ("APP002", "社員", "P002_基本設計", 5, None, 500000, "詳細_3"),
    ("APP002", "ベンダーC", "P003_詳細設計", 20, 600000, None, "詳細_3"),

    ("APP003", "ベンダーA", "P002_基本設計", 6, 180000, None, "詳細_1"),
    ("APP003", "社員", "P002_基本設計", 3, None, 300000, "詳細_1"),
    ("APP003", "ベンダーD", "P004_実装", 20, 600000, None, "詳細_4"),

    ("APP004", "ベンダーE", "P001_要件定義", 3, 90000, None, "詳細_5"),
    ("APP004", "社員", "P001_要件定義", 1, None, 100000, "詳細_5"),
    ("APP004", "ベンダーE", "P004_実装", 12, 360000, None, "詳細_5"),

    ("APP005", "ベンダーA", "P001_要件定義", 4, 120000, None, "詳細_1"),
    ("APP005", "社員", "P001_要件定義", 2, None, 200000, "詳細_1"),
    ("APP005", "ベンダーF", "P004_実装", 25, 750000, None, "詳細_6"),

    ("APP006", "ベンダーC", "P002_基本設計", 8, 240000, None, "詳細_3"),
    ("APP006", "社員", "P002_基本設計", 2, None, 200000, "詳細_3"),

    ("APP007", "ベンダーE", "P001_要件定義", 3, 90000, None, "詳細_5"),
    ("APP007", "ベンダーD", "P004_実装", 16, 480000, None, "詳細_4"),
    ("APP007", "社員", "P004_実装", 2, None, 200000, "詳細_4"),

    ("APP008", "ベンダーA", "P002_基本設計", 7, 210000, None, "詳細_1"),
    ("APP008", "ベンダーB", "P004_実装", 14, 420000, None, "詳細_2"),
    ("APP008", "社員", "P004_実装", 2, None, 200000, "詳細_2"),

    ("APP009", "ベンダーD", "P003_詳細設計", 11, 330000, None, "詳細_4"),
    ("APP009", "社員", "P003_詳細設計", 3, None, 300000, "詳細_4"),

    ("APP010", "ベンダーE", "P001_要件定義", 6, 180000, None, "詳細_5"),
    ("APP010", "社員", "P001_要件定義", 2, None, 200000, "詳細_5"),
    ("APP010", "ベンダーC", "P004_実装", 20, 600000, None, "詳細_7"),
]

for idx, (app_id, vendor_emp, phase, hours, vendor_amount, emp_cost, output_sheet) in enumerate(work_data, start=2):
    ws_work[f"A{idx}"] = app_id
    ws_work[f"B{idx}"] = vendor_emp
    ws_work[f"C{idx}"] = phase
    ws_work[f"D{idx}"] = hours
    ws_work[f"E{idx}"] = vendor_amount
    ws_work[f"F{idx}"] = emp_cost
    ws_work[f"G{idx}"] = output_sheet

ws_work.column_dimensions["A"].width = 12
ws_work.column_dimensions["B"].width = 15
ws_work.column_dimensions["C"].width = 25
ws_work.column_dimensions["D"].width = 15
ws_work.column_dimensions["E"].width = 15
ws_work.column_dimensions["F"].width = 15
ws_work.column_dimensions["G"].width = 20

# ============================================
# ヘッダー行のスタイル統一
# ============================================
header_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
header_font = Font(bold=True, size=11)
header_align = Alignment(horizontal="center", vertical="center")

for ws in [ws_apps, ws_phases, ws_employee_cost, ws_work]:
    for cell in ws[1]:
        if cell.value:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align

# ファイルを保存
Path("input").mkdir(exist_ok=True)
wb.save("input\\config.xlsx")
print("[OK] Sample configuration file created: input\\config.xlsx")
