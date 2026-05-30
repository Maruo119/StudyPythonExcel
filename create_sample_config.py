"""
サンプル設定Excelを生成するスクリプト（修正版）

新しい構成：
- シート1: 「アプリ」（アプリID, アプリ名）
- シート2: 「フェーズ」（フェーズ名, 区分）
- シート3: 「工数」（アプリID, ベンダー名, フェーズ名, ベンダー工数, 発注金額, 社員工数）
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
# シート3: 工数データ（ベンダー・社員統合）
# ============================================
ws_work = wb.create_sheet("工数")
ws_work["A1"] = "アプリID"
ws_work["B1"] = "ベンダー名"
ws_work["C1"] = "フェーズ名"
ws_work["D1"] = "ベンダー工数（人日）"
ws_work["E1"] = "発注金額"
ws_work["F1"] = "社員工数（人日）"

work_data = [
    ("APP001", "ベンダーA", "P001_要件定義", 5, 150000, 2),
    ("APP001", "ベンダーA", "P002_基本設計", 8, 240000, 3),
    ("APP001", "ベンダーB", "P004_実装", 15, 450000, 1),
    ("APP001", "ベンダーB", "P005_単体テスト", 10, 300000, 2),

    ("APP002", "ベンダーC", "P001_要件定義", 10, 300000, 4),
    ("APP002", "ベンダーC", "P002_基本設計", 12, 360000, 5),
    ("APP002", "ベンダーC", "P003_詳細設計", 20, 600000, 3),
    ("APP002", "ベンダーC", "P004_実装", 30, 900000, 6),
    ("APP002", "ベンダーC", "P006_統合テスト", 18, 540000, 4),

    ("APP003", "ベンダーA", "P002_基本設計", 6, 180000, 2),
    ("APP003", "ベンダーA", "P003_詳細設計", 10, 300000, 3),
    ("APP003", "ベンダーD", "P004_実装", 20, 600000, 4),

    ("APP004", "ベンダーE", "P001_要件定義", 3, 90000, 1),
    ("APP004", "ベンダーE", "P002_基本設計", 5, 150000, 2),
    ("APP004", "ベンダーE", "P004_実装", 12, 360000, 2),
    ("APP004", "ベンダーE", "P005_単体テスト", 8, 240000, 1),

    ("APP005", "ベンダーA", "P001_要件定義", 4, 120000, 1),
    ("APP005", "ベンダーF", "P003_詳細設計", 15, 450000, 3),
    ("APP005", "ベンダーF", "P004_実装", 25, 750000, 5),

    ("APP006", "ベンダーC", "P002_基本設計", 8, 240000, 2),
    ("APP006", "ベンダーC", "P004_実装", 18, 540000, 3),
    ("APP006", "ベンダーC", "P005_単体テスト", 10, 300000, 2),

    ("APP007", "ベンダーE", "P001_要件定義", 3, 90000, 1),
    ("APP007", "ベンダーE", "P003_詳細設計", 12, 360000, 2),
    ("APP007", "ベンダーD", "P004_実装", 16, 480000, 3),

    ("APP008", "ベンダーA", "P002_基本設計", 7, 210000, 2),
    ("APP008", "ベンダーB", "P004_実装", 14, 420000, 2),

    ("APP009", "ベンダーD", "P003_詳細設計", 11, 330000, 2),
    ("APP009", "ベンダーD", "P004_実装", 22, 660000, 4),

    ("APP010", "ベンダーE", "P001_要件定義", 6, 180000, 2),
    ("APP010", "ベンダーE", "P002_基本設計", 9, 270000, 3),
    ("APP010", "ベンダーC", "P004_実装", 20, 600000, 4),
    ("APP010", "ベンダーC", "P006_統合テスト", 12, 360000, 2),
]

for idx, (app_id, vendor, phase, vendor_hours, amount, emp_hours) in enumerate(work_data, start=2):
    ws_work[f"A{idx}"] = app_id
    ws_work[f"B{idx}"] = vendor
    ws_work[f"C{idx}"] = phase
    ws_work[f"D{idx}"] = vendor_hours
    ws_work[f"E{idx}"] = amount
    ws_work[f"F{idx}"] = emp_hours

ws_work.column_dimensions["A"].width = 12
ws_work.column_dimensions["B"].width = 12
ws_work.column_dimensions["C"].width = 25
ws_work.column_dimensions["D"].width = 18
ws_work.column_dimensions["E"].width = 12
ws_work.column_dimensions["F"].width = 18

# ============================================
# ヘッダー行のスタイル統一
# ============================================
header_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
header_font = Font(bold=True, size=11)
header_align = Alignment(horizontal="center", vertical="center")

for ws in [ws_apps, ws_phases, ws_work]:
    for cell in ws[1]:
        if cell.value:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align

# ファイルを保存
Path("input").mkdir(exist_ok=True)
wb.save("input\\config.xlsx")
print("[OK] Sample configuration file created: input\\config.xlsx")
