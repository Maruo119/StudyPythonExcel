"""
Excel工数管理自動化スクリプト（修正版）

[目的]
設定Excelの工数データを読み込み、資産/費用に分けた報告用Excelを自動生成する。

[処理フロー]
1. config.xlsx から マスタデータ（アプリ、フェーズ）と工数（ベンダー・社員）を読み込み
2. フェーズの区分（資産/費用）に基づいて、工数データを分類
3. アプリ×ベンダー単位でグループ化
4. 報告_資産.xlsx と 報告_費用.xlsx の 2ファイルを自動生成（シート名は詳細_1～）
"""

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
from typing import Dict, List, Tuple


def read_config() -> Tuple[Dict, Dict, List]:
    """
    設定Excelを読み込み、マスタデータと工数を返す。

    Returns:
        apps: {アプリID: アプリ名}
        phases: {フェーズ名: 区分}  ※区分は "資産" or "費用"
        work_data: [(アプリID, ベンダー名, フェーズ名, ベンダー工数, 発注金額, 社員工数), ...]
    """
    config_path = Path("input") / "config.xlsx"

    if not config_path.exists():
        raise FileNotFoundError(f"設定ファイルが見つかりません: {config_path}")

    wb = load_workbook(config_path)

    # アプリマスタを読み込み
    apps = {}
    ws_apps = wb["アプリ"]
    for row in ws_apps.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        app_id, app_name = row[0], row[1]
        apps[app_id] = app_name

    # フェーズマスタを読み込み（新形式: フェーズ名で管理）
    phases = {}
    ws_phases = wb["フェーズ"]
    for row in ws_phases.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        phase_name, division = row[0], row[1]
        phases[phase_name] = division

    # 工数データを読み込み（ベンダー・社員統合）
    work_data = []
    ws_work = wb["工数"]
    for row in ws_work.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        app_id, vendor_name, phase_name, vendor_hours, amount, employee_hours = row[0:6]
        work_data.append((app_id, vendor_name, phase_name, vendor_hours, amount, employee_hours))

    wb.close()
    return apps, phases, work_data


def organize_data_by_division(work_data: List, phases: Dict) -> Tuple[Dict, Dict]:
    """
    工数データをフェーズの区分（資産/費用）に基づいて分類する。
    アプリ×ベンダー単位でグループ化。

    Args:
        work_data: [(アプリID, ベンダー名, フェーズ名, ベンダー工数, 発注金額, 社員工数), ...]
        phases: {フェーズ名: 区分}

    Returns:
        asset_data: {(アプリID, ベンダー名): {フェーズ名: {vendor_hours, amount, employee_hours}}}
        expense_data: {(アプリID, ベンダー名): {フェーズ名: {vendor_hours, amount, employee_hours}}}
    """
    asset_data = {}
    expense_data = {}

    for app_id, vendor_name, phase_name, vendor_hours, amount, employee_hours in work_data:
        if phase_name not in phases:
            print(f"[WARN] フェーズ '{phase_name}' がマスタに見つかりません")
            continue

        division = phases[phase_name]
        key = (app_id, vendor_name)

        # キーの初期化
        if key not in asset_data:
            asset_data[key] = {}
            expense_data[key] = {}

        # フェーズデータを辞書で保存
        phase_data = {
            "vendor_hours": vendor_hours,
            "amount": amount,
            "employee_hours": employee_hours
        }

        # 資産/費用に分類
        if division == "資産":
            asset_data[key][phase_name] = phase_data
        elif division == "費用":
            expense_data[key][phase_name] = phase_data
        else:
            print(f"[WARN] 不正な区分 '{division}' (フェーズ: {phase_name})")

    return asset_data, expense_data


def create_workbook(apps: Dict, data_by_app_vendor: Dict, file_name: str):
    """
    報告用Excelを生成する。

    Args:
        apps: {アプリID: アプリ名}
        data_by_app_vendor: {(アプリID, ベンダー名): {フェーズ名: {...}}}
        file_name: 出力ファイル名（例: "報告_資産.xlsx"）
    """
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)

    sheet_num = 1

    # アプリ×ベンダーごとにシートを作成
    for (app_id, vendor_name) in sorted(data_by_app_vendor.keys()):
        if not data_by_app_vendor[(app_id, vendor_name)]:
            continue

        app_name = apps.get(app_id, app_id)

        # シート名は固定名：詳細_1, 詳細_2, ...
        sheet_title = f"詳細_{sheet_num}"
        ws = wb.create_sheet(title=sheet_title)
        sheet_num += 1

        # 1行目：メモ欄
        ws["A1"] = app_name
        ws["B1"] = vendor_name
        ws["C1"] = "社員工数"

        # 社員工数の合計を計算して D1 に記入
        phases_data = data_by_app_vendor[(app_id, vendor_name)]
        total_employee_hours = sum(phase["employee_hours"] for phase in phases_data.values())
        ws["D1"] = total_employee_hours

        # 2行目：ヘッダー行
        ws["A2"] = "フェーズ名"
        ws["B2"] = "ベンダー工数（人日）"
        ws["C2"] = "発注金額"
        ws["D2"] = "社員工数（人日）"

        # 3行目以降：データ行
        row = 3
        for phase_name, phase_data in sorted(phases_data.items()):
            ws[f"A{row}"] = phase_name
            ws[f"B{row}"] = phase_data["vendor_hours"]
            ws[f"C{row}"] = phase_data["amount"]
            ws[f"D{row}"] = phase_data["employee_hours"]
            row += 1

        # セルの装飾
        format_workbook(ws)

    # ファイルを保存
    output_path = Path("output") / file_name
    wb.save(output_path)
    print(f"[OK] Created: {output_path}")


def format_workbook(ws):
    """
    ワークシートのセルをフォーマットする（背景色、フォント、枠線など）。
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

    # 1行目（メモ欄）のフォーマット
    for col in ["A", "B", "C", "D"]:
        cell = ws[f"{col}1"]
        if cell.value is not None:
            cell.fill = memo_fill
            cell.font = memo_font
            cell.border = thin_border
            cell.alignment = left_align

    # 2行目（ヘッダー行）のフォーマット
    for col in ["A", "B", "C", "D"]:
        cell = ws[f"{col}2"]
        if cell.value is not None:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = center_align

    # データ行のフォーマット
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=4):
        for cell in row:
            if cell.value is not None:
                cell.border = thin_border
                if cell.column in [2, 3, 4]:  # B, C, D 列は右揃え（数値）
                    cell.alignment = right_align
                else:  # A 列は左揃え（テキスト）
                    cell.alignment = left_align

    # 列幅を調整
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 18


def main():
    """メイン処理"""
    try:
        print("=" * 50)
        print("Excel工数管理自動化を開始します")
        print("=" * 50)

        # 1. 設定Excelを読み込み
        print("\n[Step 1] Reading configuration file...")
        apps, phases, work_data = read_config()
        print(f"  [OK] Number of apps: {len(apps)}")
        print(f"  [OK] Number of phases: {len(phases)}")
        print(f"  [OK] Number of work records: {len(work_data)}")

        # 2. 工数データを資産/費用に分類
        print("\n[Step 2] Classifying work data...")
        asset_data, expense_data = organize_data_by_division(work_data, phases)
        print(f"  [OK] Asset app-vendor pairs: {len(asset_data)}")
        print(f"  [OK] Expense app-vendor pairs: {len(expense_data)}")

        # 3. 報告用Excelを生成
        print("\n[Step 3] Generating report files...")
        create_workbook(apps, asset_data, "houkoku_shisan.xlsx")
        create_workbook(apps, expense_data, "houkoku_hiyo.xlsx")

        print("\n" + "=" * 50)
        print("[OK] Process completed successfully!")
        print("=" * 50)

    except FileNotFoundError as e:
        print(f"[ERROR] {e}")
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
