"""
Excel工数管理自動化スクリプト（新仕様版）

[目的]
設定Excelの工数データを読み込み、資産/費用に分けた報告用Excelを自動生成する。
さらに、突き合わせチェック用Excelを新規生成する。

[処理フロー]
1. config.xlsx から マスタデータ（アプリ、フェーズ、社員単価）と工数を読み込み
2. フェーズの区分（資産/費用）に基づいて、工数データを分類
3. アプリ×ベンダー/社員×出力シート指定 で グループ化
4. 報告_資産.xlsx と 報告_費用.xlsx の 2ファイルを自動生成
5. 突き合わせチェック用Excelを新規生成
"""

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
from typing import Dict, List, Tuple
from datetime import datetime


def read_config() -> Tuple[Dict, Dict, float, List]:
    """
    設定Excelを読み込み、マスタデータと工数を返す。

    Returns:
        apps: {アプリID: アプリ名}
        phases: {フェーズ名: 区分}  ※区分は "資産" or "費用"
        employee_cost_per_day: 社員単価（¥/人日）
        work_data: [(アプリID, vendor_or_employee, フェーズ名, hours, vendor_amount, employee_cost, output_sheet), ...]
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

    # フェーズマスタを読み込み
    phases = {}
    ws_phases = wb["フェーズ"]
    for row in ws_phases.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        phase_name, division = row[0], row[1]
        phases[phase_name] = division

    # 社員単価を読み込み
    employee_cost_per_day = 100000  # デフォルト値
    if "社員単価" in wb.sheetnames:
        ws_employee_cost = wb["社員単価"]
        # A2に社員単価, B2に金額が入っている
        cost_value = ws_employee_cost["B2"].value
        if cost_value is not None:
            employee_cost_per_day = float(cost_value)

    # 工数データを読み込み（新形式: 8列対応）
    work_data = []
    ws_work = wb["工数"]
    for row in ws_work.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        # A:アプリID, B:ベンダー/社員, C:フェーズ名, D:工数, E:発注金額, F:空, G:社員コスト, H:出力シート指定
        app_id = row[0]
        vendor_employee = row[1]  # ベンダー名 または 「社員」
        phase_name = row[2]
        hours = row[3]
        vendor_amount = row[4]  # ベンダーの場合は金額、社員の場合はNone
        employee_cost = row[6]  # 社員の場合はコスト、ベンダーの場合はNone
        output_sheet = row[7]  # 出力シート指定

        work_data.append((app_id, vendor_employee, phase_name, hours, vendor_amount, employee_cost, output_sheet))

    wb.close()
    return apps, phases, employee_cost_per_day, work_data


def organize_data_by_division(work_data: List, phases: Dict, employee_cost_per_day: float) -> Tuple[Dict, Dict]:
    """
    工数データをフェーズの区分（資産/費用）に基づいて分類する。
    アプリ×ベンダー/社員×出力シート指定 でグループ化。

    Args:
        work_data: [(アプリID, vendor_or_employee, フェーズ名, hours, vendor_amount, employee_cost, output_sheet), ...]
        phases: {フェーズ名: 区分}
        employee_cost_per_day: 社員単価（¥/人日）

    Returns:
        asset_data: {(アプリID, vendor_employee, output_sheet): {フェーズ名: {hours, amount}}}
        expense_data: {(アプリID, vendor_employee, output_sheet): {フェーズ名: {hours, amount}}}
    """
    asset_data = {}
    expense_data = {}

    for app_id, vendor_employee, phase_name, hours, vendor_amount, employee_cost, output_sheet in work_data:
        if phase_name not in phases:
            print(f"[WARN] フェーズ '{phase_name}' がマスタに見つかりません")
            continue

        division = phases[phase_name]
        key = (app_id, vendor_employee, output_sheet)

        # キーの初期化
        if key not in asset_data:
            asset_data[key] = {}
            expense_data[key] = {}

        # 金額を計算：社員の場合は hours × 単価、ベンダーの場合は vendor_amount
        if vendor_employee == "社員":
            amount = hours * employee_cost_per_day if hours else 0
        else:
            amount = vendor_amount if vendor_amount else 0

        # フェーズデータを辞書で保存
        phase_data = {
            "hours": hours,
            "amount": amount
        }

        # 資産/費用に分類
        if division == "資産":
            asset_data[key][phase_name] = phase_data
        elif division == "費用":
            expense_data[key][phase_name] = phase_data
        else:
            print(f"[WARN] 不正な区分 '{division}' (フェーズ: {phase_name})")

    return asset_data, expense_data


def create_workbook(apps: Dict, data_by_group: Dict, file_name: str):
    """
    報告用Excelを生成する。

    Args:
        apps: {アプリID: アプリ名}
        data_by_group: {(アプリID, vendor_employee, output_sheet): {フェーズ名: {hours, amount}}}
        file_name: 出力ファイル名（例: "houkoku_shisan.xlsx"）
    """
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)

    # 出力シート指定ごとにシートを作成
    sheet_dict = {}  # {sheet_name: 作成済みシートのリスト}

    for (app_id, vendor_employee, output_sheet_name) in sorted(data_by_group.keys()):
        if not data_by_group[(app_id, vendor_employee, output_sheet_name)]:
            continue

        app_name = apps.get(app_id, app_id)
        phases_data = data_by_group[(app_id, vendor_employee, output_sheet_name)]

        # シート名を決定（H列の指定値、またはデフォルト）
        if output_sheet_name and str(output_sheet_name).strip():
            sheet_title = str(output_sheet_name)
        else:
            sheet_title = f"詳細_{len(sheet_dict) + 1}"

        # 既存シートかどうかを確認
        if sheet_title not in sheet_dict:
            ws = wb.create_sheet(title=sheet_title)
            sheet_dict[sheet_title] = ws
            # 初回：ヘッダー行を作成
            ws["A2"] = "アプリ名称"
            ws["B2"] = "フェーズ"
            ws["C2"] = "ベンダー/社員"
            ws["D2"] = "工数（人日）"
            ws["E2"] = "金額"
            current_row = 3
        else:
            ws = sheet_dict[sheet_title]
            current_row = ws.max_row + 1

        # 1行目（メモ欄）をシートの最初に追加（シートごとに1回だけ）
        if not hasattr(ws, '_first_data_written'):
            ws["A1"] = f"{app_name}_{vendor_employee}"
            ws._first_data_written = True

        # 3行目以降：データ行
        for phase_name, phase_data in sorted(phases_data.items()):
            ws[f"A{current_row}"] = app_name
            ws[f"B{current_row}"] = phase_name
            ws[f"C{current_row}"] = vendor_employee
            ws[f"D{current_row}"] = phase_data["hours"]
            ws[f"E{current_row}"] = phase_data["amount"]
            current_row += 1

    # 全シートのセルの装飾
    for sheet_title in sheet_dict:
        ws = sheet_dict[sheet_title]
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
    for col in ["A", "B", "C", "D", "E"]:
        cell = ws[f"{col}1"]
        if cell.value is not None:
            cell.fill = memo_fill
            cell.font = memo_font
            cell.border = thin_border
            cell.alignment = left_align

    # 2行目（ヘッダー行）のフォーマット
    for col in ["A", "B", "C", "D", "E"]:
        cell = ws[f"{col}2"]
        if cell.value is not None:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = center_align

    # データ行のフォーマット
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=5):
        for cell in row:
            if cell.value is not None:
                cell.border = thin_border
                if cell.column in [4, 5]:  # D, E 列は右揃え（数値）
                    cell.alignment = right_align
                else:  # A, B, C 列は左揃え（テキスト）
                    cell.alignment = left_align

    # 列幅を調整
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15


def create_matching_check_workbook():
    """
    報告用Excelファイル（資産・費用の2ファイル）から
    突き合わせチェック用Excelを新規生成する。
    """
    output_dir = Path("output")

    # 報告用Excelファイルのパスを指定
    asset_file = output_dir / "houkoku_shisan.xlsx"
    expense_file = output_dir / "houkoku_hiyo.xlsx"

    # 新しいワークブックを作成
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="突き合わせ")

    # ヘッダー行（2行目）
    ws["A2"] = "報告ファイル名"
    ws["B2"] = "シート名"
    ws["C2"] = "統合情報"
    ws["D2"] = "フェーズ"
    ws["E2"] = "ベンダー/社員"
    ws["F2"] = "工数（人日）"
    ws["G2"] = "金額"

    current_row = 3

    # 両ファイルからデータを読み込んで転記
    for file_path, file_display_name in [(asset_file, "houkoku_shisan.xlsx"),
                                          (expense_file, "houkoku_hiyo.xlsx")]:
        if not file_path.exists():
            print(f"[WARN] ファイルが見つかりません: {file_path}")
            continue

        try:
            wb_temp = load_workbook(file_path)

            for sheet_name in wb_temp.sheetnames:
                ws_temp = wb_temp[sheet_name]

                # 1行目から統合情報を取得
                merged_info = ws_temp["A1"].value if ws_temp["A1"].value else "N/A"

                # 3行目以降のデータを転記
                for row in ws_temp.iter_rows(min_row=3, max_row=ws_temp.max_row,
                                             min_col=1, max_col=5, values_only=True):
                    if row[0] is None:  # A列が空なら終了
                        break

                    app_name = row[0]          # A列：アプリ名称
                    phase_name = row[1]        # B列：フェーズ
                    vendor_employee = row[2]   # C列：ベンダー/社員
                    hours = row[3]             # D列：工数（人日）
                    amount = row[4]            # E列：金額

                    ws[f"A{current_row}"] = file_display_name
                    ws[f"B{current_row}"] = sheet_name
                    ws[f"C{current_row}"] = merged_info
                    ws[f"D{current_row}"] = phase_name
                    ws[f"E{current_row}"] = vendor_employee
                    ws[f"F{current_row}"] = hours
                    ws[f"G{current_row}"] = amount

                    current_row += 1

            wb_temp.close()
        except Exception as e:
            print(f"[ERROR] ファイル読み込みエラー {file_path}: {e}")

    # ヘッダー行のスタイル設定
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    # 2行目（ヘッダー行）のスタイル設定
    for col in ["A", "B", "C", "D", "E", "F", "G"]:
        cell = ws[f"{col}2"]
        if cell.value is not None:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = center_align

    # データ行のスタイル設定
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=7):
        for cell in row:
            if cell.value is not None:
                cell.border = thin_border
                if cell.column in [6, 7]:  # F, G 列は右揃え（数値）
                    cell.alignment = right_align
                else:
                    cell.alignment = left_align

    # 列幅を調整
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 25
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 15
    ws.column_dimensions["G"].width = 15

    # タイムスタンプ付きファイル名で保存
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = output_dir / f"matching_check_{timestamp}.xlsx"
    wb.save(output_file)
    print(f"[OK] Created: {output_file}")


def main():
    """メイン処理"""
    try:
        print("=" * 50)
        print("Excel工数管理自動化を開始します")
        print("=" * 50)

        # 1. 設定Excelを読み込み
        print("\n[Step 1] Reading configuration file...")
        apps, phases, employee_cost_per_day, work_data = read_config()
        print(f"  [OK] Number of apps: {len(apps)}")
        print(f"  [OK] Number of phases: {len(phases)}")
        print(f"  [OK] Employee cost per day: {employee_cost_per_day:,.0f} JPY")
        print(f"  [OK] Number of work records: {len(work_data)}")

        # 2. 工数データを資産/費用に分類
        print("\n[Step 2] Classifying work data...")
        asset_data, expense_data = organize_data_by_division(work_data, phases, employee_cost_per_day)
        print(f"  [OK] Asset groups: {len(asset_data)}")
        print(f"  [OK] Expense groups: {len(expense_data)}")

        # 3. 報告用Excelを生成
        print("\n[Step 3] Generating report files...")
        create_workbook(apps, asset_data, "houkoku_shisan.xlsx")
        create_workbook(apps, expense_data, "houkoku_hiyo.xlsx")

        # 4. 突き合わせチェック用Excelを生成
        print("\n[Step 4] Generating matching check file...")
        create_matching_check_workbook()

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
