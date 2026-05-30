"""
Excel工数管理自動化スクリプト

[目的]
作業用Excelの工数データを読み込み、資産/費用に分けた報告用Excelを自動生成する。

[処理フロー]
1. config.xlsx から マスタデータ（アプリ、フェーズ）と工数を読み込み
2. フェーズの区分（資産/費用）に基づいて、工数データを分類
3. 報告_資産.xlsx と 報告_費用.xlsx の 2ファイルを自動生成
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
        phases: {フェーズID: (フェーズ名, 区分)}  ※区分は "資産" or "費用"
        work_data: [(アプリID, フェーズID, 工数), ...]
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
        phase_id, phase_name, division = row[0], row[1], row[2]
        phases[phase_id] = (phase_name, division)

    # 工数データを読み込み
    work_data = []
    ws_work = wb["工数"]
    for row in ws_work.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        app_id, phase_id, hours = row[0], row[1], row[2]
        work_data.append((app_id, phase_id, hours))

    wb.close()
    return apps, phases, work_data


def organize_data_by_division(work_data: List, phases: Dict) -> Tuple[Dict, Dict]:
    """
    工数データをフェーズの区分（資産/費用）に基づいて分類する。

    Args:
        work_data: [(アプリID, フェーズID, 工数), ...]
        phases: {フェーズID: (フェーズ名, 区分)}

    Returns:
        asset_data: {アプリID: {フェーズID: 工数, ...}, ...}
        expense_data: {アプリID: {フェーズID: 工数, ...}, ...}
    """
    asset_data = {}
    expense_data = {}

    for app_id, phase_id, hours in work_data:
        if phase_id not in phases:
            print(f"警告: フェーズID '{phase_id}' がマスタに見つかりません")
            continue

        phase_name, division = phases[phase_id]

        # 辞書の初期化
        if app_id not in asset_data:
            asset_data[app_id] = {}
            expense_data[app_id] = {}

        # 資産/費用に分類
        if division == "資産":
            asset_data[app_id][phase_id] = hours
        elif division == "費用":
            expense_data[app_id][phase_id] = hours
        else:
            print(f"警告: 不正な区分 '{division}' (フェーズID: {phase_id})")

    return asset_data, expense_data


def create_workbook(apps: Dict, phases: Dict, data_by_app: Dict, file_name: str):
    """
    報告用Excelを生成する。

    Args:
        apps: {アプリID: アプリ名}
        phases: {フェーズID: (フェーズ名, 区分)}
        data_by_app: {アプリID: {フェーズID: 工数, ...}, ...}
        file_name: 出力ファイル名（例: "報告_資産.xlsx"）
    """
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)  # デフォルトシートを削除

    # アプリごとにシートを作成
    for app_id in sorted(data_by_app.keys()):
        if not data_by_app[app_id]:  # 該当するフェーズがない場合はスキップ
            continue

        app_name = apps.get(app_id, app_id)
        ws = wb.create_sheet(title=app_name[:31])  # Excelのシート名は31文字まで

        # ヘッダー行を作成
        ws["A1"] = "フェーズ"
        ws["B1"] = "工数（人日）"

        # フェーズと工数を記入
        row = 2
        for phase_id, hours in sorted(data_by_app[app_id].items()):
            phase_name = phases[phase_id][0]
            ws[f"A{row}"] = phase_name
            ws[f"B{row}"] = hours
            row += 1

        # セルの装飾
        format_workbook(ws)

    # ファイルを保存
    output_path = Path("output") / file_name
    wb.save(output_path)
    print(f"[OK] 作成完了: {output_path}")


def format_workbook(ws):
    """
    ワークシートのセルをフォーマットする（背景色、フォント、枠線など）。
    """
    # スタイル定義
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # ヘッダー行のフォーマット
    for cell in ws[1]:
        if cell.value is not None:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border

    # データ行のフォーマット
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=2):
        for cell in row:
            if cell.value is not None:
                cell.border = thin_border
                if cell.column == 2:  # B列（工数）は右揃え
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

    # 列幅を調整
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 15


def main():
    """メイン処理"""
    try:
        print("=" * 50)
        print("Excel工数管理自動化を開始します")
        print("=" * 50)

        # 1. 設定Excelを読み込み
        print("\n[ステップ1] 設定Excelを読み込み中...")
        apps, phases, work_data = read_config()
        print(f"  [OK] アプリ数: {len(apps)}")
        print(f"  [OK] フェーズ数: {len(phases)}")
        print(f"  [OK] 工数レコード数: {len(work_data)}")

        # 2. 工数データを資産/費用に分類
        print("\n[ステップ2] 工数データを分類中...")
        asset_data, expense_data = organize_data_by_division(work_data, phases)
        print(f"  [OK] 資産に該当するアプリ: {len(asset_data)}")
        print(f"  [OK] 費用に該当するアプリ: {len(expense_data)}")

        # 3. 報告用Excelを生成
        print("\n[ステップ3] 報告用Excelを生成中...")
        create_workbook(apps, phases, asset_data, "報告_資産.xlsx")
        create_workbook(apps, phases, expense_data, "報告_費用.xlsx")

        print("\n" + "=" * 50)
        print("[OK] 処理が正常に完了しました！")
        print("=" * 50)

    except FileNotFoundError as e:
        print(f"[ERROR] エラー: {e}")
    except Exception as e:
        print(f"[ERROR] 予期しないエラーが発生しました: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
