# プログラムの処理フロー詳解

このドキュメントでは、`excel_automation.py` が内部的にどのように動作しているのか、ステップバイステップで解説します。

---

## 全体の処理フロー図

```
[スタート]
    ↓
[read_config()]
  ├─ input/config.xlsx を開く
  ├─ 「アプリ」シートを読み込み → apps 辞書を作成
  ├─ 「フェーズ」シートを読み込み → phases 辞書を作成
  └─ 「工数」シートを読み込み → work_data リストを作成
    ↓
[organize_data_by_division()]
  ├─ work_data の各レコードをループ
  ├─ phases から「区分」（資産/費用）を確認
  ├─ 「資産」なら asset_data に追加
  └─ 「費用」なら expense_data に追加
    ↓
[create_workbook()（資産用）]
  ├─ 新しい Workbook を作成
  ├─ asset_data の各アプリごとにシートを作成
  ├─ format_workbook() でスタイルを適用
  └─ output/報告_資産.xlsx として保存
    ↓
[create_workbook()（費用用）]
  ├─ 新しい Workbook を作成
  ├─ expense_data の各アプリごとにシートを作成
  ├─ format_workbook() でスタイルを適用
  └─ output/報告_費用.xlsx として保存
    ↓
[エンド]
```

---

## ステップ1：read_config() の詳細

### 目的
設定Excel（config.xlsx）から必要なマスタデータと工数を読み込み、Python内のデータ構造（辞書・リスト）に変換する。

### 処理ステップ

**1-1) Excelファイルを開く**
```python
wb = load_workbook(config_path)
```

**1-2) 「アプリ」シートを読み込み**
```python
apps = {}
ws_apps = wb["アプリ"]
for row in ws_apps.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        break
    app_id, app_name = row[0], row[1]
    apps[app_id] = app_name
```

**結果**:
```python
apps = {
    "APP001": "ユーザー管理システム",
    "APP002": "決済システム",
    ...
}
```

**1-3) 「フェーズ」シートを読み込み**
```python
phases = {}
ws_phases = wb["フェーズ"]
for row in ws_phases.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        break
    phase_id, phase_name, division = row[0], row[1], row[2]
    phases[phase_id] = (phase_name, division)
```

**結果**:
```python
phases = {
    "P001": ("要件定義", "資産"),
    "P002": ("基本設計", "資産"),
    "P004": ("実装", "費用"),
    ...
}
```

**1-4) 「工数」シートを読み込み**
```python
work_data = []
ws_work = wb["工数"]
for row in ws_work.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        break
    app_id, phase_id, hours = row[0], row[1], row[2]
    work_data.append((app_id, phase_id, hours))
```

**結果**:
```python
work_data = [
    ("APP001", "P001", 5),
    ("APP001", "P002", 8),
    ("APP001", "P004", 15),
    ...
]
```

---

## ステップ2：organize_data_by_division() の詳細

### 目的
work_data の各レコードをループして、フェーズの「区分」（資産/費用）に基づいて2つのグループに分ける。

### 処理ステップ

**2-1) 初期化**
```python
asset_data = {}
expense_data = {}
```

**2-2) work_data をループ**
```python
for app_id, phase_id, hours in work_data:
    # アプリID と フェーズID が妥当か確認
    if phase_id not in phases:
        print(f"警告: フェーズID '{phase_id}' がマスタに見つかりません")
        continue
    
    phase_name, division = phases[phase_id]  # 「区分」を取得
    
    # 辞書の初期化
    if app_id not in asset_data:
        asset_data[app_id] = {}
        expense_data[app_id] = {}
    
    # 区分に応じて振り分け
    if division == "資産":
        asset_data[app_id][phase_id] = hours
    elif division == "費用":
        expense_data[app_id][phase_id] = hours
```

### 結果イメージ

**asset_data（資産に区分されたもの）**:
```python
asset_data = {
    "APP001": {"P001": 5, "P002": 8},
    "APP002": {"P001": 10, "P002": 12, "P003": 20},
    ...
}
```

**expense_data（費用に区分されたもの）**:
```python
expense_data = {
    "APP001": {"P004": 15, "P005": 10},
    "APP002": {"P004": 30, "P006": 18},
    ...
}
```

---

## ステップ3：create_workbook() の詳細

### 目的
分類されたデータから、報告用のExcelファイルを生成する。

### 処理ステップ（例：報告_資産.xlsx を生成する場合）

**3-1) 新しい Workbook を作成**
```python
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

wb = Workbook()
wb.remove(wb.active)  # デフォルトシートを削除
```

**3-2) アプリごとにシートを作成**
```python
for app_id in sorted(asset_data.keys()):
    if not asset_data[app_id]:  # データがない場合はスキップ
        continue
    
    app_name = apps.get(app_id, app_id)
    ws = wb.create_sheet(title=app_name[:31])
```

**3-3) ヘッダー行を作成**
```python
ws["A1"] = "フェーズ"
ws["B1"] = "工数（人日）"
```

**3-4) フェーズと工数を記入**
```python
row = 2
for phase_id, hours in sorted(asset_data[app_id].items()):
    phase_name = phases[phase_id][0]
    ws[f"A{row}"] = phase_name  # A列にフェーズ名
    ws[f"B{row}"] = hours        # B列に工数
    row += 1
```

**例：APP001 のシート内容**
```
フェーズ           | 工数（人日）
要件定義           | 5
基本設計           | 8
```

**3-5) セルをフォーマット**
```python
format_workbook(ws)
```

**3-6) ファイルを保存**
```python
output_path = Path("output") / file_name
wb.save(output_path)
```

---

## ステップ4：format_workbook() の詳細

### 目的
ワークシートのセルを装飾して、見やすくする。

### 処理ステップ

**4-1) スタイルを定義**
```python
# ヘッダー行のスタイル
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF", size=11)

# 中央揃え
center_align = Alignment(horizontal="center", vertical="center")

# 枠線
thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)
```

**4-2) ヘッダー行に適用**
```python
for cell in ws[1]:
    if cell.value is not None:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
```

**結果**: 
- 背景色が青色になり、フォントが白い太字に

**4-3) データ行に適用**
```python
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=2):
    for cell in row:
        if cell.value is not None:
            cell.border = thin_border
            if cell.column == 2:  # B列（工数）
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:  # A列（フェーズ）
                cell.alignment = Alignment(horizontal="left", vertical="center")
```

**4-4) 列幅を調整**
```python
ws.column_dimensions["A"].width = 25
ws.column_dimensions["B"].width = 15
```

---

## データフロー例

### 入力例

**config.xlsx - 「アプリ」**
```
アプリID | アプリ名
APP001  | ユーザー管理
APP002  | 決済システム
```

**config.xlsx - 「フェーズ」**
```
フェーズID | フェーズ名 | 区分
P001      | 要件定義   | 資産
P002      | 基本設計   | 資産
P004      | 実装       | 費用
```

**config.xlsx - 「工数」**
```
アプリID | フェーズID | 工数
APP001  | P001       | 5
APP001  | P002       | 8
APP001  | P004       | 15
APP002  | P001       | 10
APP002  | P002       | 12
APP002  | P004       | 30
```

### 処理中のデータ

**ステップ1後**:
```python
apps = {"APP001": "ユーザー管理", "APP002": "決済システム"}
phases = {"P001": ("要件定義", "資産"), "P002": ("基本設計", "資産"), "P004": ("実装", "費用")}
work_data = [("APP001", "P001", 5), ("APP001", "P002", 8), ("APP001", "P004", 15), ...]
```

**ステップ2後**:
```python
asset_data = {
    "APP001": {"P001": 5, "P002": 8},
    "APP002": {"P001": 10, "P002": 12}
}

expense_data = {
    "APP001": {"P004": 15},
    "APP002": {"P004": 30}
}
```

### 出力例

**報告_資産.xlsx**
```
[シート：ユーザー管理]
フェーズ   | 工数（人日）
要件定義   | 5
基本設計   | 8

[シート：決済システム]
フェーズ   | 工数（人日）
要件定義   | 10
基本設計   | 12
```

**報告_費用.xlsx**
```
[シート：ユーザー管理]
フェーズ   | 工数（人日）
実装       | 15

[シート：決済システム]
フェーズ   | 工数（人日）
実装       | 30
```

---

## エラーハンドリング

プログラムは以下の例外に対応しています：

```python
try:
    # メイン処理
except FileNotFoundError as e:
    print(f"[ERROR] エラー: {e}")
except Exception as e:
    print(f"[ERROR] 予期しないエラーが発生しました: {e}")
    import traceback
    traceback.print_exc()
```

**よくあるエラー**:

| エラー | 原因 | 対策 |
|--------|------|------|
| `FileNotFoundError` | config.xlsx が見つからない | `create_sample_config.py` を実行 |
| `KeyError` | 工数シートのアプリ/フェーズが不正 | config.xlsx を確認 |
| `ValueError` | 工数が数値でない | config.xlsx の工数列を確認 |

---

## メモリ効率とパフォーマンス

このプログラムは以下のように最適化されています：

1. **ジェネレータの使用**: `iter_rows()` でメモリ効率良く行をループ
2. **辞書のキャッシング**: アプリ・フェーズを辞書に保存して高速検索
3. **ループの最小化**: 必要な処理だけを実行

---

## 拡張例

### 例1：工数の合計を自動計算

`create_workbook()` 関数内に以下を追加：

```python
# データ行の後に合計行を追加
total_hours = sum(data_by_app[app_id].values())
ws[f"A{row}"] = "合計"
ws[f"B{row}"] = total_hours

# 合計行を太字・背景色付き
ws[f"A{row}"].fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
ws[f"B{row}"].fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
ws[f"B{row}"].font = Font(bold=True)
```

### 例2：複数の区分（資産・費用の他に「保守」など）を対応

フェーズマスタに「保守」を追加し、`organize_data_by_division()` を拡張：

```python
asset_data = {}
expense_data = {}
maintenance_data = {}

if division == "資産":
    asset_data[app_id][phase_id] = hours
elif division == "費用":
    expense_data[app_id][phase_id] = hours
elif division == "保守":
    maintenance_data[app_id][phase_id] = hours
```

その後、メイン処理で報告_保守.xlsx も生成：

```python
create_workbook(apps, phases, maintenance_data, "報告_保守.xlsx")
```

---

## まとめ

このプログラムは、**入力 → 処理 → 出力** の3段階で構成されています：

1. **入力（read_config）**: マスタデータと工数を読み込み
2. **処理（organize_data_by_division）**: データを資産/費用に分類
3. **出力（create_workbook）**: 報告用Excelを生成

各関数は独立していて、修正・拡張が容易な設計になっています。
