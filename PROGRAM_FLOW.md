# プログラムの処理フロー詳解（修正版）

このドキュメントでは、`excel_automation.py` が内部的にどのように動作しているのか、ステップバイステップで解説します。

---

## 全体の処理フロー図

```
[スタート]
    ↓
[read_config()]
  ├─ input/config.xlsx を開く
  ├─ 「アプリ」シートを読み込み → apps 辞書を作成
  ├─ 「フェーズ」シートを読み込み → phases 辞書を作成（フェーズ名で管理）
  └─ 「工数」シートを読み込み → work_data リストを作成
    ↓
[organize_data_by_division()]
  ├─ work_data の各レコードをループ
  ├─ phases から「区分」（資産/費用）を確認
  ├─ アプリ×ベンダーをキーにしてグループ化
  ├─ 「資産」なら asset_data に追加
  └─ 「費用」なら expense_data に追加
    ↓
[create_workbook()（資産用）]
  ├─ 新しい Workbook を作成
  ├─ asset_data の各（アプリ×ベンダー）ごとにシートを作成（詳細_1, 詳細_2, ...）
  ├─ 1行目：メモ欄（アプリ名、ベンダー名、社員工数）
  ├─ 2行目：ヘッダー行（フェーズ名、ベンダー工数、発注金額、社員工数）
  ├─ 3行目以降：データ行
  ├─ format_workbook() でスタイルを適用
  └─ output/houkoku_shisan.xlsx として保存
    ↓
[create_workbook()（費用用）]
  ├─ 新しい Workbook を作成
  ├─ expense_data の各（アプリ×ベンダー）ごとにシートを作成
  ├─ 同様の行・列構成で出力
  └─ output/houkoku_hiyo.xlsx として保存
    ↓
[エンド]
```

---

## ステップ1：read_config() の詳細

### 目的
設定Excel（config.xlsx）から必要なマスタデータと工数（ベンダー・社員）を読み込み、Python内のデータ構造（辞書・リスト）に変換する。

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

**1-3) 「フェーズ」シートを読み込み（新形式）**
```python
phases = {}
ws_phases = wb["フェーズ"]
for row in ws_phases.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        break
    phase_name, division = row[0], row[1]  # フェーズ名で管理
    phases[phase_name] = division
```

**結果**:
```python
phases = {
    "P001_要件定義": "資産",
    "P002_基本設計": "資産",
    "P004_実装": "費用",
    ...
}
```

**1-4) 「工数」シートを読み込み（ベンダー・社員統合）**
```python
work_data = []
ws_work = wb["工数"]
for row in ws_work.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        break
    app_id, vendor_name, phase_name, vendor_hours, amount, employee_hours = row[0:6]
    work_data.append((app_id, vendor_name, phase_name, vendor_hours, amount, employee_hours))
```

**結果**:
```python
work_data = [
    ("APP001", "ベンダーA", "P001_要件定義", 5, 150000, 2),
    ("APP001", "ベンダーA", "P002_基本設計", 8, 240000, 3),
    ("APP001", "ベンダーB", "P004_実装", 15, 450000, 1),
    ...
]
```

---

## ステップ2：organize_data_by_division() の詳細

### 目的
work_data の各レコードをループして、フェーズの「区分」（資産/費用）に基づいて2つのグループに分け、さらにアプリ×ベンダー単位でグループ化する。

### 処理ステップ

**2-1) 初期化**
```python
asset_data = {}
expense_data = {}
```

**2-2) work_data をループ（アプリ×ベンダー単位でグループ化）**
```python
for app_id, vendor_name, phase_name, vendor_hours, amount, employee_hours in work_data:
    # フェーズが存在するか確認
    if phase_name not in phases:
        print(f"[WARN] フェーズ '{phase_name}' がマスタに見つかりません")
        continue
    
    division = phases[phase_name]  # 「区分」を取得
    key = (app_id, vendor_name)    # アプリ×ベンダーをキーにする
    
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
    
    # 区分に応じて振り分け
    if division == "資産":
        asset_data[key][phase_name] = phase_data
    elif division == "費用":
        expense_data[key][phase_name] = phase_data
```

### 結果イメージ

**asset_data（資産に区分されたもの）**:
```python
asset_data = {
    ("APP001", "ベンダーA"): {
        "P001_要件定義": {
            "vendor_hours": 5,
            "amount": 150000,
            "employee_hours": 2
        },
        "P002_基本設計": {
            "vendor_hours": 8,
            "amount": 240000,
            "employee_hours": 3
        }
    },
    ("APP001", "ベンダーB"): {
        ...
    },
    ...
}
```

**expense_data（費用に区分されたもの）**:
```python
expense_data = {
    ("APP001", "ベンダーB"): {
        "P004_実装": {
            "vendor_hours": 15,
            "amount": 450000,
            "employee_hours": 1
        },
        "P005_単体テスト": {
            "vendor_hours": 10,
            "amount": 300000,
            "employee_hours": 2
        }
    },
    ...
}
```

---

## ステップ3：create_workbook() の詳細

### 目的
分類されたデータから、報告用のExcelファイルを生成する。

### 処理ステップ（例：houkoku_shisan.xlsx を生成する場合）

**3-1) 新しい Workbook を作成**
```python
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

wb = Workbook()
wb.remove(wb.active)  # デフォルトシートを削除
```

**3-2) アプリ×ベンダーごとにシートを作成（固定シート名で採番）**
```python
sheet_num = 1

for (app_id, vendor_name) in sorted(data_by_app_vendor.keys()):
    if not data_by_app_vendor[(app_id, vendor_name)]:  # データがない場合はスキップ
        continue
    
    app_name = apps.get(app_id, app_id)
    sheet_title = f"詳細_{sheet_num}"  # シート名は固定：詳細_1, 詳細_2, ...
    ws = wb.create_sheet(title=sheet_title)
    sheet_num += 1
```

**3-3) 1行目：メモ欄を作成**
```python
ws["A1"] = app_name           # アプリ名
ws["B1"] = vendor_name        # ベンダー名
ws["C1"] = "社員工数"          # ラベル

# 社員工数の合計を計算
phases_data = data_by_app_vendor[(app_id, vendor_name)]
total_employee_hours = sum(phase["employee_hours"] for phase in phases_data.values())
ws["D1"] = total_employee_hours  # 社員工数の合計
```

**結果（例）**:
```
A1: ユーザー管理システム  | B1: ベンダーA | C1: 社員工数 | D1: 5
```

**3-4) 2行目：ヘッダー行を作成**
```python
ws["A2"] = "フェーズ名"
ws["B2"] = "ベンダー工数（人日）"
ws["C2"] = "発注金額"
ws["D2"] = "社員工数（人日）"
```

**3-5) 3行目以降：データ行を作成**
```python
row = 3
for phase_name, phase_data in sorted(phases_data.items()):
    ws[f"A{row}"] = phase_name
    ws[f"B{row}"] = phase_data["vendor_hours"]
    ws[f"C{row}"] = phase_data["amount"]
    ws[f"D{row}"] = phase_data["employee_hours"]
    row += 1
```

**例：APP001 × ベンダーA のシート内容**
```
1行目: ユーザー管理システム | ベンダーA | 社員工数 | 5
2行目: フェーズ名 | ベンダー工数（人日） | 発注金額 | 社員工数（人日）
3行目: P001_要件定義 | 5 | 150000 | 2
4行目: P002_基本設計 | 8 | 240000 | 3
```

**3-6) セルをフォーマット**
```python
format_workbook(ws)
```

**3-7) ファイルを保存**
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
# メモ欄（1行目）のスタイル
memo_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
memo_font = Font(bold=False, size=10)

# ヘッダー行（2行目）のスタイル
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF", size=11)

# 配置
center_align = Alignment(horizontal="center", vertical="center")
left_align = Alignment(horizontal="left", vertical="center")
right_align = Alignment(horizontal="right", vertical="center")

# 枠線
thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)
```

**4-2) 1行目（メモ欄）に適用**
```python
for col in ["A", "B", "C", "D"]:
    cell = ws[f"{col}1"]
    if cell.value is not None:
        cell.fill = memo_fill      # 黄色背景
        cell.font = memo_font
        cell.border = thin_border
        cell.alignment = left_align
```

**結果**:
- 背景色が黄色になり、枠線で囲まれる

**4-3) 2行目（ヘッダー行）に適用**
```python
for col in ["A", "B", "C", "D"]:
    cell = ws[f"{col}2"]
    if cell.value is not None:
        cell.fill = header_fill     # 青色背景
        cell.font = header_font     # 白い太字
        cell.border = thin_border
        cell.alignment = center_align  # 中央揃え
```

**結果**:
- 背景色が青色になり、フォントが白い太字に

**4-4) 3行目以降（データ行）に適用**
```python
for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=4):
    for cell in row:
        if cell.value is not None:
            cell.border = thin_border
            if cell.column in [2, 3, 4]:  # B, C, D 列（数値）は右揃え
                cell.alignment = right_align
            else:  # A 列（テキスト）は左揃え
                cell.alignment = left_align
```

**4-5) 列幅を調整**
```python
ws.column_dimensions["A"].width = 25
ws.column_dimensions["B"].width = 18
ws.column_dimensions["C"].width = 15
ws.column_dimensions["D"].width = 18
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
フェーズ名 | 区分
P001_要件定義 | 資産
P002_基本設計 | 資産
P004_実装 | 費用
```

**config.xlsx - 「工数」**
```
アプリID | ベンダー名 | フェーズ名 | ベンダー工数 | 発注金額 | 社員工数
APP001 | ベンダーA | P001_要件定義 | 5 | 150000 | 2
APP001 | ベンダーA | P002_基本設計 | 8 | 240000 | 3
APP001 | ベンダーB | P004_実装 | 15 | 450000 | 1
APP002 | ベンダーC | P001_要件定義 | 10 | 300000 | 4
```

### 処理中のデータ

**ステップ1後**:
```python
apps = {"APP001": "ユーザー管理", "APP002": "決済システム"}
phases = {"P001_要件定義": "資産", "P002_基本設計": "資産", "P004_実装": "費用"}
work_data = [
    ("APP001", "ベンダーA", "P001_要件定義", 5, 150000, 2),
    ("APP001", "ベンダーA", "P002_基本設計", 8, 240000, 3),
    ("APP001", "ベンダーB", "P004_実装", 15, 450000, 1),
    ("APP002", "ベンダーC", "P001_要件定義", 10, 300000, 4)
]
```

**ステップ2後**:
```python
asset_data = {
    ("APP001", "ベンダーA"): {
        "P001_要件定義": {"vendor_hours": 5, "amount": 150000, "employee_hours": 2},
        "P002_基本設計": {"vendor_hours": 8, "amount": 240000, "employee_hours": 3}
    },
    ("APP002", "ベンダーC"): {
        "P001_要件定義": {"vendor_hours": 10, "amount": 300000, "employee_hours": 4}
    }
}

expense_data = {
    ("APP001", "ベンダーB"): {
        "P004_実装": {"vendor_hours": 15, "amount": 450000, "employee_hours": 1}
    }
}
```

### 出力例

**houkoku_shisan.xlsx**
```
[詳細_1 シート]
1行目: ユーザー管理 | ベンダーA | 社員工数 | 5
2行目: フェーズ名 | ベンダー工数（人日） | 発注金額 | 社員工数（人日）
3行目: P001_要件定義 | 5 | 150000 | 2
4行目: P002_基本設計 | 8 | 240000 | 3

[詳細_2 シート]
1行目: 決済システム | ベンダーC | 社員工数 | 4
2行目: フェーズ名 | ベンダー工数（人日） | 発注金額 | 社員工数（人日）
3行目: P001_要件定義 | 10 | 300000 | 4
```

**houkoku_hiyo.xlsx**
```
[詳細_1 シート]
1行目: ユーザー管理 | ベンダーB | 社員工数 | 1
2行目: フェーズ名 | ベンダー工数（人日） | 発注金額 | 社員工数（人日）
3行目: P004_実装 | 15 | 450000 | 1
```

---

## エラーハンドリング

プログラムは以下の例外に対応しています：

```python
try:
    # メイン処理
except FileNotFoundError as e:
    print(f"[ERROR] {e}")
except Exception as e:
    print(f"[ERROR] Unexpected error: {e}")
    import traceback
    traceback.print_exc()
```

**よくあるエラー**:

| エラー | 原因 | 対策 |
|--------|------|------|
| `FileNotFoundError` | config.xlsx が見つからない | `create_sample_config.py` を実行 |
| `KeyError` | 工数シートのフェーズ名が不正 | config.xlsx を確認 |
| `ValueError` | 工数が数値でない | config.xlsx の工数列を確認 |

---

## パフォーマンスと最適化

このプログラムは以下のように最適化されています：

1. **ジェネレータの使用**: `iter_rows()` でメモリ効率良く行をループ
2. **辞書のキャッシング**: アプリ・フェーズ・（アプリ×ベンダー）を辞書に保存して高速検索
3. **ループの最小化**: 必要な処理だけを実行（ベンダーやアプリで不要なループなし）
4. **タプルをキーに使用**: (app_id, vendor_name) をキーにしてO(1)で検索可能

---

## 拡張例

### 例1：社員工数の合計を自動計算して表示

`create_workbook()` 関数内で既に実装されています：

```python
total_employee_hours = sum(phase["employee_hours"] for phase in phases_data.values())
ws["D1"] = total_employee_hours
```

### 例2：ベンダーごとの発注金額の合計を追加

`create_workbook()` 関数内で、データ行の後に合計行を追加：

```python
total_amount = sum(phase["amount"] for phase in phases_data.values())
ws[f"A{row}"] = "合計"
ws[f"C{row}"] = total_amount
```

### 例3：複数の区分（資産/費用の他に「保守」など）を対応

`organize_data_by_division()` 関数を拡張し、main() で新しい出力を追加：

```python
# organize_data_by_division() で
maintenance_data = {}  # 新しい区分用の辞書
if division == "保守":
    maintenance_data[key][phase_name] = phase_data

# main() で
create_workbook(apps, maintenance_data, "houkoku_hoshu.xlsx")
```

---

## まとめ

このプログラムは、**入力 → 処理 → 出力** の3段階で構成されています：

1. **入力（read_config）**: マスタデータと工数を読み込み
2. **処理（organize_data_by_division）**: データを資産/費用に分類、アプリ×ベンダー単位でグループ化
3. **出力（create_workbook）**: 報告用Excelを生成（1行目メモ欄、2行目ヘッダー、3行目以降データ）

各関数は独立していて、修正・拡張が容易な設計になっています。
