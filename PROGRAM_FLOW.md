# プログラムの処理フロー詳解（新仕様版）

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
  ├─ アプリ×ベンダー/社員をキーにしてグループ化
  ├─ 「資産」なら asset_data に追加
  └─ 「費用」なら expense_data に追加
    ↓
[create_workbook()（資産用）]
  ├─ テンプレート template/houkoku_template.xlsx をコピー
  ├─ output/houkoku_shisan.xlsx として複製
  ├─ テンプレートの15シート（詳細_1～詳細_15）に対してデータを上書き
  ├─ 1行目：メモ欄（アプリ名_ベンダー/社員）
  ├─ 2行目：ヘッダー行（アプリ名、フェーズ、ベンダー/社員、工数、金額）
  ├─ 3行目以降：データ行（カンマ区切り形式で表示）
  └─ ファイルを保存
    ↓
[create_workbook()（費用用）]
  ├─ テンプレート template/houkoku_template.xlsx をコピー
  ├─ output/houkoku_hiyo.xlsx として複製
  ├─ テンプレートの15シートに対してデータを上書き
  └─ ファイルを保存
    ↓
[create_matching_check_workbook()]
  ├─ output/houkoku_shisan.xlsx と output/houkoku_hiyo.xlsx を読み込み
  ├─ 新しい Workbook を作成
  ├─ 「突き合わせ」シートを作成
  ├─ 両ファイルのデータを統合
  └─ output/matching_check_[タイムスタンプ].xlsx として保存
    ↓
[エンド]
```

---

## CLIオプション

### 通常モード（引数なし）
```bash
python excel_automation.py
```
報告Excel（資産・費用）を作成し、自動で突き合わせチェックを生成

### チェックのみモード
```bash
python excel_automation.py --check-only
```
既存の報告Excel（output/houkoku_shisan.xlsx, houkoku_hiyo.xlsx）から突き合わせチェックを生成

カスタムパス指定：
```bash
python excel_automation.py --check-only --asset-file path/to/asset.xlsx --expense-file path/to/expense.xlsx
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

**1-4) 「工数」シートを読み込み（7列構成）**
```python
work_data = []
ws_work = wb["工数"]
for row in ws_work.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        break
    # A:アプリID, B:ベンダー/社員, C:フェーズ名, D:工数, E:発注金額, F:社員コスト, G:出力シート指定
    app_id = row[0]
    vendor_employee = row[1]  # ベンダー名 または 「社員」
    phase_name = row[2]
    hours = row[3]
    vendor_amount = row[4]  # ベンダーの場合は金額、社員の場合はNone
    employee_cost = row[5]  # 社員の場合はコスト、ベンダーの場合はNone
    output_sheet = row[6]  # 出力シート指定
    
    work_data.append((app_id, vendor_employee, phase_name, hours, vendor_amount, employee_cost, output_sheet))
```

**結果**:
```python
work_data = [
    ("APP001", "ベンダーA", "P001_要件定義", 5, 150000, None, "詳細_1"),
    ("APP001", "社員", "P001_要件定義", 2, None, 200000, "詳細_1"),
    ("APP001", "ベンダーA", "P002_基本設計", 8, 240000, None, "詳細_1"),
    ("APP001", "ベンダーB", "P004_実装", 15, 450000, None, "詳細_2"),
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

**2-2) work_data をループ（アプリ×ベンダー/社員×出力シート単位でグループ化）**
```python
for app_id, vendor_employee, phase_name, hours, vendor_amount, employee_cost, output_sheet in work_data:
    # フェーズが存在するか確認
    if phase_name not in phases:
        print(f"[WARN] フェーズ '{phase_name}' がマスタに見つかりません")
        continue
    
    division = phases[phase_name]  # 「区分」を取得
    key = (app_id, vendor_employee, output_sheet)  # アプリ×ベンダー/社員×出力シート指定をキーにする
    
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
    ("APP001", "ベンダーA", "詳細_1"): {
        "P001_要件定義": {
            "hours": 5,
            "amount": 150000
        },
        "P002_基本設計": {
            "hours": 8,
            "amount": 240000
        }
    },
    ("APP001", "社員", "詳細_1"): {
        "P001_要件定義": {
            "hours": 2,
            "amount": 200000  # 2 × 100,000
        }
    },
    ...
}
```

**expense_data（費用に区分されたもの）**:
```python
expense_data = {
    ("APP001", "ベンダーB", "詳細_2"): {
        "P004_実装": {
            "hours": 15,
            "amount": 450000
        }
    },
    ...
}
```

---

## ステップ3：create_workbook() の詳細

### 目的
テンプレートをコピーして、報告用のExcelファイルを生成する。
テンプレートには15シート（詳細_1～詳細_15）が事前に用意されているため、シート作成は不要。

### 処理ステップ（例：houkoku_shisan.xlsx を生成する場合）

**3-1) テンプレートをコピー**
```python
template_path = Path("template") / "houkoku_template.xlsx"
output_dir = Path("output")
output_path = output_dir / file_name

shutil.copy(template_path, output_path)
```

**3-2) コピーしたファイルを開く**
```python
wb = load_workbook(output_path)
```

**3-3) アプリ×ベンダー/社員×出力シート指定に基づいて、既存シートにデータを上書き**
```python
sheet_row_dict = {}  # {sheet_name: 次に書き込む行番号}

for (app_id, vendor_employee, output_sheet_name) in sorted(data_by_group.keys()):
    if not data_by_group[(app_id, vendor_employee, output_sheet_name)]:
        continue
    
    app_name = apps.get(app_id, app_id)
    phases_data = data_by_group[(app_id, vendor_employee, output_sheet_name)]
    
    # シート名を決定（出力シート指定値）
    sheet_title = str(output_sheet_name) if output_sheet_name else "詳細_1"
    
    # シートが存在するか確認
    if sheet_title not in wb.sheetnames:
        continue
    
    ws = wb[sheet_title]
```

**3-4) 初回時に1行目（メモ欄）を設定**
```python
if sheet_title not in sheet_row_dict:
    ws["A1"] = f"{app_name}_{vendor_employee}"
    sheet_row_dict[sheet_title] = 3
```

**3-5) 3行目以降：データ行を上書き**
```python
current_row = sheet_row_dict[sheet_title]

for phase_name, phase_data in sorted(phases_data.items()):
    ws[f"A{current_row}"] = app_name
    ws[f"B{current_row}"] = phase_name
    ws[f"C{current_row}"] = vendor_employee
    ws[f"D{current_row}"] = phase_data["hours"]         # 工数（カンマ区切り）
    ws[f"E{current_row}"] = phase_data["amount"]        # 金額（カンマ区切り）
    current_row += 1

sheet_row_dict[sheet_title] = current_row
```

**例：詳細_1 シートのデータ上書き後**
```
1行目: ユーザー管理システム_ベンダーA
2行目: アプリ名称 | フェーズ | ベンダー/社員 | 工数（人日） | 金額
3行目: ユーザー管理システム | P001_要件定義 | ベンダーA | 5 | 150,000
4行目: ユーザー管理システム | P002_基本設計 | ベンダーA | 8 | 240,000
```

**3-6) ファイルを保存**
```python
wb.save(output_path)
```

**注**: 
- テンプレートには既に1行目（メモ欄）、2行目（ヘッダー）、3行目以降（データエリア）のスタイルが適用されている
- 金額列（E列）と工数列（D列）には自動的にカンマ区切り書式が適用される
- 未使用シートは空のまま保持される（将来の拡張用）

---

## ステップ4：generate_matching_check() の詳細

### 目的
複数の報告用Excelファイルから突き合わせチェック用Excelを生成する。

### 処理ステップ

**4-1) 新しい Workbook を作成**
```python
wb = Workbook()
wb.remove(wb.active)
ws = wb.create_sheet(title="突き合わせ")
```

**4-2) ヘッダー行（2行目）を作成**
```python
ws["A2"] = "報告ファイル名"
ws["B2"] = "シート名"
ws["C2"] = "統合情報"
ws["D2"] = "フェーズ"
ws["E2"] = "ベンダー/社員"
ws["F2"] = "工数（人日）"
ws["G2"] = "金額"
```

**4-3) 複数ファイルからデータを読み込んで転記**
```python
for file_path, file_display_name in file_paths:
    if not file_path.exists():
        continue
    
    wb_temp = load_workbook(file_path)
    
    for sheet_name in wb_temp.sheetnames:
        ws_temp = wb_temp[sheet_name]
        merged_info = ws_temp["A1"].value if ws_temp["A1"].value else "N/A"
        
        # 3行目以降のデータを転記
        for row in ws_temp.iter_rows(min_row=3, max_row=ws_temp.max_row,
                                     min_col=1, max_col=5, values_only=True):
            if row[0] is None:
                break
            
            ws[f"A{current_row}"] = file_display_name
            ws[f"B{current_row}"] = sheet_name
            ws[f"C{current_row}"] = merged_info
            ws[f"D{current_row}"] = row[1]  # フェーズ
            ws[f"E{current_row}"] = row[2]  # ベンダー/社員
            ws[f"F{current_row}"] = row[3]  # 工数
            ws[f"G{current_row}"] = row[4]  # 金額
            current_row += 1
```

**4-4) スタイル適用とファイル保存**
```python
# ヘッダー行にスタイルを適用
for col in ["A", "B", "C", "D", "E", "F", "G"]:
    cell = ws[f"{col}2"]
    cell.fill = header_fill
    cell.font = header_font
    # ...

# ファイルを保存
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = output_dir / f"matching_check_{timestamp}.xlsx"
wb.save(output_file)
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
アプリID | ベンダー/社員 | フェーズ名 | 工数 | 発注金額 | 社員コスト | 出力シート指定
APP001 | ベンダーA | P001_要件定義 | 5 | 150000 | | 詳細_1
APP001 | 社員 | P001_要件定義 | 2 | | 200000 | 詳細_1
APP001 | ベンダーA | P002_基本設計 | 8 | 240000 | | 詳細_1
APP001 | ベンダーB | P004_実装 | 15 | 450000 | | 詳細_2
APP002 | ベンダーC | P001_要件定義 | 10 | 300000 | | 詳細_3
```

### 処理中のデータ

**ステップ1後**:
```python
apps = {"APP001": "ユーザー管理", "APP002": "決済システム"}
phases = {"P001_要件定義": "資産", "P002_基本設計": "資産", "P004_実装": "費用"}
work_data = [
    ("APP001", "ベンダーA", "P001_要件定義", 5, 150000, None, "詳細_1"),
    ("APP001", "社員", "P001_要件定義", 2, None, 200000, "詳細_1"),
    ("APP001", "ベンダーA", "P002_基本設計", 8, 240000, None, "詳細_1"),
    ("APP001", "ベンダーB", "P004_実装", 15, 450000, None, "詳細_2"),
    ("APP002", "ベンダーC", "P001_要件定義", 10, 300000, None, "詳細_3")
]
```

**ステップ2後**:
```python
asset_data = {
    ("APP001", "ベンダーA", "詳細_1"): {
        "P001_要件定義": {"hours": 5, "amount": 150000},
        "P002_基本設計": {"hours": 8, "amount": 240000}
    },
    ("APP001", "社員", "詳細_1"): {
        "P001_要件定義": {"hours": 2, "amount": 200000}
    },
    ("APP002", "ベンダーC", "詳細_3"): {
        "P001_要件定義": {"hours": 10, "amount": 300000}
    }
}

expense_data = {
    ("APP001", "ベンダーB", "詳細_2"): {
        "P004_実装": {"hours": 15, "amount": 450000}
    }
}
```

### 出力例

**houkoku_shisan.xlsx**
```
[詳細_1 シート]
1行目: [メモ] | （未使用）
2行目: アプリ名称 | フェーズ | ベンダー/社員 | 工数（人日） | 金額
3行目: ユーザー管理システム | P001_要件定義 | ベンダーA | 5 | 150,000
4行目: ユーザー管理システム | P002_基本設計 | ベンダーA | 8 | 240,000
5行目: ユーザー管理システム | P001_要件定義 | 社員 | 2 | 200,000

[詳細_2 シート]
1行目: [メモ] | （未使用）
2行目: アプリ名称 | フェーズ | ベンダー/社員 | 工数（人日） | 金額
3行目: ユーザー管理システム | P001_要件定義 | ベンダーA | 5 | 150,000
...

[詳細_3 シート]
1行目: [メモ] | （未使用）
2行目: アプリ名称 | フェーズ | ベンダー/社員 | 工数（人日） | 金額
3行目: 決済システム | P001_要件定義 | ベンダーC | 10 | 300,000

[詳細_4～15 シート]
（未使用、空のまま）
```

**houkoku_hiyo.xlsx**
```
[詳細_1 シート]
1行目: [メモ] | （未使用）
2行目: アプリ名称 | フェーズ | ベンダー/社員 | 工数（人日） | 金額
3行目: ユーザー管理システム | P004_実装 | ベンダーB | 15 | 450,000

[詳細_2～15 シート]
（未使用、空のまま）
```

**matching_check_20260530_231904.xlsx**
```
[突き合わせ シート]
2行目: 報告ファイル名 | シート名 | 統合情報 | フェーズ | ベンダー/社員 | 工数（人日） | 金額
3行目: houkoku_shisan.xlsx | 詳細_1 | [メモ] | P001_要件定義 | ベンダーA | 5 | 150,000
4行目: houkoku_shisan.xlsx | 詳細_1 | [メモ] | P002_基本設計 | ベンダーA | 8 | 240,000
...
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
