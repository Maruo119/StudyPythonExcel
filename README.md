# Pythonを使ったExcel工数管理自動化

このプロジェクトは、大規模開発案件のコスト管理をPythonで自動化する学習教材です。

## プロジェクト構成

```
D:\StudyPythonExcel/
├── excel_automation.py          # メインプログラム（自動生成の処理）
├── create_sample_config.py      # サンプル設定Excelを生成するスクリプト
├── create_template.py           # テンプレートExcelを生成するスクリプト
├── PROGRAM_FLOW.md              # プログラムの処理フロー詳解
├── README.md                    # このファイル
├── template/
│   └── houkoku_template.xlsx    # テンプレート（15シート事前用意）【Git管理対象】
├── input/
│   └── config.xlsx              # ユーザーが入力する設定Excel（マスタデータ）【Git管理対象】
└── output/
    ├── houkoku_shisan.xlsx      # 自動生成：資産に区分されたフェーズのシート
    ├── houkoku_hiyo.xlsx        # 自動生成：費用に区分されたフェーズのシート
    └── matching_check_*.xlsx    # 自動生成：突き合わせチェック用ファイル
```

---

## 実装の全体像

```
┌──────────────────────────┐
│ テンプレート初期化       │
│ create_template.py       │
│ (15シート用意)           │
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│ config.xlsx を作成       │
│ create_sample_config.py  │
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│ excel_automation.py      │
│ を実行                   │
└──────────┬───────────────┘
           │
           ├─► read_config()
           │    └─ マスタデータ・工数を読み込み
           │
           ├─► organize_data_by_division()
           │    └─ 工数を資産/費用に分類
           │
           ├─► create_workbook()（資産/費用）
           │    ├─ テンプレートをコピー
           │    └─ 15シートに対してデータを上書き
           │
           └─► create_matching_check_workbook()
                └─ 報告Excel から突き合わせチェック生成
           │
           ▼
┌──────────────────────────┐
│ houkoku_shisan.xlsx      │
│ houkoku_hiyo.xlsx        │
│ matching_check_*.xlsx    │
│ を出力                   │
└──────────────────────────┘
```

---

## 使用方法

### ステップ1: テンプレートを生成（初回のみ）

```powershell
python create_template.py
```

このコマンドで `template/houkoku_template.xlsx` が生成されます（15シート事前用意）。

### ステップ2: サンプル設定ファイルを生成（初回のみ）

```powershell
python create_sample_config.py
```

このコマンドで `input/config.xlsx` が生成されます。
（以降はこのファイルを編集して使用）

### ステップ3: メインプログラムを実行（通常モード）

```powershell
python excel_automation.py
```

このコマンドで自動的に以下が生成されます：
- `output/houkoku_shisan.xlsx` （資産に区分されたフェーズ）
- `output/houkoku_hiyo.xlsx` （費用に区分されたフェーズ）
- `output/matching_check_[タイムスタンプ].xlsx` （突き合わせチェック）

### ステップ4-1: 任意のタイミングで突き合わせチェックを生成（チェックのみモード）

```powershell
python excel_automation.py --check-only
```

既存の報告Excel（`output/houkoku_shisan.xlsx`、`houkoku_hiyo.xlsx`）から突き合わせチェックを生成します。

### ステップ4-2: カスタムパスを指定してチェックを生成

```powershell
python excel_automation.py --check-only --asset-file C:\path\to\asset.xlsx --expense-file C:\path\to\expense.xlsx
```

---

## config.xlsx（入力ファイル）の構成

### シート1：「アプリ」
アプリケーションのマスタデータを定義します。

| アプリID | アプリ名 |
|---------|---------|
| APP001 | ユーザー管理システム |
| APP002 | 決済システム |
| ... | ... |

**カスタマイズ方法**：
- 新しいアプリを追加したい場合、行を追加してアプリIDとアプリ名を入力します。

### シート2：「フェーズ」
プロジェクトフェーズと、それが「資産」か「費用」かを定義します。

| フェーズ名 | 区分 |
|-----------|------|
| P001_要件定義 | 資産 |
| P002_基本設計 | 資産 |
| P003_詳細設計 | 資産 |
| P004_実装 | 費用 |
| P005_単体テスト | 費用 |
| P006_統合テスト | 費用 |
| P007_本番環境構築 | 費用 |

**カスタマイズ方法**：
- フェーズ名は **フェーズID + アンダースコア + 名称** の形式で統一してください
- 資産に区分したいフェーズの「区分」列を「資産」に設定
- 費用に区分したいフェーズの「区分」列を「費用」に設定
- 新しいフェーズを追加する場合、行を追加します。

### シート3：「社員単価」
社員1人あたりの日当単価を定義します。

| 単価ラベル | 単価 |
|----------|------|
| 社員単価 | 100000 |

**カスタマイズ方法**：
- 社員単価を変更する場合、B2セルの値を更新してください

### シート4：「工数」
各アプリのベンダー/社員ごと、フェーズごとの工数（人日）と発注金額を記入します。

| アプリID | ベンダー/社員 | フェーズ名 | 工数（人日） | 発注金額 | 社員コスト | 出力シート指定 |
|---------|-------------|----------|-----------|--------|----------|------------|
| APP001 | ベンダーA | P001_要件定義 | 5 | 150000 | | 詳細_1 |
| APP001 | 社員 | P001_要件定義 | 2 | | 200000 | 詳細_1 |
| APP001 | ベンダーA | P002_基本設計 | 8 | 240000 | | 詳細_1 |
| APP001 | ベンダーB | P004_実装 | 15 | 450000 | | 詳細_2 |
| ... | ... | ... | ... | ... | ... | ... |

**カスタマイズ方法**：
- **ベンダーの場合**: ベンダー/社員列に「ベンダー名」、発注金額を入力。社員コストは空
- **社員の場合**: ベンダー/社員列に「社員」と入力。発注金額は空。社員コストは工数 × 社員単価を入力
- **出力シート指定**: データを配置するシート名を指定（詳細_1～詳細_15）。未指定の場合は「詳細_1」
- 工数がない場合は、この行を削除
- 同じアプリでも複数のベンダーに依頼する場合は、ベンダー名を変えて複数行追加

---

## 出力ファイル（houkoku_shisan.xlsx / houkoku_hiyo.xlsx）

### 内容
- **15シート（詳細_1～詳細_15）**: テンプレートに事前用意。使用するシートだけにデータを上書き
- **1行目（メモ欄）**: アプリ名_ベンダー/社員
- **2行目（ヘッダー）**: アプリ名称、フェーズ、ベンダー/社員、工数（人日）、金額
- **3行目以降（データ）**: 実際の工数・金額情報（**金額と工数はカンマ区切りで表示**）

### 例：houkoku_shisan.xlsx
```
[詳細_1 シート]
行 1: [メモ]
行 2: アプリ名称 | フェーズ | ベンダー/社員 | 工数（人日） | 金額
行 3: ユーザー管理システム | P001_要件定義 | ベンダーA | 5 | 150,000
行 4: ユーザー管理システム | P002_基本設計 | ベンダーA | 8 | 240,000
行 5: ユーザー管理システム | P001_要件定義 | 社員 | 2 | 200,000

[詳細_2 シート]
行 1: [メモ]
行 2: アプリ名称 | フェーズ | ベンダー/社員 | 工数（人日） | 金額
行 3: ユーザー管理システム | P004_実装 | ベンダーB | 15 | 450,000

[詳細_3～15 シート]
（未使用、空のまま保持）
```

### 金額・工数の表示形式
- **工数（D列）**: カンマ区切り（例：5,125）
- **金額（E列）**: カンマ区切り（例：150,000）

---

## プログラムの重要な関数

### 1. `read_config()`
**役割**: 入力ファイル（config.xlsx）を読み込む

```python
apps, phases, work_data = read_config()
```

**戻り値**:
- `apps`: `{アプリID: アプリ名}` の辞書
- `phases`: `{フェーズ名: 区分}` の辞書（フェーズ名で管理）
- `work_data`: `[(アプリID, ベンダー名, フェーズ名, ベンダー工数, 発注金額, 社員工数), ...]` のリスト

---

### 2. `organize_data_by_division(work_data, phases)`
**役割**: 工数データを資産/費用に分類し、アプリ×ベンダー単位でグループ化

```python
asset_data, expense_data = organize_data_by_division(work_data, phases)
```

**戻り値**:
- `asset_data`: `{(アプリID, ベンダー名): {フェーズ名: {vendor_hours, amount, employee_hours}}, ...}`
- `expense_data`: 同じ構造の費用データ

---

### 3. `create_workbook(apps, data_by_group, file_name)`
**役割**: テンプレートをコピーして報告用Excelファイルを生成

```python
create_workbook(apps, asset_data, "houkoku_shisan.xlsx")
create_workbook(apps, expense_data, "houkoku_hiyo.xlsx")
```

**処理内容**:
1. テンプレート（`template/houkoku_template.xlsx`）をコピー
2. コピーしたファイルを開く
3. データを対応するシート（詳細_1, 詳細_2, ...）に上書き
4. 1行目にメモ欄（アプリ名_ベンダー/社員）を記入
5. 3行目以降にフェーズと工数を記入
6. ファイルを保存
   
**注**: テンプレートには既に以下が設定されている
- 2行目ヘッダー行（青色背景、白い太字）
- セルの装飾（背景色、枠線）
- 数値書式（金額・工数にカンマ区切り）

---

### 4. `generate_matching_check(file_paths, output_dir)`
**役割**: 複数の報告用Excelから突き合わせチェック用Excelを生成

```python
# 内部的に create_matching_check_workbook() から呼び出される
generate_matching_check(
    [(Path("output/houkoku_shisan.xlsx"), "houkoku_shisan.xlsx"),
     (Path("output/houkoku_hiyo.xlsx"), "houkoku_hiyo.xlsx")],
    Path("output")
)
```

**処理内容**:
1. 複数の報告Excelファイルを読み込み
2. 新しいWorkbookを作成（「突き合わせ」シート）
3. 各ファイルのデータを統合して転記
4. タイムスタンプ付きでファイルを保存

---

## 修正・カスタマイズする際のポイント

### 🔧 報告ファイルの名前を変更したい

`excel_automation.py` の main() 関数内で、create_workbook() の呼び出しを修正：

```python
# 現在
create_workbook(apps, asset_data, "houkoku_shisan.xlsx")
create_workbook(apps, expense_data, "houkoku_hiyo.xlsx")

# 変更例
create_workbook(apps, asset_data, "報告_資産.xlsx")
create_workbook(apps, expense_data, "報告_費用.xlsx")
```

### 🔧 テンプレートのスタイルを変更したい

`create_template.py` 関数内のスタイル定義を修正し、再度テンプレートを生成：

```python
# メモ欄の色を変更（現在は黄色 FFF2CC）
memo_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

# ヘッダーの色を変更（現在は青色 4472C4）
header_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")

# フォントサイズを変更（現在は11pt）
header_font = Font(bold=True, color="FFFFFF", size=14)
```

修正後、`python create_template.py` を実行してテンプレートを再生成してください。

### 🔧 出力シート数を変更したい

`create_template.py` の `create_template()` 関数内の「15シート作成」の部分を修正：

```python
# 現在（15シート）
for sheet_num in range(1, 16):

# 変更例（20シート）
for sheet_num in range(1, 21):
```

修正後、`python create_template.py` を実行してテンプレートを再生成してください。

### 🔧 出力ファイルの名前を変更したい

`excel_automation.py` の main() 関数内で、create_workbook() の呼び出しを修正：

```python
# 現在
create_workbook(apps, asset_data, "houkoku_shisan.xlsx")
create_workbook(apps, expense_data, "houkoku_hiyo.xlsx")

# 変更例
create_workbook(apps, asset_data, "報告_資産.xlsx")
create_workbook(apps, expense_data, "報告_費用.xlsx")
```

### 📊 複数の区分（資産/費用の他に「保守」など）を対応

`organize_data_by_division()` 関数で、新しい区分を追加：

```python
# 現在
if division == "資産":
    asset_data[key][phase_name] = phase_data
elif division == "費用":
    expense_data[key][phase_name] = phase_data

# 変更例（保守区分を追加）
maintenance_data = {}  # main() で定義
if division == "資産":
    asset_data[key][phase_name] = phase_data
elif division == "費用":
    expense_data[key][phase_name] = phase_data
elif division == "保守":
    maintenance_data[key][phase_name] = phase_data

# main() で報告ファイルを追加
create_workbook(apps, maintenance_data, "houkoku_hoshu.xlsx")
```

---

## よくある質問

**Q: アプリやフェーズを削除したい場合は？**

A: `config.xlsx` の該当する行を削除してから、プログラムを再実行してください。

**Q: 複数の工数ファイルを統合したい場合は？**

A: `config.xlsx` の「工数」シートに複数のファイルのデータを統合（コピー＆ペースト）してから、プログラムを実行してください。

**Q: 新しいベンダーを追加したい場合は？**

A: `config.xlsx` の「工数」シートに、アプリID、新しいベンダー名、フェーズ名、工数などを入力する行を追加してください。

**Q: 社員工数だけを入力したい場合は？**

A: `config.xlsx` の「工数」シートで、ベンダー/社員列に「社員」と入力し、社員コストを指定してください。

**Q: 出力シートを指定したい場合は？**

A: `config.xlsx` の「工数」シートの「出力シート指定」列に、シート名（詳細_1～詳細_15）を入力してください。

**Q: 突き合わせチェックを再度生成したい場合は？**

A: 以下のコマンドを実行してください：
```powershell
python excel_automation.py --check-only
```

**Q: テンプレートの見た目を変更したい場合は？**

A: `create_template.py` のスタイル定義を修正してから、`python create_template.py` を実行してテンプレートを再生成してください。

---

## 学習のポイント

このプログラムで学べることは：

1. **ファイルI/O**: Excelファイルの読み書き（openpyxl）
2. **データ処理**: 辞書やリストを使った複雑なデータ処理
3. **関数設計**: 責務を明確に分けた関数設計
4. **キー値の工夫**: タプルをキーにした複雑なデータグループ化
5. **スタイリング**: セルの装飾や書式設定
6. **エラーハンドリング**: 予期しないエラーへの対応

---

## トラブルシューティング

### エラー: `FileNotFoundError: 設定ファイルが見つかりません`

→ `input/config.xlsx` が存在するか確認してください。
→ `create_sample_config.py` を実行して、サンプルファイルを生成してください。

### エラー: `KeyError` が発生した

→ config.xlsx の「工数」シートで、フェーズ名が「フェーズ」シートに存在しているか確認してください。

### 出力ファイルが文字化けしている

→ Excelで出力ファイルを開き、ファイル→オプション→詳細→文字エンコーディングを確認してください。

---

## ライセンス

このプログラムはPythonの学習用です。自由に修正・改変して、業務に活用してください。
