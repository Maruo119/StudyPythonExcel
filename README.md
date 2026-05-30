# Pythonを使ったExcel工数管理自動化

このプロジェクトは、大規模開発案件のコスト管理をPythonで自動化する学習教材です。

## プロジェクト構成

```
D:\StudyPythonExcel/
├── excel_automation.py          # メインプログラム（自動生成の処理）
├── create_sample_config.py      # サンプル設定Excelを生成するスクリプト
├── README.md                    # このファイル
├── input/
│   └── config.xlsx              # ユーザーが入力する設定Excel（マスタデータ）
└── output/
    ├── houkoku_shisan.xlsx      # 自動生成：資産に区分されたフェーズのシート
    └── houkoku_hiyo.xlsx        # 自動生成：費用に区分されたフェーズのシート
```

---

## 実装の全体像

```
┌──────────────────────────┐
│   config.xlsx を作成     │
│ （アプリ、フェーズ、工数）│
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
           └─► create_workbook()
                └─ 報告用Excel生成 ×2
           │
           ▼
┌──────────────────────────┐
│  houkoku_shisan.xlsx     │
│  houkoku_hiyo.xlsx       │
│ を出力                   │
└──────────────────────────┘
```

---

## 使用方法

### ステップ1: サンプルファイルを生成

```powershell
python create_sample_config.py
```

このコマンドで `input/config.xlsx` が生成されます。
（初回のみ必要。以降はこのファイルを編集して使用）

### ステップ2: メインプログラムを実行

```powershell
python excel_automation.py
```

このコマンドで自動的に以下が生成されます：
- `output/houkoku_shisan.xlsx` （資産に区分されたフェーズ）
- `output/houkoku_hiyo.xlsx` （費用に区分されたフェーズ）

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

### シート3：「工数」
各アプリのベンダーごと、フェーズごとの工数（人日）と発注金額を記入します。

| アプリID | ベンダー名 | フェーズ名 | ベンダー工数（人日） | 発注金額 | 社員工数（人日） |
|---------|-----------|----------|-------------------|--------|---------------|
| APP001 | ベンダーA | P001_要件定義 | 5 | 150000 | 2 |
| APP001 | ベンダーA | P002_基本設計 | 8 | 240000 | 3 |
| APP001 | ベンダーB | P004_実装 | 15 | 450000 | 1 |
| ... | ... | ... | ... | ... | ... |

**カスタマイズ方法**：
- アプリID、ベンダー名、フェーズ名の組み合わせで、ベンダー工数・発注金額・社員工数を入力
- 工数がない場合は、この行を削除
- 同じアプリでも複数のベンダーに依頼する場合は、ベンダー名を変えて複数行追加

---

## 出力ファイル（houkoku_shisan.xlsx / houkoku_hiyo.xlsx）

### 内容
- **複数シート（詳細_1, 詳細_2, ...）**: アプリ×ベンダーの組み合わせごとにシートが作成
- **1行目（メモ欄）**: アプリ名、ベンダー名、社員工数の合計
- **2行目（ヘッダー）**: フェーズ名、ベンダー工数、発注金額、社員工数
- **3行目以降（データ）**: 実際の工数・金額情報

### 例：houkoku_shisan.xlsx
```
[詳細_1 シート]
行 1: ユーザー管理システム | ベンダーA | 社員工数 | 5
行 2: フェーズ名 | ベンダー工数（人日） | 発注金額 | 社員工数（人日）
行 3: P001_要件定義 | 5 | 150000 | 2
行 4: P002_基本設計 | 8 | 240000 | 3

[詳細_2 シート]
行 1: ユーザー管理システム | ベンダーB | 社員工数 | 3
行 2: フェーズ名 | ベンダー工数（人日） | 発注金額 | 社員工数（人日）
行 3: P004_実装 | 15 | 450000 | 1
行 4: P005_単体テスト | 10 | 300000 | 2

[詳細_3 シート]
行 1: 決済システム | ベンダーC | 社員工数 | 12
...
```

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

### 3. `create_workbook(apps, data_by_app_vendor, file_name)`
**役割**: 報告用Excelファイルを生成

```python
create_workbook(apps, asset_data, "houkoku_shisan.xlsx")
create_workbook(apps, expense_data, "houkoku_hiyo.xlsx")
```

**処理内容**:
1. Workbookオブジェクトを作成
2. アプリ×ベンダーごとにシートを追加（シート名は詳細_1, 詳細_2, ...）
3. 1行目にメモ欄（アプリ名、ベンダー名、社員工数の合計）を記入
4. 2行目にヘッダー行を記入
5. 3行目以降にフェーズと工数を記入
6. セルを装飾（背景色、フォント、枠線）
7. ファイルを保存

---

### 4. `format_workbook(ws)`
**役割**: ワークシートをフォーマット（色、フォント、枠線など）

```python
format_workbook(ws)
```

**装飾内容**:
- 1行目（メモ欄）：黄色背景
- 2行目（ヘッダー行）：青色背景、白い太字
- データ行：枠線、右揃え（数値）/左揃え（テキスト）

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

### 🔧 出力ファイルのレイアウトを変更したい

`create_workbook()` 関数内で、1行目～3行目以降の構成を修正：

```python
# 1行目の内容を変更
ws["A1"] = app_name
ws["B1"] = vendor_name
# 必要な情報を追加

# 2行目のヘッダーを追加/削除
ws["A2"] = "フェーズ名"
ws["B2"] = "ベンダー工数（人日）"
# カラムを追加したい場合は E2, F2... に追加

# 3行目以降のデータ行も対応
ws[f"E{row}"] = phase_data["新しい項目"]
```

### 🎨 色やフォントを変更したい

`format_workbook()` 関数内のスタイル定義を修正：

```python
# メモ欄の色を変更（現在は黄色 FFF2CC）
memo_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

# ヘッダーの色を変更（現在は青色 4472C4）
header_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")

# フォントサイズを変更（現在は11pt）
header_font = Font(bold=True, color="FFFFFF", size=14)
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

A: `config.xlsx` の「工数」シートで、ベンダー工数や発注金額を 0 または空白にしてください。

**Q: 出力ファイルのシート数を制限したい場合は？**

A: `create_workbook()` 関数内で、sheet_num が 20 に達したらループを抜ける処理を追加してください。

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
