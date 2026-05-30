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
    ├── 報告_資産.xlsx            # 自動生成：資産に区分されたフェーズのシート
    └── 報告_費用.xlsx            # 自動生成：費用に区分されたフェーズのシート
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
           │    └─ マスタデータ読み込み
           │
           ├─► organize_data_by_division()
           │    └─ 工数を資産/費用に分類
           │
           └─► create_workbook()
                └─ 報告用Excel生成 ×2
           │
           ▼
┌──────────────────────────┐
│  報告_資産.xlsx           │
│  報告_費用.xlsx           │
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
- `output/報告_資産.xlsx` （資産に区分されたフェーズ）
- `output/報告_費用.xlsx` （費用に区分されたフェーズ）

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

| フェーズID | フェーズ名 | 区分 |
|-----------|----------|------|
| P001 | 要件定義 | 資産 |
| P002 | 基本設計 | 資産 |
| P003 | 詳細設計 | 資産 |
| P004 | 実装 | 費用 |
| P005 | 単体テスト | 費用 |
| ... | ... | ... |

**カスタマイズ方法**：
- 資産に区分したいフェーズの「区分」列を「資産」に設定
- 費用に区分したいフェーズの「区分」列を「費用」に設定
- 新しいフェーズを追加する場合、行を追加します。

### シート3：「工数」
各アプリのフェーズごとの工数（人日）を記入します。

| アプリID | フェーズID | 工数（人日） |
|---------|-----------|-----------|
| APP001 | P001 | 5 |
| APP001 | P002 | 8 |
| APP001 | P004 | 15 |
| ... | ... | ... |

**カスタマイズ方法**：
- アプリIDとフェーズIDの組み合わせで、実際の工数を数値で入力
- 工数がない場合は、この行を削除

---

## 出力ファイル（報告_資産.xlsx / 報告_費用.xlsx）

### 内容
- **アプリごとにシートが作成**されます
- 各シート内には、該当するフェーズと工数が記載されます
  - 資産ファイル → 資産に区分されたフェーズのみ記載
  - 費用ファイル → 費用に区分されたフェーズのみ記載

### 例：報告_資産.xlsx
- シート「ユーザー管理システム」：要件定義（5人日）、基本設計（8人日）
- シート「決済システム」：要件定義（10人日）、基本設計（12人日）、詳細設計（20人日）
- ...

---

## プログラムの重要な関数

### 1. `read_config()`
**役割**: 入力ファイル（config.xlsx）を読み込む

```python
apps, phases, work_data = read_config()
```

**戻り値**:
- `apps`: `{アプリID: アプリ名}` の辞書
- `phases`: `{フェーズID: (フェーズ名, 区分)}` の辞書
- `work_data`: `[(アプリID, フェーズID, 工数), ...]` のリスト

---

### 2. `organize_data_by_division(work_data, phases)`
**役割**: 工数データを資産/費用に分類

```python
asset_data, expense_data = organize_data_by_division(work_data, phases)
```

**戻り値**:
- `asset_data`: `{アプリID: {フェーズID: 工数, ...}, ...}` (資産に区分)
- `expense_data`: `{アプリID: {フェーズID: 工数, ...}, ...}` (費用に区分)

---

### 3. `create_workbook(apps, phases, data_by_app, file_name)`
**役割**: 報告用Excelファイルを生成

```python
create_workbook(apps, phases, asset_data, "報告_資産.xlsx")
create_workbook(apps, phases, expense_data, "報告_費用.xlsx")
```

**処理内容**:
1. Workbookオブジェクトを作成
2. アプリごとにシートを追加
3. フェーズと工数を記入
4. セルを装飾（背景色、フォント、枠線）
5. ファイルを保存

---

### 4. `format_workbook(ws)`
**役割**: ワークシートをフォーマット（色、フォント、枠線など）

```python
format_workbook(ws)
```

**装飾内容**:
- ヘッダー行：青色背景、白い太字
- データ行：枠線、右揃え（工数列）

---

## 修正・カスタマイズする際のポイント

### 🔧 出力ファイルのレイアウトを変更したい

`create_workbook()` 関数内で、シートにデータを書き込む部分（ここ）を修正：

```python
# フェーズと工数を記入
row = 2
for phase_id, hours in sorted(data_by_app[app_id].items()):
    phase_name = phases[phase_id][0]
    ws[f"A{row}"] = phase_name
    ws[f"B{row}"] = hours
    row += 1
```

例：フェーズIDも表示したい場合：
```python
ws[f"C{row}"] = phase_id  # フェーズIDを追加
```

### 🎨 色やフォントを変更したい

`format_workbook()` 関数内のスタイル定義を修正：

```python
# ヘッダーの色を変更（現在は青色 4472C4）
header_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")  # 赤色に変更

# フォントサイズを変更（現在は11pt）
header_font = Font(bold=True, color="FFFFFF", size=14)  # 14ptに変更
```

### 📊 資産と費用の判定ルールを変更したい

`organize_data_by_division()` 関数のこの部分を修正：

```python
if division == "資産":
    asset_data[app_id][phase_id] = hours
elif division == "費用":
    expense_data[app_id][phase_id] = hours
```

### ⚠️ エラーハンドリングを強化したい

`read_config()` 関数に検証ロジックを追加：

```python
for app_id, phase_id, hours in work_data:
    # アプリIDが存在するか確認
    if app_id not in apps:
        raise ValueError(f"不正なアプリID: {app_id}")
    # フェーズIDが存在するか確認
    if phase_id not in phases:
        raise ValueError(f"不正なフェーズID: {phase_id}")
    # 工数が正の数か確認
    if not isinstance(hours, (int, float)) or hours < 0:
        raise ValueError(f"工数が不正です: {hours}")
```

---

## よくある質問

**Q: アプリやフェーズを削除したい場合は？**

A: `config.xlsx` の該当する行を削除してから、プログラムを再実行してください。

**Q: 複数の工数ファイルを統合したい場合は？**

A: `config.xlsx` の「工数」シートに複数のファイルのデータを統合（コピー＆ペースト）してから、プログラムを実行してください。

**Q: 出力ファイルのシート名を長くしたい場合は？**

A: Excelのシート名は31文字までの制限があります。プログラム内で以下の部分を修正：

```python
ws = wb.create_sheet(title=app_name[:31])  # [:31] を削除するか、数字を変更
```

**Q: 工数の小計や合計を自動計算したい場合は？**

A: `create_workbook()` 関数内で、小計行を追加：

```python
# データ行の後に合計行を追加
total_hours = sum(data_by_app[app_id].values())
ws[f"A{row}"] = "合計"
ws[f"B{row}"] = total_hours
```

---

## 学習のポイント

このプログラムで学べることは：

1. **ファイルI/O**: Excelファイルの読み書き（openpyxl）
2. **データ処理**: 辞書やリストを使った効率的なデータ処理
3. **関数設計**: 責務を明確に分けた関数設計
4. **エラーハンドリング**: 予期しないエラーへの対応
5. **スタイリング**: セルの装飾や書式設定

---

## トラブルシューティング

### エラー: `FileNotFoundError: 設定ファイルが見つかりません`

→ `input/config.xlsx` が存在するか確認してください。
→ `create_sample_config.py` を実行して、サンプルファイルを生成してください。

### エラー: `KeyError` が発生した

→ config.xlsx の「工数」シートで、アプリIDやフェーズIDが「アプリ」「フェーズ」シートに存在しているか確認してください。

### 出力ファイルが文字化けしている

→ Excelで出力ファイルを開き、ファイル→オプション→詳細→文字エンコーディングを確認してください。

---

## ライセンス

このプログラムはPythonの学習用です。自由に修正・改変して、業務に活用してください。
