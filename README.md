# Excel to Word 変換ツール

このプロジェクトは、Excelファイルからデータを読み込み、各人の情報をもとにWordドキュメントを生成するツールです。具体的には、ゴルフ場利用税の非課税申請書を自動生成します。

## 機能

- Excelファイルから氏名、住所、生年月日のデータを読み込む
- 各人ごとにWordドキュメントを生成する
- 生成したドキュメントに適切なフォーマットを適用する

## 必要なライブラリ

このツールを使用するには、以下のライブラリが必要です：

- pandas: データ処理用
- openpyxl: Excelファイル操作用
- python-docx: Wordドキュメント操作用
- datetime: 日付処理用
- python-dotenv: 環境変数の管理用（オプション）

## インストール方法

1. 必要なライブラリをインストールします：

```bash
pip install pandas openpyxl python-docx python-dotenv
```

## 使い方

### 1. 環境変数の設定

すべての設定は環境変数から取得されます。以下の2つの方法で環境変数を設定できます：

#### 1.1 .envファイルを使用する方法（推奨）

プロジェクトのルートディレクトリに`.env`ファイルを作成し、以下のように設定を記述します：

```
# 設定セクション
INPUT_FILE_PATH=input.xlsx
OUTPUT_DIRECTORY=output
SHEET_NAME=全部員
GOVERNOR_NAME=知事名
GOLF_COURSE_NAME=石川カントリークラブ
USAGE_DATE=2023年10月15日
```

この方法を使用するには、python-dotenvライブラリが必要です：

```bash
pip install python-dotenv
```

#### 1.2 直接環境変数を設定する方法

環境変数を直接設定することもできます：

##### Linuxまたは macOS:

```bash
export INPUT_FILE_PATH="input.xlsx"
export OUTPUT_DIRECTORY="output"
export SHEET_NAME="全部員"
export GOVERNOR_NAME="知事名"
export GOLF_COURSE_NAME="石川カントリークラブ"
export USAGE_DATE="2023年10月15日"
python excel_to_word.py
```

##### Windows (コマンドプロンプト):

```cmd
set INPUT_FILE_PATH=input.xlsx
set OUTPUT_DIRECTORY=output
set SHEET_NAME=全部員
set GOVERNOR_NAME=知事名
set GOLF_COURSE_NAME=石川カントリークラブ
set USAGE_DATE=2023年10月15日
python excel_to_word.py
```

##### Windows (PowerShell):

```powershell
$env:INPUT_FILE_PATH = "input.xlsx"
$env:OUTPUT_DIRECTORY = "output"
$env:SHEET_NAME = "全部員"
$env:GOVERNOR_NAME = "知事名"
$env:GOLF_COURSE_NAME = "石川カントリークラブ"
$env:USAGE_DATE = "2023年10月15日"
python excel_to_word.py
```

### 2. 環境変数の説明

| 環境変数 | 説明 | デフォルト値 |
|----------|------|------------|
| `INPUT_FILE_PATH` | 入力Excelファイルのパス | `input.xlsx` |
| `OUTPUT_DIRECTORY` | 出力ディレクトリ | `output` |
| `SHEET_NAME` | データが含まれるシート名 | `全部員` |
| `GOVERNOR_NAME` | 知事の名前 | `知事名` |
| `GOLF_COURSE_NAME` | ゴルフ場名 | `""` (空文字) |
| `USAGE_DATE` | 利用年月日 | `年　　　月　　　日` |

環境変数を設定しない場合、上記のデフォルト値が使用されます。

### 3. Excelファイルの準備

入力Excelファイルには、少なくとも以下の列が必要です：

| 列名 | データ型 | 説明 | 形式 | 例 |
|------|----------|------|------|------|
| `氏名` | 文字列 | 申請者の氏名 | フルネーム | 山田太郎 |
| `住所` | 文字列 | 申請者の住所 | 都道府県から番地まで | 東京都千代田区霞が関1-1-1 |
| `生年月日` | 日付 | 申請者の生年月日 | YYYY-MM-DD形式 | 1990-01-01 |

以下は入力Excelファイルの例です：

| 氏名 | 住所 | 生年月日 |
|------|------|----------|
| 山田太郎 | 東京都千代田区霞が関1-1-1 | 1990-01-01 |
| 佐藤花子 | 大阪府大阪市中央区大手前2-2-2 | 1985-05-15 |
| 鈴木一郎 | 愛知県名古屋市中区三の丸3-3-3 | 1978-12-30 |

### 4. スクリプトの実行

```bash
python excel_to_word.py
```

### 5. 結果の確認

スクリプトが正常に実行されると、指定した出力ディレクトリに各人のWordドキュメントが生成されます。ファイル名は「連番_氏名.docx」の形式になります。

## コードの説明

### 主な機能

1. **環境のセットアップ**：出力ディレクトリが存在しない場合は作成します。
2. **Excelデータの読み込み**：指定されたExcelファイルからデータを読み込みます。
3. **ユニークなデータの抽出**：氏名、住所、生年月日のユニークな値を抽出します。
4. **Wordドキュメントの作成**：各人の情報をもとにWordドキュメントを生成します。
5. **フォント設定の適用**：生成したドキュメントに適切なフォント設定を適用します。

### 関数の説明

- `setup_environment()`: 出力ディレクトリが存在しない場合は作成します。
- `load_excel_data(file_path, sheet_name)`: Excelファイルからデータを読み込みます。
- `extract_unique_data(df)`: データフレームから氏名、住所、生年月日のユニークな値を抽出します。
- `parse_date(date_string)`: 日付文字列をパースして年、月、日に分解します。
- `set_cell_borders_black(cell)`: セルの枠線を黒色に設定します。
- `apply_font_settings(file_path)`: ドキュメントのフォント設定を適用します。
- `create_word_document(name, address, birthday, index, governor_name)`: Wordドキュメントを作成します。
- `main()`: メイン処理を実行します。

## 改良点

元のJupyterノートブックから以下の改良を行いました：

1. **コードの構造化**：機能ごとに関数に分割し、メイン処理を`main()`関数にまとめました。
2. **エラー処理の追加**：各関数にtry-except文を追加し、エラーが発生した場合も適切に処理できるようにしました。
3. **ハードコードされた値の削除**：設定値を変数として定義し、簡単に変更できるようにしました。
4. **ドキュメントの追加**：各関数に日本語のドキュメントを追加し、コードの理解を助けるようにしました。
5. **ファイル名の改善**：出力ファイル名に連番を追加し、整理しやすくしました。

## 注意事項

- 入力Excelファイルには、必要な列（氏名、住所、生年月日）が含まれている必要があります。
- 生年月日は、YYYY-MM-DD形式である必要があります。
- 出力ディレクトリが存在しない場合は自動的に作成されます。
- 同名のファイルが既に存在する場合は上書きされます。

## トラブルシューティング

### エラー：「必要なカラムがExcelファイルに見つかりません」

入力Excelファイルに必要な列（氏名、住所、生年月日）が含まれていることを確認してください。

### エラー：「日付のパースに失敗しました」

生年月日の形式が正しいことを確認してください。サポートされている形式は「YYYY-MM-DD」または「YYYY-MM-DD HH:MM:SS」です。

### エラー：「ドキュメントの作成中に問題が発生しました」

出力ディレクトリに書き込み権限があることを確認してください。また、同名のファイルが開かれていないことを確認してください。
