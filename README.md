# Excel_To_Word-by-Python
This program retrieves information from an Excel file and outputs it to a Word file with the information appended to it. I have commented it out in the code and hope it is helpful.
---
Excelファイルからデータを読み込み、特定の形式のWordドキュメントを生成するプロセスを行います。以下は、各セルのコードについての詳細な説明です。

### 1. 必要なパッケージのインストール

```python
!pip install openpyxl
!pip install pandas
```
- `openpyxl`: Excelファイルを読み書きするためのライブラリ。
- `pandas`: データフレーム操作のためのライブラリ。

### 2. パッケージのインポート

```python
import openpyxl
import pandas as pd
import glob
```
- `openpyxl`と`pandas`をインポートして、Excelファイル操作とデータフレーム操作を行います。
- `glob`はファイルパスのパターンマッチングを行うためのモジュールです。

### 3. ファイルパスとシート名の設定

```python
import_file_path = 'your_file_location'
excel_sheet_name = '全部員'
export_file_path = 'your_file_location'
```
- ここではExcelファイルのパスとシート名を設定しています。

### 4. Excelファイルの読み込み

```python
df_order = pd.read_excel(import_file_path, sheet_name = excel_sheet_name)
```
- 指定されたExcelファイルのシートをデータフレームに読み込みます。

### 5. データフレームの表示

```python
df_order
```
- データフレームの内容を表示します。

### 6. ユニークな値の抽出

```python
people_name = df_order['氏名'].unique()
people_location = df_order['住所'].unique()
people_birthday = df_order['生年月日'].unique()
```
- データフレームから「氏名」、「住所」、「生年月日」のユニークな値を抽出します。

### 7. ユニークな値の表示

```python
people_name
people_location
people_birthday
```
- 抽出したユニークな値を表示します。

### 8. 各値の出力

```python
for person_name in people_name:
    print(person_name)
for person_location in people_location:
    print(person_location)
for person_birthday in people_birthday:
    print(person_birthday)
```
- 各ユニークな値を順に出力します。

### 9. 日付のパースと分解

```python
from datetime import datetime
for person_birthday in people_birthday:
    date_string = str(person_birthday)
    dates = [datetime.strptime(date_string, '%Y-%m-%d %H:%M:%S')]
    years = [date.year for date in dates]
    months = [date.month for date in dates]
    days = [date.day for date in dates]
    print("Years:", years)
    print("Months:", months)
    print("Days:", days)
```
- 生年月日をパースして年、月、日に分解し、表示します。

### 10. `python-docx`のインストール

```python
pip install python-docx
```
- `python-docx`: Wordドキュメントを操作するためのライブラリ。

### 11. `Document`のインポート

```python
from docx import Document
```
- `Document`クラスをインポートしてWordドキュメントを操作します。

### 12. セルの枠線を設定する関数

```python
from docx.oxml import OxmlElement
from docx.shared import RGBColor
from docx.oxml.ns import qn

def set_cell_borders_black(cell):
    tcPr = cell._element.get_or_add_tcPr()
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '4')
    top.set(qn('w:color'), '000000')
    tcPr.append(top)
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single')
    left.set(qn('w:sz'), '4')
    left.set(qn('w:color'), '000000')
    tcPr.append(left)
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:color'), '000000')
    tcPr.append(bottom)
    right = OxmlElement('w:right')
    right.set(qn('w:val'), 'single')
    right.set(qn('w:sz'), '4')
    right.set(qn('w:color'), '000000')
    tcPr.append(right)
```
- セルの枠線を黒色に設定する関数を定義します。

### 13. Wordドキュメントの生成と保存

```python
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

for i in range(len(people_birthday)):
    doc = Document()
    chiji_name = "name"
    doc.add_paragraph('第38号の5様式\n\n')
    doc.add_paragraph('comment\n').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    doc.add_paragraph('　　石川県知事　'+chiji_name+'様\n')
    doc.add_paragraph('coment')
    table = doc.add_table(rows=9, cols=2)
    for row in table.rows:
        for cell in row.cells:
            set_cell_borders_black(cell)
    for row in table.rows:
        row.cells[0].width = Inches(1.4)
        row.cells[1].width = Inches(4.3)
    date_string = str(people_birthday[i])
    date = datetime.strptime(date_string, '%Y-%m-%d %H:%M:%S')
    years = date.year
    months = date.month
    days = date.day
    rows_contents = [
        ('利用ゴルフ場名', ''),
        ('利用年月日', '年　　　月　　　日'),
        ('住所', people_location[i]),
        ('区分', '□　メンバー　　□　ビジター'),
        ('氏名', people_name[i]),
        ('生年月日', '大・昭・平'+str(years-1988)+'年'+str(months)+'月'+str(days)+'日生'),
        ('非課税等適用区分', '□70歳以上　□18歳未満　□障害者等\n□教育活動等　□国民スポーツ大会　□スポーツマスターズ等'),
        ('証明書の種類', '□運転免許証　□学生証　□職員証　□パスポート\n□障害者手帳等　□学校長の証明　□教育委員会の証明\n□その他(　　　　　　　　)'),
        ('備考','')
    ]
    for row, (label, content) in zip(table.rows, rows_contents):
        row.cells[0].text = label
        row.cells[1].text = content
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(10)
    remarks = '''備考　1　該当する□の中にレ点を付けてください。
      2　70歳以上、18歳未満及び障害者等の方は、この申請書を、利用するゴルフ場が最初の
      利用である場合にゴルフ場に提出してください。また、受付の際に非課税利用に該当する
      ことを証明する証明書をゴルフ場に提示してください。
      3　教育活動等、国民スポーツ大会、スポーツマスターズ等の利用の場合は、利用の都度
      この申請書を提出してください。その際には、受付に非課税・課税免除利用に該当するこ
      とを証明する証明書をゴルフ場に提出してください。
      4　この申請書を提出しない場合、2又は3の証明書を提示又は提出しない場合は、非課
      税・課税免除の適用を受けられない場合があります。
'''
    doc.add_paragraph(remarks).alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    file_path = 'your_file_location'
    doc.save(file_path)
    file_path
```
- 各人の情報をもとにWordドキュメントを生成し、指定されたファイルパスに保存します。

### 14. フォントとサイズを設定

```python
from docx import Document
from docx.shared import Pt

for i in range(len(people_birthday)):
    doc = Document('your_file_location')
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(10)
            run.font.name = 'ＭＳ 明朝'


    doc.save('your_file_location')
```
- 保存されたWordドキュメントを開き、フォントとサイズを設定して再度保存します。

---

このコードは、指定されたExcelファイルからデータを読み込み、各人の情報をもとにWordドキュメントを生成して保存します。Wordドキュメントのフォントとサイズを設定するプロセスも含まれています。
