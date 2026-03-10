# AI 時代 Excel 教學課件 — 一天完整課程
# 以下課程生成大量範例，讓學員可以跟著實際操作
> **適用對象：** Excel 新手 / 想用 AI 取代繁瑣操作的工作者
> **課程時長：** 一天（約 6-7 小時，含休息）
> **前置需求：** 有基本電腦操作能力，無需程式基礎
>
> **課程學習目標：**
> - ✅ 能用 ChatGPT 生成各種 Excel 公式與 Python 腳本
> - ✅ 能在 Google Colab 執行 Python 自動化處理 Excel
> - ✅ 能在 Colab 內建 AI 面板用自然語言生成與修改資料處理程式
> - ✅ 能用 Gemini CLI 以自然語言操作 Excel 檔案
> - ✅ 能整合多種 AI 工具完成月報自動化流程

---

## 課程時間表

| 時段 | 主題 | 時長 |
|------|------|------|
| 09:00-10:30 | **單元一：ChatGPT 操作 Excel** | 90 min |
| 10:30-10:45 | 休息 | 15 min |
| 10:45-12:15 | **單元二：Google Colab + AI（含 Colab 內建 AI）自動操作 Excel** | 90 min |
| 12:15-13:30 | 午餐 | 75 min |
| 13:30-15:00 | **單元三：Gemini CLI 操作 Excel** | 90 min |
| 15:00-15:15 | 休息 | 15 min |
| 15:15-16:30 | **單元四：綜合實戰演練** | 75 min |
| 16:30-17:00 | Q&A + 總結 | 30 min |

---

# 單元一：ChatGPT 操作 Excel（09:00-10:30）

> **本單元學習目標：**
> - 理解 AI 輔助 Excel 的核心優勢
> - 能用 ChatGPT 生成 VLOOKUP、SUMIFS、IF 等常見公式
> - 能描述需求讓 ChatGPT 生成可執行的 Python 腳本
>
> **難度：** ⭐☆☆（入門）

## 1.1 觀念建立：為什麼用 AI 操作 Excel？

傳統方式：手動輸入公式 → 查函數語法 → 除錯 → 反覆修改
AI 方式：**用自然語言描述需求** → AI 生成公式/Python 腳本 → 直接套用

### AI 輔助的三大優勢

```
┌─────────────────────────────────────────────────┐
│  1. 零記憶成本  — 不需要記住任何函數語法           │
│  2. 快速迭代    — 描述錯了再說一次，AI 立即修正   │
│  3. 能力倍增    — 過去要學 3 個月的技能，現在 1 天 │
└─────────────────────────────────────────────────┘
```

### 什麼時候用 AI，什麼時候自己做？

| 情境 | 建議做法 |
|------|---------|
| 臨時查詢一個公式 | 問 ChatGPT，貼上即用 |
| 每週重複的報表 | 讓 AI 寫 Python 腳本，一鍵執行 |
| 超大資料（10萬+ 筆）| 用 Colab + Python 處理 |
| 需要讀取本機多個檔案 | 用 Gemini CLI 或 Claude Code CLI |
| 簡單的幾個數字 | 直接在 Excel 手動做，更快 |

## 1.2 ChatGPT 生成 Excel 公式

### 範例 1：VLOOKUP 查找

**傳統做法：** 自己記住 `=VLOOKUP(查找值, 範圍, 欄號, FALSE)` 語法

**AI 做法：** 直接跟 ChatGPT 說：

```
我的 Excel 中：
- A 欄是「員工編號」
- B 欄是「姓名」
- C 欄是「部門」
- D 欄是「薪資」

我在另一個工作表的 A2 有一個員工編號，
想查出這個員工的薪資，請給我公式。
```

**ChatGPT 回覆：**
```excel
=VLOOKUP(A2, Sheet1!A:D, 4, FALSE)
```

### 範例 2：條件加總 SUMIFS

**對 ChatGPT 說：**
```
我的銷售表格：
- A 欄：日期
- B 欄：業務員姓名
- C 欄：產品類別
- D 欄：銷售金額

我想算出「王小明」在「2024年1月」賣出「電子產品」的總金額。
```

**ChatGPT 回覆：**
```excel
=SUMIFS(D:D, B:B, "王小明", C:C, "電子產品", A:A, ">="&DATE(2024,1,1), A:A, "<"&DATE(2024,2,1))
```

### 範例 3：巢狀 IF 判斷成績等級

**對 ChatGPT 說：**
```
B2 是學生分數（0-100），幫我寫公式：
- 90 以上：A
- 80-89：B
- 70-79：C
- 60-69：D
- 60 以下：F
```

**ChatGPT 回覆：**
```excel
=IF(B2>=90,"A",IF(B2>=80,"B",IF(B2>=70,"C",IF(B2>=60,"D","F"))))
```

### 範例 4：文字處理 — 從身分證字號提取生日

**對 ChatGPT 說：**
```
A2 是身分證字號（例如 A123456789），
我想從中提取出生日（民國年/月/日格式），請給我公式。
```

**ChatGPT 回覆：**
```excel
=MID(A2,2,2)&"/"&MID(A2,4,2)&"/"&MID(A2,6,2)
```

### 範例 5：動態排名

**對 ChatGPT 說：**
```
D 欄是各業務的銷售金額，我想在 E 欄顯示排名（由高到低），請給公式。
```

**ChatGPT 回覆：**
```excel
=RANK(D2, $D$2:$D$100, 0)
```

---

## 1.3 ChatGPT 生成 Python 腳本

### 範例 6：自動合併多個工作表

**對 ChatGPT 說：**
```
我的 Excel 活頁簿有 12 個工作表（1月, 2月, ... 12月），
每個工作表的格式一樣：A 欄是日期、B 欄是項目、C 欄是金額。
請幫我寫 Python，把 12 個工作表（1月, 2月, ... 12月）
合併成一個新的 Excel 檔案，工作表名稱叫「年度彙總」，
並新增「來源月份」欄位。
```

**ChatGPT 回覆的 Python：**
```python
import pandas as pd

# 讀取來源活頁簿（含 1月~12月工作表）
input_file = '每月資料.xlsx'
output_file = '年度彙總.xlsx'

# 取得所有工作表
all_sheets = pd.read_excel(input_file, sheet_name=None)

merged = []
for sheet_name, df in all_sheets.items():
    if sheet_name != '年度彙總' and not df.empty:
        temp = df.copy()
        temp['來源月份'] = sheet_name
        merged.append(temp)

if not merged:
    raise ValueError('沒有可合併的資料')

summary_df = pd.concat(merged, ignore_index=True)

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    summary_df.to_excel(writer, sheet_name='年度彙總', index=False)

print(f'合併完成！共 {len(summary_df)} 筆資料，已輸出 {output_file}')
```

**如何使用：**
1. 安裝套件：`pip install pandas openpyxl`
2. 把程式存成 `merge_sheets.py`
3. 將 `每月資料.xlsx` 放在同一資料夾
4. 執行：`python merge_sheets.py`

### 範例 7：自動寄送 Email（根據 Excel 名單）

**對 ChatGPT 說：**
```
我的 Excel：
- A 欄：收件人姓名
- B 欄：Email 地址
- C 欄：應付金額

請幫我寫 Python，讀取這份 Excel，
然後自動寄送催款通知信給每位收件人。
```

**ChatGPT 回覆的 Python：**
```python
import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

excel_file = '催款名單.xlsx'
smtp_host = 'smtp.gmail.com'
smtp_port = 587
sender_email = os.getenv('SENDER_EMAIL')
sender_password = os.getenv('SENDER_PASSWORD')

if not sender_email or not sender_password:
    raise ValueError('請先設定環境變數 SENDER_EMAIL 與 SENDER_PASSWORD')

df = pd.read_excel(excel_file)

with smtplib.SMTP(smtp_host, smtp_port) as server:
    server.starttls()
    server.login(sender_email, sender_password)

    for _, row in df.iterrows():
        receiver_name = str(row['收件人姓名']).strip()
        receiver_email = str(row['Email 地址']).strip()
        amount = int(row['應付金額'])

        subject = '付款提醒通知'
        body = (
            f'{receiver_name} 您好，\n\n'
            f'提醒您尚有應付款項 NT${amount:,}，敬請安排付款。\n\n'
            '如已付款請忽略此信，謝謝。'
        )

        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain', 'utf-8'))

        server.sendmail(sender_email, receiver_email, message.as_string())
        print(f'已寄送給：{receiver_email}')

print('全部郵件已寄送完成！')
```

---

## 1.4 ChatGPT + Excel 進階技巧

### 範例 8：請 ChatGPT 解釋別人寫的複雜公式

**對 ChatGPT 說：**
```
請幫我解釋這個 Excel 公式在做什麼：
=IFERROR(INDEX($B$2:$B$100,MATCH(1,(A2=$D$2:$D$100)*(B2=$E$2:$E$100),0)),"無資料")
```

### 範例 9：請 ChatGPT 優化現有公式

**對 ChatGPT 說：**
```
我用了很多巢狀 IF，很難維護：
=IF(A2="台北",100,IF(A2="台中",200,IF(A2="高雄",150,IF(A2="花蓮",300,0))))

有沒有更好的寫法？
```

**ChatGPT 建議用 SWITCH 或 XLOOKUP：**
```excel
=SWITCH(A2,"台北",100,"台中",200,"高雄",150,"花蓮",300,0)
```

### 範例 10：請 ChatGPT 生成範例數據

**對 ChatGPT 說：**
```
幫我生成 20 筆假的銷售數據，格式如下：
日期 | 業務員 | 客戶名稱 | 產品 | 數量 | 單價 | 金額

要有合理的數值，可以直接貼到 Excel。
```

---

## 1.5 課堂練習（15 分鐘）

> **練習題：** 打開 ChatGPT，用自然語言完成以下任務：
> 1. 請 ChatGPT 生成一個「員工考勤表」的公式，計算遲到次數
> 2. 請 ChatGPT 寫一個公式，把「2024/01/15」格式轉成「民國113年1月15日」
> 3. 請 ChatGPT 生成 30 筆假的「訂單資料」供後續練習使用

---

# 單元二：Google Colab + AI（含 Colab 內建 AI）自動操作 Excel（10:45-12:15）

> **本單元學習目標：**
> - 能在 Google Colab 環境中上傳並讀取 Excel 檔案
> - 能用 pandas 完成篩選、統計、樞紐分析等操作
> - 能在 Colab 內直接呼叫 AI（Gemini）協助分析與產生程式
> - 能生成視覺化圖表並下載格式化報表
>
> **難度：** ⭐⭐☆（初階）

## 2.1 什麼是 Google Colab？

- Google 免費提供的雲端 Python 環境
- **不需安裝任何軟體**，瀏覽器直接使用
- 已整合 AI 助理能力，可在筆記本中用自然語言請 AI 協助處理資料
- 網址：https://colab.research.google.com
- 用 Google 帳號即可登入

## 2.2 環境準備

### 第一步：開啟 Colab 並安裝套件

```python
# 在 Colab 的第一個儲存格執行
!pip install openpyxl pandas xlsxwriter
```

### 第二步：上傳 Excel 檔案到 Colab

```python
from google.colab import files

# 這會跳出檔案選擇視窗，選擇你的 Excel 檔案
uploaded = files.upload()
```

### 第三步：使用 Colab 內建 AI 直接協助操作

- 在 Colab 右上角開啟 AI（Gemini）面板（不同版本介面名稱可能是「Ask AI」或 Gemini 圖示）
- 直接輸入需求，讓 AI 幫你產生或修改目前筆記本程式碼
- 建議先描述資料欄位，再說清楚輸出結果格式

可直接貼上的提示詞範例：

```text
我已上傳 銷售資料.xlsx，欄位有：日期、業務員、產品類別、金額。
請幫我在下一個 Colab 儲存格產生 Python 程式，完成：
1. 讀取 Excel
2. 計算每位業務員總銷售額（由高到低）
3. 畫出長條圖
4. 匯出結果為 業務員銷售排行.xlsx
```

---

## 2.3 用 Python 讀取 Excel

### 範例 11：讀取並顯示 Excel 內容

```python
import pandas as pd

# 讀取 Excel 檔案
df = pd.read_excel('銷售資料.xlsx')

# 顯示前 10 筆
print("=== 前 10 筆資料 ===")
print(df.head(10))

# 顯示基本資訊
print(f"\n共 {len(df)} 筆資料，{len(df.columns)} 個欄位")
print(f"欄位名稱：{list(df.columns)}")
```

### 範例 12：讀取特定工作表

```python
# 讀取特定工作表
df = pd.read_excel('銷售資料.xlsx', sheet_name='1月')

# 讀取所有工作表
all_sheets = pd.read_excel('銷售資料.xlsx', sheet_name=None)
for sheet_name, data in all_sheets.items():
    print(f"工作表：{sheet_name}，共 {len(data)} 筆")
```

---

## 2.4 常見 Excel 操作的 Python 替代方案

### 範例 13：篩選資料（取代 Excel 篩選功能）

```python
import pandas as pd

df = pd.read_excel('銷售資料.xlsx')

# 篩選金額大於 10000 的訂單
big_orders = df[df['金額'] > 10000]
print("大額訂單：")
print(big_orders)

# 多條件篩選：業務員是王小明 且 產品是電子產品
filtered = df[(df['業務員'] == '王小明') & (df['產品類別'] == '電子產品')]
print("\n王小明的電子產品訂單：")
print(filtered)

# 篩選日期範圍
df['日期'] = pd.to_datetime(df['日期'])
jan_data = df[(df['日期'] >= '2024-01-01') & (df['日期'] <= '2024-01-31')]
print(f"\n一月份共 {len(jan_data)} 筆")
```

### 範例 14：樞紐分析表（取代 Excel Pivot Table）

```python
import pandas as pd

df = pd.read_excel('銷售資料.xlsx')

# 依業務員統計銷售金額（等同 Excel 樞紐分析表）
pivot = df.pivot_table(
    values='金額',
    index='業務員',
    columns='產品類別',
    aggfunc='sum',
    fill_value=0,
    margins=True  # 加上總計
)
print("=== 業務員 × 產品類別 銷售統計 ===")
print(pivot)
```

### 範例 15：VLOOKUP 的替代 — merge

```python
import pandas as pd

# 主表：訂單資料
orders = pd.read_excel('訂單.xlsx')
# 查詢表：產品價目表
products = pd.read_excel('產品.xlsx')

# 合併（等同 VLOOKUP）
result = orders.merge(products, on='產品編號', how='left')
print(result)
```

### 範例 16：條件新增欄位（取代 IF 公式）

```python
import pandas as pd

df = pd.read_excel('成績單.xlsx')

# 等同 Excel 的 IF 巢狀公式
def grade(score):
    if score >= 90: return 'A'
    elif score >= 80: return 'B'
    elif score >= 70: return 'C'
    elif score >= 60: return 'D'
    else: return 'F'

df['等第'] = df['分數'].apply(grade)
print(df)
```

### 範例 17：分組統計（取代 SUMIF / COUNTIF）

```python
import pandas as pd

df = pd.read_excel('銷售資料.xlsx')

# SUMIF 的替代：各部門銷售總額
dept_sum = df.groupby('部門')['金額'].sum()
print("=== 各部門銷售總額 ===")
print(dept_sum)

# COUNTIF 的替代：各產品銷售筆數
product_count = df.groupby('產品類別')['金額'].count()
print("\n=== 各產品銷售筆數 ===")
print(product_count)

# AVERAGEIF 的替代：各業務員平均單筆金額
avg_sales = df.groupby('業務員')['金額'].mean().round(0)
print("\n=== 各業務員平均單筆金額 ===")
print(avg_sales)
```

### 範例 18：排序與排名（取代 RANK / 排序功能）

```python
import pandas as pd

df = pd.read_excel('銷售資料.xlsx')

# 依金額排序（由大到小）
df_sorted = df.sort_values('金額', ascending=False)
print("=== 銷售金額排行 ===")
print(df_sorted.head(10))

# 加上排名欄位
df['排名'] = df['金額'].rank(ascending=False, method='min').astype(int)
print(df[['業務員', '金額', '排名']].sort_values('排名'))
```

---

## 2.5 自動產生圖表（取代 Excel 圖表）

### 範例 19：自動產生長條圖與圓餅圖

```python
import pandas as pd
import matplotlib.pyplot as plt

# 設定中文字體
plt.rcParams['font.sans-serif'] = ['Noto Sans CJK TC', 'Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False

df = pd.read_excel('銷售資料.xlsx')

# 長條圖：各業務員銷售額
fig, axes = plt.subplots(1, 2, figsize=(14, 5))

sales_by_person = df.groupby('業務員')['金額'].sum().sort_values(ascending=False)
sales_by_person.plot(kind='bar', ax=axes[0], color='steelblue')
axes[0].set_title('各業務員銷售總額')
axes[0].set_ylabel('金額 (NT$)')

# 圓餅圖：各產品類別占比
sales_by_product = df.groupby('產品類別')['金額'].sum()
sales_by_product.plot(kind='pie', ax=axes[1], autopct='%1.1f%%')
axes[1].set_title('各產品類別銷售占比')
axes[1].set_ylabel('')

plt.tight_layout()
plt.savefig('銷售分析圖.png', dpi=150, bbox_inches='tight')
plt.show()
print("圖表已儲存為 銷售分析圖.png")
```

---

## 2.6 自動生成 Excel 報表並下載

### 範例 20：產生格式化 Excel 報表

```python
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

df = pd.read_excel('銷售資料.xlsx')

# 建立統計資料
summary = df.groupby('業務員').agg(
    訂單數=('金額', 'count'),
    總金額=('金額', 'sum'),
    平均金額=('金額', 'mean'),
    最大單筆=('金額', 'max')
).round(0).reset_index()

# 建立格式化 Excel
wb = Workbook()
ws = wb.active
ws.title = "業務員績效報表"

# 標題
ws.merge_cells('A1:E1')
ws['A1'] = '2024 年度業務員績效報表'
ws['A1'].font = Font(size=16, bold=True, color='FFFFFF')
ws['A1'].fill = PatternFill(start_color='4472C4', fill_type='solid')
ws['A1'].alignment = Alignment(horizontal='center')

# 表頭
headers = ['業務員', '訂單數', '總金額', '平均金額', '最大單筆']
header_fill = PatternFill(start_color='D9E2F3', fill_type='solid')
for col, header in enumerate(headers, 1):
    cell = ws.cell(row=3, column=col, value=header)
    cell.font = Font(bold=True)
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center')

# 寫入資料
for r_idx, row in enumerate(dataframe_to_rows(summary, index=False, header=False), 4):
    for c_idx, value in enumerate(row, 1):
        cell = ws.cell(row=r_idx, column=c_idx, value=value)
        if c_idx >= 3:  # 金額欄位加千分位
            cell.number_format = '#,##0'

# 自動調整欄寬
for col in ws.columns:
    ws.column_dimensions[col[0].column_letter].width = 15

# 儲存
wb.save('業務員績效報表.xlsx')
print("報表已生成：業務員績效報表.xlsx")

# 在 Colab 中下載
from google.colab import files
files.download('業務員績效報表.xlsx')
```

### 範例 21：批次處理多個 Excel 檔案

```python
import pandas as pd
import glob

# 讀取資料夾中所有 Excel 檔案
all_files = glob.glob('*.xlsx')
print(f"找到 {len(all_files)} 個 Excel 檔案")

all_data = []
for file in all_files:
    df = pd.read_excel(file)
    df['來源檔案'] = file  # 標記來源
    all_data.append(df)
    print(f"  已讀取：{file}（{len(df)} 筆）")

# 合併所有資料
combined = pd.concat(all_data, ignore_index=True)
print(f"\n合併完成，共 {len(combined)} 筆資料")

# 儲存合併結果
combined.to_excel('合併結果.xlsx', index=False)
print("已儲存為：合併結果.xlsx")
```

### 範例 22：自動去除重複資料 & 清洗

```python
import pandas as pd

df = pd.read_excel('原始資料.xlsx')
print(f"原始資料：{len(df)} 筆")

# 去除完全重複的列
df = df.drop_duplicates()
print(f"去除重複後：{len(df)} 筆")

# 去除空值
df = df.dropna(subset=['姓名', '金額'])  # 姓名和金額不可為空
print(f"去除空值後：{len(df)} 筆")

# 修正資料格式
df['姓名'] = df['姓名'].str.strip()       # 去除前後空白
df['電話'] = df['電話'].astype(str)         # 確保電話是文字
df['金額'] = pd.to_numeric(df['金額'], errors='coerce')  # 確保金額是數字

# 儲存清洗後的資料
df.to_excel('清洗後資料.xlsx', index=False)
print("清洗完成！已儲存。")
```

---

## 2.7 在 Colab 內直接呼叫 AI 操作 Excel（Gemini）

### 範例 23：用自然語言請 Colab AI 分析並產生程式

在 Colab 的 AI 面板輸入：

```text
請讀取目前工作目錄的 銷售資料.xlsx，
先告訴我這份資料的欄位與資料品質風險（空值、格式不一致），
再產生一段可執行的 pandas 程式：
1. 統計各產品類別銷售總額
2. 列出前 5 名業務員
3. 將結果輸出為 AI分析結果.xlsx
最後用繁體中文摘要 3 個發現。
```

### 範例 23-進階：需要程式化串接時再使用 Gemini API

若你希望把 AI 分析流程包成可重複執行的程式，再改用 API 版本：

```python
# 先安裝 Google AI SDK
!pip install google-generativeai

import pandas as pd
import google.generativeai as genai

# 設定 API Key（從 https://aistudio.google.com/apikey 取得）
genai.configure(api_key='你的_API_KEY')

# 讀取 Excel
df = pd.read_excel('銷售資料.xlsx')

# 把資料摘要傳給 AI 分析
data_summary = f"""
以下是銷售資料的摘要：

欄位：{list(df.columns)}
資料筆數：{len(df)}
數值欄位統計：
{df.describe().to_string()}

前 5 筆範例：
{df.head().to_string()}
"""

model = genai.GenerativeModel('gemini-2.0-flash')
response = model.generate_content(
    f"請用繁體中文分析以下銷售資料，給出 3 個重要發現和建議：\n{data_summary}"
)
print("=== AI 分析結果 ===")
print(response.text)
```

---

## 2.8 課堂練習（15 分鐘）

> **練習題：** 在 Colab 中完成以下任務：
> 1. 上傳任意一個 Excel 檔案，用 pandas 讀取並顯示前 10 筆
> 2. 使用 Colab 內建 AI（Gemini）產生 `groupby` 分組統計程式
> 3. 生成一個長條圖
> 4. 將結果儲存為新的 Excel 檔案並下載

---

# 單元三：Gemini CLI 操作 Excel（13:30-15:00）

> **本單元學習目標：**
> - 能安裝並設定 Gemini CLI
> - 能用自然語言指令讓 Gemini CLI 讀取、分析、生成 Excel
> - 能用 Gemini CLI 比對多個 Excel 檔案並輸出差異報告
>
> **難度：** ⭐⭐☆（初階）

## 3.1 什麼是 Gemini CLI？

- Google 推出的命令列 AI 工具
- 可以直接在終端機中用自然語言操作檔案
- **免費使用**（有每日額度限制）
- 可以直接讀取、修改、生成 Excel 檔案

> **和單元二的 Colab 內建 AI 差別：**
> - Colab 內建 AI：適合在筆記本內協助寫/改程式、分析已上傳資料
> - Gemini CLI：適合在本機終端機直接批次處理資料夾中的多個檔案

## 3.2 安裝 Gemini CLI

### Mac / Linux / Windows（WSL）

```bash
# 安裝 Node.js（如果沒有的話）
# 從 https://nodejs.org 下載安裝（建議 v18 以上）

# 安裝 Gemini CLI（Google 官方套件）
npm install -g @google/gemini-cli

# 驗證安裝
gemini --version

# 首次使用：登入 Google 帳號
gemini auth login
```

### Windows（PowerShell）

```powershell
# 安裝 Gemini CLI
npm install -g @google/gemini-cli

# 設定環境變數（永久）
[System.Environment]::SetEnvironmentVariable("GEMINI_API_KEY", "你的API金鑰", "User")
```

### 替代方案：直接使用 Python SDK

```bash
# 安裝 Google GenAI SDK
pip install google-genai

# 設定環境變數
export GEMINI_API_KEY="你的API金鑰"  # Mac/Linux
# set GEMINI_API_KEY=你的API金鑰      # Windows CMD
```

> **取得 API Key：** 前往 https://aistudio.google.com/apikey → 點擊「Create API Key」→ 複製金鑰

## 3.3 Gemini CLI 操作 Excel 範例

### 範例 24：直接請 Gemini CLI 讀取 Excel

在終端機中輸入：

```
gemini

> 請讀取 銷售資料.xlsx 這個檔案，告訴我裡面有什麼內容
```

**Gemini 會自動：**
1. 識別檔案格式
2. 讀取所有工作表
3. 顯示資料摘要

### 範例 25：請 Gemini CLI 做資料分析

```
> 分析 銷售資料.xlsx 中的銷售趨勢，
  哪個業務員表現最好？哪個產品最暢銷？
  請用表格方式呈現分析結果。
```

### 範例 26：請 Gemini CLI 生成新的 Excel

```
> 請幫我讀取 銷售資料.xlsx，
  然後產生一個新的 Excel 檔案叫「月報.xlsx」，
  包含以下工作表：
  1.「月度摘要」：每月銷售總額、訂單數、平均單價
  2.「業務排行」：業務員依銷售額排名
  3.「產品分析」：各產品銷售量和占比

  每個表格要有格式：標題粗體、金額有千分位、加上框線。
```

### 範例 27：請 Gemini CLI 比對兩個 Excel

```
> 我有兩個 Excel 檔案：
  - 上月庫存.xlsx
  - 本月庫存.xlsx

  兩個檔案都有「產品編號」和「庫存量」欄位。
  請幫我比對差異，找出：
  1. 庫存增加的產品
  2. 庫存減少的產品
  3. 新增的產品
  4. 消失的產品

  結果存成「庫存異動.xlsx」。
```

### 範例 28：請 Gemini CLI 自動整理凌亂資料

```
> 請讀取 原始資料.xlsx，這份資料很亂，請幫我：
  1. 移除重複資料
  2. 統一日期格式為 YYYY-MM-DD
  3. 姓名欄位去除多餘空白
  4. 金額欄位中的文字（如「元」「NTD」）去掉，只保留數字
  5. 電話號碼統一格式為 09XX-XXX-XXX

  整理完存成「清理後資料.xlsx」。
```

---

## 3.4 Gemini CLI 批次自動化

### 範例 29：批次處理多個檔案

```
> 目前資料夾中有多個 Excel 檔案（Jan.xlsx, Feb.xlsx, ... Dec.xlsx），
  請幫我：
  1. 讀取全部檔案
  2. 合併成一個「全年度.xlsx」
  3. 每個月份的資料加上「月份」欄位
  4. 在最後加一個「統計」工作表，顯示每月小計
```

### 範例 30：自動產生報表 + 圖表

```
> 請讀取 銷售資料.xlsx，然後生成一個完整的報告：
  1. 產生「分析報告.xlsx」，包含：
     - 摘要頁：KPI 數據（總營收、成長率、平均客單價）
     - 趨勢圖：月營收走勢（以嵌入圖表方式）
     - 明細頁：原始資料
  2. 另外產生一個「圖表.png」，包含營收趨勢線圖
```

---

## 3.5 Gemini CLI 寫 Python 腳本處理 Excel

### 範例 31：請 Gemini CLI 寫可重複使用的腳本

```
> 幫我寫一個 Python 腳本 process_excel.py，功能是：
  1. 接受命令列參數指定輸入檔案
  2. 讀取 Excel
  3. 自動偵測數值欄位並做統計
  4. 輸出統計結果到新的 Excel
  5. 附上使用說明

  使用方式要像這樣：
  python process_excel.py 輸入檔案.xlsx
```

**Gemini CLI 生成的腳本：**

```python
#!/usr/bin/env python3
"""
Excel 自動統計分析腳本
使用方式：python process_excel.py <輸入檔案.xlsx> [輸出檔案.xlsx]
"""

import sys
import pandas as pd
from datetime import datetime

def analyze_excel(input_file, output_file=None):
    if output_file is None:
        output_file = f"分析結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    # 讀取資料
    print(f"正在讀取：{input_file}")
    df = pd.read_excel(input_file)
    print(f"共 {len(df)} 筆資料，{len(df.columns)} 個欄位")

    # 自動偵測數值欄位
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    text_cols = df.select_dtypes(include='object').columns.tolist()

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 原始資料
        df.to_excel(writer, sheet_name='原始資料', index=False)

        # 數值統計
        if numeric_cols:
            stats = df[numeric_cols].describe().round(2)
            stats.to_excel(writer, sheet_name='數值統計')

        # 文字欄位統計
        for col in text_cols:
            value_counts = df[col].value_counts().reset_index()
            value_counts.columns = [col, '出現次數']
            value_counts.to_excel(writer, sheet_name=f'{col}_統計'[:31], index=False)

    print(f"分析完成！結果已儲存至：{output_file}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("使用方式：python process_excel.py <輸入檔案.xlsx> [輸出檔案.xlsx]")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    analyze_excel(input_file, output_file)
```

---

## 3.6 課堂練習（15 分鐘）

> **練習題：**
> 1. 用 Gemini CLI（或類似 AI CLI 工具）讀取一個 Excel 檔案
> 2. 請它分析資料並指出 3 個重要發現
> 3. 請它生成一個新的統計報表 Excel
> 4. 請它寫一個可重複使用的 Python 腳本

---

# 單元四：綜合實戰演練（15:15-16:30）

> **本單元學習目標：**
> - 能整合 ChatGPT + Colab + Gemini CLI 完成完整自動化流程
> - 能依照實際需求選擇最適合的 AI 工具
> - 完成一個從原始資料到自動寄出月報的完整範例
>
> **難度：** ⭐⭐⭐（中階）

## 4.1 實戰情境：公司月報自動化

### 情境描述

你是公司的業務助理，每月需要：
1. 收到 5 個部門的 Excel 銷售報表
2. 合併所有資料
3. 製作月報（含圖表）
4. 寄給主管

### Step 1：用 ChatGPT 規劃工作流程

**對 ChatGPT 說：**
```
我每個月要處理 5 個部門的 Excel 銷售報表，然後合併、分析、做圖表、寄給主管。
請幫我規劃自動化的流程，並建議適合的工具。
```

### Step 2：用 Colab（可搭配內建 AI）寫自動化程式

```python
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
import warnings
warnings.filterwarnings('ignore')

# ===== 步驟 1：合併所有部門資料 =====
departments = ['業務部', '行銷部', '客服部', '研發部', '管理部']
all_data = []

for dept in departments:
    try:
        df = pd.read_excel(f'{dept}.xlsx')
        df['部門'] = dept
        all_data.append(df)
        print(f"已讀取：{dept} ({len(df)} 筆)")
    except FileNotFoundError:
        print(f"找不到：{dept}.xlsx，跳過")

combined = pd.concat(all_data, ignore_index=True)
print(f"\n合併完成，共 {len(combined)} 筆資料")

# ===== 步驟 2：統計分析 =====
# 各部門銷售統計
dept_summary = combined.groupby('部門').agg(
    訂單數=('金額', 'count'),
    總營收=('金額', 'sum'),
    平均客單價=('金額', 'mean'),
    最大訂單=('金額', 'max')
).round(0)

# 月度趨勢
combined['月份'] = pd.to_datetime(combined['日期']).dt.month
monthly = combined.groupby('月份')['金額'].sum()

# Top 10 業務員
top_sales = combined.groupby('業務員')['金額'].sum().nlargest(10)

print("\n=== 各部門銷售統計 ===")
print(dept_summary)

# ===== 步驟 3：生成報表 =====
with pd.ExcelWriter('月報.xlsx', engine='openpyxl') as writer:
    dept_summary.to_excel(writer, sheet_name='部門摘要')
    top_sales.to_excel(writer, sheet_name='業務員排行')
    combined.to_excel(writer, sheet_name='完整明細', index=False)

print("\n月報已生成：月報.xlsx")

# ===== 步驟 4：生成圖表 =====
fig, axes = plt.subplots(2, 2, figsize=(14, 10))

# 各部門營收
dept_summary['總營收'].plot(kind='bar', ax=axes[0,0], color='steelblue')
axes[0,0].set_title('各部門營收')

# 月度趨勢
monthly.plot(kind='line', ax=axes[0,1], marker='o', color='green')
axes[0,1].set_title('月度營收趨勢')

# Top 10 業務員
top_sales.plot(kind='barh', ax=axes[1,0], color='coral')
axes[1,0].set_title('Top 10 業務員')

# 各部門占比
dept_summary['總營收'].plot(kind='pie', ax=axes[1,1], autopct='%1.1f%%')
axes[1,1].set_title('營收占比')

plt.tight_layout()
plt.savefig('月報圖表.png', dpi=150)
plt.show()
print("圖表已儲存：月報圖表.png")
```

### Step 3：用 Gemini CLI 或 ChatGPT 撰寫報告

```
> 以下是本月銷售數據摘要：
  （貼上步驟 2 的統計結果）

  請用繁體中文撰寫一份簡短的月報摘要（200 字以內），
  包含：
  1. 本月整體表現
  2. 表現最好的部門
  3. 需要注意的地方
  4. 下月建議
```

---

## 4.2 快速對照表：Excel 操作 → AI 替代方案

| Excel 傳統操作 | ChatGPT | Colab (Python) | Gemini CLI |
|:---|:---|:---|:---|
| VLOOKUP | 描述需求→生成公式 | `pd.merge()` | 「幫我查找...」 |
| SUMIF / COUNTIF | 描述需求→生成公式 | `df.groupby().sum()` | 「幫我統計...」 |
| 樞紐分析表 | 描述需求→生成操作步驟 | `df.pivot_table()` | 「幫我做樞紐分析」 |
| IF 巢狀判斷 | 描述需求→生成公式 | `df.apply(函數)` | 「幫我分類...」 |
| 排序 / 篩選 | 描述需求→生成操作步驟 | `df.sort_values()` / `df[條件]` | 「幫我排序...」 |
| 圖表製作 | 描述需求→教你操作 | `matplotlib` / `plotly` | 「幫我畫圖...」 |
| 合併多檔案 | 生成 Python 腳本 | `pd.concat()` | 「幫我合併...」 |
| 去除重複 | 描述需求→教你操作 | `df.drop_duplicates()` | 「幫我去重...」 |
| 格式化報表 | 生成 Python 腳本 | `openpyxl` 設定格式 | 「幫我美化報表」 |
| 發送 Email | 生成 Python 郵件程式 | `smtplib` 自動寄信 | 「幫我寄信...」 |

---

## 4.3 各工具適用場景

```
┌─────────────┬──────────────────────────────────────┐
│   工具       │   最適合的場景                         │
├─────────────┼──────────────────────────────────────┤
│ ChatGPT     │ ● 不會寫公式，需要即時幫助              │
│             │ ● 需要解釋複雜公式                      │
│             │ ● 一次性的 Python 腳本需求              │
│             │ ● 學習 Excel 操作的好老師               │
├─────────────┼──────────────────────────────────────┤
│ Colab       │ ● 大量資料處理（超過 10 萬筆）           │
│ (Python)    │ ● 需要重複執行的分析流程                 │
│             │ ● 複雜的資料清洗和轉換                   │
│             │ ● 專業圖表和視覺化                      │
│             │ ● 串接 API 做自動化                     │
├─────────────┼──────────────────────────────────────┤
│ Gemini CLI  │ ● 快速讀取並理解 Excel 內容              │
│             │ ● 用自然語言完成複雜操作                  │
│             │ ● 不想寫程式但需要自動化                  │
│             │ ● 比對、合併多個檔案                     │
│             │ ● 直接在終端機中操作最高效                │
└─────────────┴──────────────────────────────────────┘
```

---

# 附錄 A：常用 Prompt 模板（Prompt 工程技巧）

## A.0 寫好 Prompt 的五個原則

| 原則 | 說明 | 例子 |
|------|------|------|
| **具體說明資料結構** | 告訴 AI 欄位名稱與格式 | 「A 欄是日期（YYYY/MM/DD 格式）」 |
| **描述期望結果** | 說清楚你要什麼輸出 | 「輸出一個 Excel，有千分位格式」 |
| **提供範例** | 給出輸入/輸出範例 | 「輸入 A123456789，輸出 123/45/67」 |
| **說明限制條件** | 提醒邊界情況 | 「如果找不到資料，顯示空白而非錯誤」 |
| **要求說明** | 請 AI 解釋其做法 | 「請同時說明每一步的用途」 |

---

## A.1 公式生成 Prompt

```
我的 Excel 表格結構如下：
- A 欄：[欄位名稱]
- B 欄：[欄位名稱]
- C 欄：[欄位名稱]

我想要在 [X] 欄計算 [描述你要什麼結果]。
資料從第 2 行開始（第 1 行是標題）。
請給我 Excel 公式。
```

## A.2 Python 腳本生成 Prompt（初階）

```
請幫我寫 Python 腳本處理 Excel：
- 輸入檔案：[檔案名稱和格式]
- 我想完成的任務：[一句話描述]
- 輸出檔案：[期望結果檔名]

請提供：
1. 完整可執行程式碼
2. 套件安裝指令
3. 執行方式（command）
4. 程式每一步在做什麼（用繁體中文簡短說明）
```

## A.3 資料分析 Prompt

```
以下是我的資料摘要：
[貼上資料的前幾筆或統計描述]

請幫我：
1. 找出 3 個關鍵發現
2. 指出異常數據
3. 給出可行的建議
4. 用表格整理重點
```

## A.4 Python 腳本生成 Prompt（進階）

```
請幫我寫 Python 腳本處理 Excel：
- 輸入檔案：[檔案名稱和格式]
- 處理步驟：
  1. [步驟一]
  2. [步驟二]
  3. [步驟三]
- 輸出：[期望的結果檔案]
- 需要的套件：pandas, openpyxl
- 額外要求：
    - 加上錯誤處理（檔案不存在、欄位缺失）
    - 將主要流程拆成函式
    - 支援命令列參數（python script.py input.xlsx output.xlsx）
    - 請加上繁體中文註解與使用說明
```

## A.5 除錯 Prompt（當公式或程式出錯時）

```
我用了以下的 [公式/Python 程式]，但出現了錯誤：

[貼上你的公式或程式碼]

錯誤訊息是：
[貼上錯誤訊息]

我的資料格式是：
[描述資料結構]

請幫我找出問題並修正。
```

## A.6 轉換既有 Excel 流程 Prompt

```
我目前用 Excel 的做法是：
1. [步驟一，例如：手動篩選 B 欄等於「台北」的資料]
2. [步驟二]
3. [步驟三]

這個流程每次要花我大約 [時間]，
請幫我用 Python 將這整個流程自動化，
讓我只需要執行一個腳本就能完成。
```

---

# 附錄 B：環境建置快速指南

## B.1 ChatGPT

1. 前往 https://chat.openai.com
2. 註冊/登入 Google 帳號
3. 免費版即可使用（依平台當下可用模型為準）
4. 若需更穩定品質與進階功能，可評估升級付費方案

## B.2 Google Colab

1. 前往 https://colab.research.google.com
2. 用 Google 帳號登入
3. 新增筆記本
4. 第一格執行：`!pip install openpyxl pandas matplotlib`
5. 需要 AI 協助時，可開啟 Colab 右上角 AI（Gemini）面板，直接用自然語言請它產生或修改程式

## B.3 Gemini CLI

```bash
# 方法一：使用 npm（推薦）
npm install -g @google/gemini-cli

# 驗證安裝成功
gemini --version

# 登入 Google 帳號並授權
gemini auth login

# 或設定 API Key（從 https://aistudio.google.com/apikey 取得）
export GEMINI_API_KEY="你的金鑰"  # Mac/Linux
# Windows PowerShell: $env:GEMINI_API_KEY="你的金鑰"

# 方法二：使用 Python SDK
pip install google-genai pandas openpyxl
```

> **注意：** 確認使用的套件名稱是 `@google/gemini-cli`（Google 官方），
> 而非其他第三方套件。

> **補充：什麼時候用 Colab 內建 AI、什麼時候用 Gemini CLI？**
> - 在 Colab 內做資料分析、教學示範：優先用 Colab 內建 AI
> - 需要操作本機資料夾、多檔批次流程：優先用 Gemini CLI

---

# 附錄 C：課後延伸學習資源

| 資源 | 連結 | 說明 |
|------|------|------|
| Pandas 官方文件 | https://pandas.pydata.org/docs/ | Python 資料分析必備 |
| OpenPyXL 文件 | https://openpyxl.readthedocs.io/ | Python 操作 Excel |
| Google AI Studio | https://aistudio.google.com/ | 免費使用 Gemini API |
| ChatGPT | https://chat.openai.com | AI 對話助手 |
| Colab 教學 | https://colab.research.google.com | 雲端 Python 環境 |

---

# 附錄 D：常見錯誤與排解

## D.1 Python / Colab 常見錯誤

| 錯誤訊息 | 可能原因 | 解決方式 |
|---------|---------|---------|
| `ModuleNotFoundError: No module named 'openpyxl'` | 套件未安裝 | 執行 `!pip install openpyxl` |
| `FileNotFoundError: 銷售資料.xlsx` | 檔案未上傳或路徑錯誤 | 重新上傳檔案，確認名稱一致 |
| `KeyError: '金額'` | 欄位名稱不符 | 用 `print(df.columns)` 確認實際欄位名 |
| `UnicodeDecodeError` | 編碼問題 | 在 `read_excel` 加上 `encoding='utf-8-sig'` |
| `ValueError: could not convert string to float` | 金額欄有文字 | 用 `pd.to_numeric(df['金額'], errors='coerce')` |
| 中文圖表顯示方框 | 缺少中文字體 | 加上 `plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']` |

## D.2 Gemini CLI 常見問題

| 問題 | 解決方式 |
|------|---------|
| `command not found: gemini` | 重新執行 `npm install -g @google/gemini-cli`，確認 Node.js 已安裝 |
| API Key 失效 | 前往 aistudio.google.com 重新生成 Key |
| 回應慢或逾時 | 檢查網路連線，或改用較小的 Excel 檔案測試 |
| 中文輸入有問題 | 改用英文輸入或在腳本中以文字傳入中文 |
| 達到每日配額上限 | 等待隔天或升級付費方案 |

## D.4 Colab 內建 AI 常見問題

| 問題 | 解決方式 |
|------|---------|
| 找不到 Colab 的 AI 面板 | 確認已登入 Google 帳號，重新整理頁面或切換新筆記本再試 |
| AI 產生的程式無法執行 | 先檢查檔名與欄位是否一致，再把錯誤訊息貼回 AI 讓它修正 |
| AI 沒讀到最新資料 | 重新執行上傳與讀檔儲存格，確認 DataFrame 變數已更新 |
| 需要操作本機多檔案但 Colab 不方便 | 改用 Gemini CLI 在本機終端機直接處理 |

## D.3 ChatGPT 公式常見問題

| 問題 | 解決方式 |
|------|---------|
| 公式出現 `#N/A` | VLOOKUP 查無此值，加上 `IFERROR(公式, "無資料")` |
| 公式出現 `#REF!` | 參照範圍被刪除，重新確認範圍 |
| 公式出現 `#VALUE!` | 資料型態錯誤，確認數字欄位不含文字 |
| 日期顯示為數字 | 選取儲存格，格式→日期 |
| 公式複製後結果不對 | 確認是否需要用 `$` 鎖定參照（如 `$D$2:$D$100`）|

---

# 附錄 E：進階工具比較（2024-2025 年）

## E.1 AI 工具功能比較

| 功能 | ChatGPT | Gemini | Claude |
|------|---------|--------|--------|
| 生成 Excel 公式 | ★★★★★ | ★★★★★ | ★★★★★ |
| 生成 Python 腳本 | ★★★★★ | ★★★★★ | ★★★★★ |
| 直接操作本機檔案 | ✗ | ★★★★（CLI；Colab 版需上傳檔案）| ★★★★（Claude Code CLI）|
| 解釋複雜公式 | ★★★★★ | ★★★★★ | ★★★★★ |
| 資料分析洞察 | ★★★★ | ★★★★★ | ★★★★★ |
| 免費使用 | 有限制 | 有限制 | 有限制 |
| 中文支援 | ★★★★★ | ★★★★★ | ★★★★★ |

## E.2 使用 Claude Code CLI 操作 Excel（補充）

Claude Code 是 Anthropic 推出的命令列 AI 工具，可直接在終端機操作本機 Excel 檔案：

```bash
# 安裝 Claude Code CLI
npm install -g @anthropic-ai/claude-code

# 啟動
claude

# 使用範例
> 請讀取 銷售資料.xlsx，幫我找出銷售額最高的前 5 名業務員
> 幫我合併這個資料夾裡所有的 xlsx 檔案，輸出到 合併.xlsx
```

**Claude Code 優勢：**
- 可直接讀取、修改本機 Excel 檔案
- 能執行生成的 Python 腳本並即時回饋結果
- 支援多輪對話持續優化
- 適合複雜的多步驟自動化任務

---

> **課程總結：**
> AI 時代的 Excel 操作不再需要死記公式。
> 善用 ChatGPT、Colab、Gemini CLI，
> 你可以用自然語言完成過去需要數小時的工作。
> **關鍵心法：描述清楚你要什麼，AI 就能幫你做到。**

---

*本教材由 AI 輔助生成，適用於教學與學習用途。*
