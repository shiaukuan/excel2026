#!/usr/bin/env python3
"""
AI 時代 Excel 教學課程 — 範例資料生成腳本
執行後會在 sample_data/ 資料夾中產生所有教學用 Excel 檔案。

使用方式：
    pip install pandas openpyxl xlsxwriter
    python generate_sample_data.py
"""

import os
import random
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

random.seed(42)
np.random.seed(42)

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sample_data')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ============================================================
# 共用假資料
# ============================================================
NAMES_SALES = ['王小明', '李美玲', '張志豪', '陳雅婷', '林建宏', '黃淑芬']
NAMES_EXTRA = ['吳俊傑', '蔡佳蓉', '鄭文彬', '許雅琪']
PRODUCTS = ['電子產品', '辦公用品', '食品飲料', '生活用品']
CUSTOMERS = [
    '台灣科技股份有限公司', '大同企業', '鴻海精密', '統一集團',
    '中華電信', '台積電', '富邦金控', '國泰人壽',
    '遠東百貨', '全聯實業', '家樂福', '寶雅國際',
    '微風廣場', '新光三越', '台灣大哥大', '亞太電信',
]
DEPARTMENTS = ['業務部', '行銷部', '客服部', '研發部', '管理部']
SUBJECTS = ['國文', '英文', '數學', '自然', '社會']
STUDENT_NAMES = [
    '王大明', '李小華', '張美玲', '陳志豪', '林雅婷',
    '黃建宏', '吳淑芬', '蔡俊傑', '鄭佳蓉', '許文彬',
    '趙雅琪', '周美慧', '孫志偉', '朱淑惠', '何建志',
    '呂美玲', '施俊宏', '羅雅雯', '傅建中', '蕭美華',
    '葉志明', '潘淑玲', '范建華', '曹雅芳', '彭志遠',
    '宋美芬', '田建民', '廖雅萍', '姚志強', '邱美如',
]
ITEMS = [
    '辦公桌椅', '影印紙', '墨水匣', '筆記本', '文件夾',
    '白板筆', '投影機', '網路設備', '伺服器維護', '軟體授權',
    '員工餐費', '交通費', '會議室租借', '廣告費', '行銷物料',
]
PRODUCT_LIST = [
    ('P001', '筆記型電腦', 28000),
    ('P002', '無線滑鼠', 590),
    ('P003', '機械鍵盤', 2490),
    ('P004', '27吋螢幕', 8900),
    ('P005', 'USB隨身碟', 350),
    ('P006', '印表機', 5600),
    ('P007', '網路攝影機', 1290),
    ('P008', '耳罩式耳機', 1890),
    ('P009', '行動電源', 790),
    ('P010', '平板保護套', 490),
]
INVENTORY_PRODUCTS = [
    ('SKU001', 'A4影印紙'),
    ('SKU002', '原子筆(藍)'),
    ('SKU003', '釘書機'),
    ('SKU004', '膠帶'),
    ('SKU005', '白板筆'),
    ('SKU006', '資料夾'),
    ('SKU007', '便利貼'),
    ('SKU008', '修正帶'),
    ('SKU009', '迴紋針'),
    ('SKU010', '信封'),
    ('SKU011', '計算機'),
    ('SKU012', '剪刀'),
]


def random_date(start, end):
    delta = end - start
    return start + timedelta(days=random.randint(0, delta.days))


# ============================================================
# 1. 銷售資料.xlsx — 貫穿全課程的主要練習檔
# ============================================================
def generate_sales_data():
    print('生成：銷售資料.xlsx ...')
    rows = []
    start = datetime(2024, 1, 1)
    end = datetime(2024, 12, 31)
    for _ in range(200):
        date = random_date(start, end)
        salesperson = random.choice(NAMES_SALES)
        product = random.choice(PRODUCTS)
        customer = random.choice(CUSTOMERS)
        qty = random.randint(1, 50)
        unit_price = random.choice([100, 250, 500, 800, 1200, 2000, 3500, 5000])
        amount = qty * unit_price
        rows.append({
            '日期': date.strftime('%Y/%m/%d'),
            '業務員': salesperson,
            '產品類別': product,
            '客戶名稱': customer,
            '數量': qty,
            '單價': unit_price,
            '金額': amount,
        })
    df = pd.DataFrame(rows)
    df['日期'] = pd.to_datetime(df['日期'])
    df = df.sort_values('日期').reset_index(drop=True)
    df.to_excel(os.path.join(OUTPUT_DIR, '銷售資料.xlsx'), index=False)
    print(f'  完成，共 {len(df)} 筆')


# ============================================================
# 2. 每月資料.xlsx — 12 個月工作表
# ============================================================
def generate_monthly_data():
    print('生成：每月資料.xlsx ...')
    months = ['1月', '2月', '3月', '4月', '5月', '6月',
              '7月', '8月', '9月', '10月', '11月', '12月']
    with pd.ExcelWriter(os.path.join(OUTPUT_DIR, '每月資料.xlsx'), engine='openpyxl') as writer:
        for i, month_name in enumerate(months):
            n_rows = random.randint(15, 25)
            month_num = i + 1
            rows = []
            for _ in range(n_rows):
                day = random.randint(1, 28)
                date = datetime(2024, month_num, day)
                item = random.choice(ITEMS)
                amount = random.randint(500, 50000)
                rows.append({
                    '日期': date.strftime('%Y/%m/%d'),
                    '項目': item,
                    '金額': amount,
                })
            df = pd.DataFrame(rows).sort_values('日期').reset_index(drop=True)
            df.to_excel(writer, sheet_name=month_name, index=False)
    print('  完成，共 12 個工作表')


# ============================================================
# 3. 催款名單.xlsx
# ============================================================
def generate_collection_list():
    print('生成：催款名單.xlsx ...')
    names = ['張三', '李四', '王五', '趙六', '孫七', '周八', '吳九', '鄭十', '陳十一', '林十二']
    rows = []
    for name in names:
        rows.append({
            '收件人姓名': name,
            'Email 地址': f'{name.lower()}@example.com',
            '應付金額': random.choice([5000, 10000, 15000, 20000, 25000, 30000, 50000]),
        })
    df = pd.DataFrame(rows)
    df.to_excel(os.path.join(OUTPUT_DIR, '催款名單.xlsx'), index=False)
    print(f'  完成，共 {len(df)} 筆')


# ============================================================
# 4. 訂單.xlsx + 產品.xlsx — merge/VLOOKUP 練習
# ============================================================
def generate_orders_and_products():
    print('生成：訂單.xlsx + 產品.xlsx ...')
    # 產品表
    products_df = pd.DataFrame(PRODUCT_LIST, columns=['產品編號', '產品名稱', '單價'])
    products_df.to_excel(os.path.join(OUTPUT_DIR, '產品.xlsx'), index=False)

    # 訂單表
    rows = []
    for i in range(50):
        product = random.choice(PRODUCT_LIST)
        rows.append({
            '訂單編號': f'ORD-{i+1:04d}',
            '日期': random_date(datetime(2024, 1, 1), datetime(2024, 12, 31)).strftime('%Y/%m/%d'),
            '客戶名稱': random.choice(CUSTOMERS),
            '產品編號': product[0],
            '數量': random.randint(1, 20),
        })
    orders_df = pd.DataFrame(rows)
    orders_df.to_excel(os.path.join(OUTPUT_DIR, '訂單.xlsx'), index=False)
    print(f'  完成，訂單 {len(orders_df)} 筆，產品 {len(products_df)} 筆')


# ============================================================
# 5. 成績單.xlsx — IF 判斷練習
# ============================================================
def generate_scores():
    print('生成：成績單.xlsx ...')
    rows = []
    for name in STUDENT_NAMES:
        subject = random.choice(SUBJECTS)
        score = random.randint(30, 100)
        rows.append({
            '姓名': name,
            '科目': subject,
            '分數': score,
        })
    df = pd.DataFrame(rows)
    df.to_excel(os.path.join(OUTPUT_DIR, '成績單.xlsx'), index=False)
    print(f'  完成，共 {len(df)} 筆')


# ============================================================
# 6. 原始資料.xlsx — 故意包含髒資料供清洗練習
# ============================================================
def generate_dirty_data():
    print('生成：原始資料.xlsx ...')
    rows = []
    base_names = NAMES_SALES + NAMES_EXTRA
    for i in range(60):
        name = random.choice(base_names)
        # 故意加空白
        if random.random() < 0.15:
            name = '  ' + name + ' '
        # 日期格式不一致
        date_val = random_date(datetime(2024, 1, 1), datetime(2024, 6, 30))
        fmt = random.choice(['%Y/%m/%d', '%Y-%m-%d', '%m/%d/%Y', '%Y年%m月%d日'])
        date_str = date_val.strftime(fmt)
        # 金額故意包含文字
        amount = random.randint(1000, 80000)
        if random.random() < 0.1:
            amount_str = f'NTD {amount}'
        elif random.random() < 0.1:
            amount_str = f'{amount}元'
        else:
            amount_str = str(amount)
        # 電話格式不一致
        phone_base = f'09{random.randint(10,99)}{random.randint(100000,999999)}'
        phone_fmt = random.choice([
            phone_base,
            f'{phone_base[:4]}-{phone_base[4:7]}-{phone_base[7:]}',
            f'{phone_base[:4]} {phone_base[4:7]} {phone_base[7:]}',
        ])

        rows.append({
            '姓名': name,
            '日期': date_str,
            '金額': amount_str,
            '電話': phone_fmt,
            '備註': random.choice(['', '已確認', '待確認', '', '', '需跟進', '']),
        })

    # 故意加入重複資料（複製一些行）
    for _ in range(8):
        rows.append(rows[random.randint(0, len(rows) - 1)].copy())

    # 故意加入空值
    for idx in random.sample(range(len(rows)), 5):
        rows[idx]['姓名'] = None
    for idx in random.sample(range(len(rows)), 3):
        rows[idx]['金額'] = None

    random.shuffle(rows)
    df = pd.DataFrame(rows)
    df.to_excel(os.path.join(OUTPUT_DIR, '原始資料.xlsx'), index=False)
    print(f'  完成，共 {len(df)} 筆（含重複與空值）')


# ============================================================
# 7. 上月庫存.xlsx + 本月庫存.xlsx — 比對差異練習
# ============================================================
def generate_inventory():
    print('生成：上月庫存.xlsx + 本月庫存.xlsx ...')
    # 上月庫存：所有 12 項
    last_month = []
    for sku, name in INVENTORY_PRODUCTS:
        last_month.append({
            '產品編號': sku,
            '產品名稱': name,
            '庫存量': random.randint(50, 500),
        })
    df_last = pd.DataFrame(last_month)
    df_last.to_excel(os.path.join(OUTPUT_DIR, '上月庫存.xlsx'), index=False)

    # 本月庫存：移除 1 項、新增 1 項、其他調整數量
    this_month = []
    removed_idx = random.randint(0, len(INVENTORY_PRODUCTS) - 1)
    for i, (sku, name) in enumerate(INVENTORY_PRODUCTS):
        if i == removed_idx:
            continue  # 消失的產品
        qty_change = random.randint(-100, 200)
        old_qty = df_last.loc[i, '庫存量']
        new_qty = max(0, old_qty + qty_change)
        this_month.append({
            '產品編號': sku,
            '產品名稱': name,
            '庫存量': new_qty,
        })
    # 新增一項
    this_month.append({
        '產品編號': 'SKU013',
        '產品名稱': '環保紙杯',
        '庫存量': random.randint(100, 300),
    })
    df_this = pd.DataFrame(this_month)
    df_this.to_excel(os.path.join(OUTPUT_DIR, '本月庫存.xlsx'), index=False)
    print(f'  完成，上月 {len(df_last)} 項，本月 {len(df_this)} 項')


# ============================================================
# 8. 五個部門 Excel — 單元四綜合演練
# ============================================================
def generate_department_files():
    print('生成：五個部門 Excel ...')
    for dept in DEPARTMENTS:
        n_rows = random.randint(30, 50)
        rows = []
        # 每個部門有 2-3 名業務員
        dept_people = random.sample(NAMES_SALES + NAMES_EXTRA, random.randint(2, 3))
        for _ in range(n_rows):
            date = random_date(datetime(2024, 1, 1), datetime(2024, 12, 31))
            rows.append({
                '日期': date.strftime('%Y/%m/%d'),
                '業務員': random.choice(dept_people),
                '產品類別': random.choice(PRODUCTS),
                '金額': random.randint(1000, 100000),
            })
        df = pd.DataFrame(rows)
        df['日期'] = pd.to_datetime(df['日期'])
        df = df.sort_values('日期').reset_index(drop=True)
        filename = f'{dept}.xlsx'
        df.to_excel(os.path.join(OUTPUT_DIR, filename), index=False)
        print(f'  {filename}：{len(df)} 筆')
    print('  五個部門檔案生成完成')


# ============================================================
# 主程式
# ============================================================
if __name__ == '__main__':
    print('=' * 50)
    print('AI 時代 Excel 教學課程 — 範例資料生成')
    print('=' * 50)
    print(f'輸出目錄：{OUTPUT_DIR}\n')

    generate_sales_data()
    generate_monthly_data()
    generate_collection_list()
    generate_orders_and_products()
    generate_scores()
    generate_dirty_data()
    generate_inventory()
    generate_department_files()

    print('\n' + '=' * 50)
    print('全部範例資料生成完成！')
    print('=' * 50)

    # 列出生成的檔案
    files = sorted(os.listdir(OUTPUT_DIR))
    print(f'\n共 {len(files)} 個檔案：')
    for f in files:
        size = os.path.getsize(os.path.join(OUTPUT_DIR, f))
        print(f'  {f} ({size:,} bytes)')
