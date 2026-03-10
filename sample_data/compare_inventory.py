import zipfile
import xml.etree.ElementTree as ET
import collections

def parse_excel(file_path):
    inventory_data = {}
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            # Get shared strings
            strings = []
            try:
                with z.open('xl/sharedStrings.xml') as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    for si in root.findall('ns:si', ns):
                        t = si.find('ns:t', ns)
                        if t is not None:
                            strings.append(t.text)
                        else:
                            t_text = ""
                            for r in si.findall('ns:r', ns):
                                t_part = r.find('ns:t', ns)
                                if t_part is not None:
                                    t_text += t_part.text
                            strings.append(t_text)
            except KeyError:
                # No shared strings
                pass

            # Get sheet data
            with z.open('xl/worksheets/sheet1.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                
                sheet_data = root.find('ns:sheetData', ns)
                
                id_col = None
                name_col = None
                stock_col = None
                
                for row in sheet_data.findall('ns:row', ns):
                    row_idx = int(row.get('r'))
                    row_values = {}
                    for cell in row.findall('ns:c', ns):
                        r = cell.get('r')
                        col = ''.join([c for c in r if not c.isdigit()])
                        t = cell.get('t')
                        v = cell.find('ns:v', ns)
                        val = None
                        if v is not None:
                            val = v.text
                            if t == 's':
                                try:
                                    val = strings[int(val)]
                                except:
                                    pass
                        row_values[col] = val
                    
                    if row_idx == 1:
                        # Look for headers
                        for col, val in row_values.items():
                            if val == '產品編號':
                                id_col = col
                            elif val == '產品名稱':
                                name_col = col
                            elif val == '庫存量':
                                stock_col = col
                    else:
                        if id_col and name_col and stock_col:
                            pid = row_values.get(id_col)
                            pname = row_values.get(name_col)
                            pstock = row_values.get(stock_col)
                            if pid:
                                try:
                                    inventory_data[pid] = (pname, float(pstock) if pstock is not None else 0.0)
                                except ValueError:
                                    inventory_data[pid] = (pname, 0.0)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return inventory_data

def compare():
    old_inv = parse_excel('上月庫存.xlsx')
    new_inv = parse_excel('本月庫存.xlsx')
    
    increased = []
    decreased = []
    added = []
    removed = []
    
    for pid, (pname, pstock) in old_inv.items():
        if pid in new_inv:
            new_name, new_stock = new_inv[pid]
            if new_stock > pstock:
                increased.append((pid, pname, pstock, new_stock))
            elif new_stock < pstock:
                decreased.append((pid, pname, pstock, new_stock))
        else:
            removed.append((pid, pname, pstock))
            
    for pid, (pname, pstock) in new_inv.items():
        if pid not in old_inv:
            added.append((pid, pname, pstock))
            
    print("\n### 1. 庫存增加的產品")
    print("| 產品編號 | 產品名稱 | 上月庫存 | 本月庫存 | 增量 |")
    print("| --- | --- | --- | --- | --- |")
    for pid, pname, old, new in increased:
        print(f"| {pid} | {pname} | {old:.0f} | {new:.0f} | +{new-old:.0f} |")
        
    print("\n### 2. 庫存減少的產品")
    print("| 產品編號 | 產品名稱 | 上月庫存 | 本月庫存 | 減量 |")
    print("| --- | --- | --- | --- | --- |")
    for pid, pname, old, new in decreased:
        print(f"| {pid} | {pname} | {old:.0f} | {new:.0f} | {new-old:.0f} |")
        
    print("\n### 3. 本月新增的產品")
    print("| 產品編號 | 產品名稱 | 本月庫存 |")
    print("| --- | --- | --- |")
    for pid, pname, stock in added:
        print(f"| {pid} | {pname} | {stock:.0f} |")
        
    print("\n### 4. 本月消失的產品")
    print("| 產品編號 | 產品名稱 | 上月庫存 |")
    print("| --- | --- | --- |")
    for pid, pname, stock in removed:
        print(f"| {pid} | {pname} | {stock:.0f} |")

if __name__ == "__main__":
    compare()
