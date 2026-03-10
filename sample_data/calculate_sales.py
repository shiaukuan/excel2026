import zipfile
import xml.etree.ElementTree as ET
import collections

def parse_excel(file_path):
    with zipfile.ZipFile(file_path, 'r') as z:
        # Get shared strings
        strings = []
        with z.open('xl/sharedStrings.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in root.findall('ns:si', ns):
                t = si.find('ns:t', ns)
                if t is not None:
                    strings.append(t.text)
                else:
                    # Sometimes strings are rich text
                    t_text = ""
                    for r in si.findall('ns:r', ns):
                        t_part = r.find('ns:t', ns)
                        if t_part is not None:
                            t_text += t_part.text
                    strings.append(t_text)

        # Get sheet data
        sales_data = collections.defaultdict(float)
        with z.open('xl/worksheets/sheet1.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            # Find row data
            sheet_data = root.find('ns:sheetData', ns)
            rows = sheet_data.findall('ns:row', ns)
            
            headers_mapping = {}
            salesperson_col = None
            amount_col = None
            
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
                        if val == '業務員':
                            salesperson_col = col
                        if val == '金 額':
                            amount_col = col
                    if not salesperson_col or not amount_col:
                        # Sometimes header names have spaces or are slightly different
                        for col, val in row_values.items():
                            if val and '業務員' in str(val):
                                salesperson_col = col
                            if val and '金' in str(val) and '額' in str(val):
                                amount_col = col
                else:
                    if salesperson_col and amount_col:
                        name = row_values.get(salesperson_col)
                        amount = row_values.get(amount_col)
                        if name and amount is not None:
                            try:
                                sales_data[name] += float(amount)
                            except:
                                pass
    return sales_data

if __name__ == "__main__":
    sales = parse_excel('銷售資料.xlsx')
    # Sort by value descending
    sorted_sales = sorted(sales.items(), key=lambda x: x[1], reverse=True)
    
    print("| 業務員 | 銷售總金額 |")
    print("| --- | --- |")
    for name, amount in sorted_sales:
        print(f"| {name} | {amount:,.0f} |")
