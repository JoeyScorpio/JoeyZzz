#!/usr/bin/env python3
"""Generate consolidated Excel with all invoices, with Invoice Date column, sorted by date (oldest first)."""

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Invoice Summary"

# === All 3 invoices data ===
invoices = [
    {
        'date': '01/02/2026',
        'invoice_no': 'DFM8101729',
        'driver': '陈桥桥',
        'note': 'AMENDED',
        'items': [
            ('BCS02',  '豆福豆皮（散装1KG） DOLFORD BEANCURD SKIN',            'KGS', 3.60,  2.00,   7.20),
            ('CST01',  '芝士海鲜豆腐 CHEESE SEAFOOD TOFU',                     'PKT', 3.50,  6.00,  21.00),
            ('ET01P',  '蛋豆腐 EGG TOFU 130g',                                'PKT', 0.25,  6.00,   1.50),
            ('EZ01C',  'EZ快热面 E-ZEE INSTANT NOODLES',                       'CTN', 18.00, 1.00,  18.00),
            ('FD01P',  '炸豆枝片 EB FRIED BEAN CURD SHEET',                    'PKT', 9.30,  2.00,  18.60),
            ('LB02P',  'FG 龙虾球 LOBSTER BALL（FG）',                          'PKT', 3.50,  3.00,  10.50),
            ('NTU01',  '豆福手工卤水大豆腐 DOLFORD NORTH TOFU',                  'PCS', 1.20,  4.00,   4.80),
            ('SAT01',  '豆福五香豆干（散装）1KG DOLFORD SPICED',                  'KGS', 3.60,  1.00,   3.60),
            ('SDO01',  '烟熏鸭胸原味 Smoked Duck Original 5PKT/KG',             'KGS', 9.50,  1.27,  12.07),
            ('VTC02',  '豆福素鸡（散装） DOLFORD VEGETARIAN',                    'KGS', 3.60,  1.00,   3.60),
            ('WFB01',  '白鱼丸 (FG)_1KG*10PKT/CTN WHITE FISH',                'KGS', 5.00,  1.00,   5.00),
            ('ZSB0080','金包银 MUSHROOM FISHN SOY 300G*30PKT/CTN',             'PKT', 2.50,  2.00,   5.00),
            ('NTU01',  'Exchange 交换：豆福手工卤水大豆腐 DOLFORD',               'PCS', None,  4.00,   None),
            ('DB01P',  '豆卜 DARK BROWN BEAN CURD 10PCS X 200G',              'PKT', 1.40, 15.00,  21.00),
        ],
        'subtotal': 131.87,
        'gst': 11.87,
        'total': 143.74,
    },
    {
        'date': '07/02/2026',
        'invoice_no': 'DFM8102534',
        'driver': '顾洪亮',
        'note': '',
        'items': [
            ('BCS02',  '豆福豆皮（散装1KG） DOLFORD BEANCURD SKIN',            'KGS', 3.60,  2.00,   7.20),
            ('CST01',  '芝士海鲜豆腐 CHEESE SEAFOOD TOFU',                     'PKT', 3.50,  5.00,  17.50),
            ('ET01P',  '蛋豆腐 EGG TOFU 130g',                                'PKT', 0.25,  6.00,   1.50),
            ('LB02P',  'FG 龙虾球 LOBSTER BALL（FG）',                          'PKT', 3.50,  3.00,  10.50),
            ('NTU01',  '豆福手工卤水大豆腐 DOLFORD NORTH TOFU',                  'PCS', 1.20,  8.00,   9.60),
            ('SDO01',  '烟熏鸭胸原味 Smoked Duck Original 5PKT/KG',             'KGS', 9.50,  1.14,  10.83),
            ('ZSB0080','金包银 MUSHROOM FISHN SOY 300G*30PKT/CTN',             'PKT', 2.65,  2.00,   5.30),
            ('DB01P',  '豆卜 DARK BROWN BEAN CURD 10PCS X 200G',              'PKT', 1.40, 10.00,  14.00),
            ('SAT01',  '豆福五香豆干（散装）1KG DOLFORD SPICED',                  'KGS', 3.60,  1.00,   3.60),
            ('VYC03',  '豆福素鸡（散装） DOLFORD VEGETARIAN',                    'KGS', 3.60,  1.00,   3.60),
            ('ZSB0082','蟹棒 MUSHROOM IMI.CRAB STICK',                         'PKT', 1.70, 15.00,  25.50),
        ],
        'subtotal': 109.13,
        'gst': 9.82,
        'total': 118.95,
    },
    {
        'date': '08/02/2026',
        'invoice_no': 'DFM8102644',
        'driver': '顾洪亮',
        'note': '',
        'items': [
            ('BCS02',  '豆福豆皮（散装1KG） DOLFORD BEANCURD SKIN',            'KGS', 3.60,  2.00,   7.20),
            ('CST01',  '芝士海鲜豆腐 CHEESE SEAFOOD TOFU',                     'PKT', 3.50,  5.00,  17.50),
            ('ET01P',  '蛋豆腐 EGG TOFU 130g',                                'PKT', 0.25, 10.00,   2.50),
            ('FD01P',  '炸豆枝片 EB FRIED BEAN CURD SHEET',                    'PKT', 9.30,  1.00,   9.30),
            ('LB02P',  'FG 龙虾球 LOBSTER BALL（FG）',                          'PKT', 3.50,  2.00,   7.00),
            ('NTU01',  '豆福手工卤水大豆腐 DOLFORD NORTH TOFU',                  'PCS', 1.20,  4.00,   4.80),
            ('SDO01',  '烟熏鸭胸原味 Smoked Duck Original 5PKT/KG',             'KGS', 9.50,  1.26,  11.97),
            ('WDM010', '日式乌冬面 INSTANT JAPANESE FRESH UDON',               'CTN', 21.00, 1.00,  21.00),
            ('WFB01',  '白鱼丸 (FG)_1KG*10PKT/CTN WHITE FISH',                'KGS', 5.00,  1.00,   5.00),
            ('ZSB0080','金包银 MUSHROOM FISHN SOY 300G*30PKT/CTN',             'PKT', 2.65,  1.00,   2.65),
        ],
        'subtotal': 88.92,
        'gst': 8.00,
        'total': 96.92,
    },
]

# Sort invoices by date (oldest first: dd/mm/yyyy)
invoices.sort(key=lambda inv: tuple(reversed(inv['date'].split('/'))))

# Flatten all rows: each item gets an Invoice Date column
all_rows = []
for inv in invoices:
    for code, desc, unit, price, qty, amt in inv['items']:
        all_rows.append({
            'date': inv['date'],
            'invoice_no': inv['invoice_no'],
            'note': inv['note'],
            'code': code,
            'desc': desc,
            'unit': unit,
            'price': price,
            'qty': qty,
            'amt': amt,
        })

# Styles
title_font = Font(name='Arial', size=14, bold=True)
header_font = Font(name='Arial', size=10, bold=True)
header_font_white = Font(name='Arial', size=10, bold=True, color='FFFFFF')
normal_font = Font(name='Arial', size=10)
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
center = Alignment(horizontal='center', vertical='center')
left_align = Alignment(horizontal='left', vertical='center')
right_align = Alignment(horizontal='right', vertical='center')
header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
subtotal_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
total_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
# Alternating date group colors
date_colors = [
    PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid'),
    PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid'),
]

# Column widths
ws.column_dimensions['A'].width = 6     # No.
ws.column_dimensions['B'].width = 14    # Invoice Date
ws.column_dimensions['C'].width = 16    # Invoice No.
ws.column_dimensions['D'].width = 12    # Code
ws.column_dimensions['E'].width = 45    # Description
ws.column_dimensions['F'].width = 8     # QTY
ws.column_dimensions['G'].width = 8     # UNIT
ws.column_dimensions['H'].width = 10    # U_PRICE
ws.column_dimensions['I'].width = 12    # AMOUNT

# === Title ===
ws.merge_cells('A1:I1')
ws['A1'] = '豆福食品制造有限公司 DOLFORD FOOD MANUFACTURING PTE LTD - Invoice Summary'
ws['A1'].font = title_font
ws['A1'].alignment = center

ws.merge_cells('A2:I2')
ws['A2'] = 'Customer: ZHANG LIANG MALA TANG 张亮麻辣烫 (远东广场)  |  Customer ID: ZLML07  |  Term: 30 DAYS'
ws['A2'].font = Font(name='Arial', size=10)
ws['A2'].alignment = center

# === Header Row ===
row = 4
headers = [
    ('No.', center),
    ('Invoice Date\n发票日期', center),
    ('Invoice No.\n发票号码', center),
    ('Code\n编码', center),
    ('ITEM NAME / DESCRIPTION\n货名/摘要', center),
    ('QTY\n数量', center),
    ('UNIT\n单位', center),
    ('U_PRICE\n单价', center),
    ('AMOUNT($)\n金额', center),
]
for col_idx, (h, align) in enumerate(headers, 1):
    cell = ws.cell(row=row, column=col_idx, value=h)
    cell.font = header_font_white
    cell.fill = header_fill
    cell.border = thin_border
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws.row_dimensions[row].height = 35

# === Data Rows ===
row = 5
# Track date groups for alternating colors
date_list = []
for inv in invoices:
    if inv['date'] not in date_list:
        date_list.append(inv['date'])

for idx, item in enumerate(all_rows):
    r = row + idx
    ws.row_dimensions[r].height = 20
    date_idx = date_list.index(item['date'])
    row_fill = date_colors[date_idx % 2]

    # No.
    cell = ws.cell(row=r, column=1, value=idx + 1)
    cell.font = normal_font; cell.alignment = center; cell.border = thin_border; cell.fill = row_fill

    # Invoice Date
    date_label = item['date']
    if item['note']:
        date_label += f" ({item['note']})"
    cell = ws.cell(row=r, column=2, value=date_label)
    cell.font = normal_font; cell.alignment = center; cell.border = thin_border; cell.fill = row_fill

    # Invoice No.
    cell = ws.cell(row=r, column=3, value=item['invoice_no'])
    cell.font = normal_font; cell.alignment = center; cell.border = thin_border; cell.fill = row_fill

    # Code
    cell = ws.cell(row=r, column=4, value=item['code'])
    cell.font = normal_font; cell.alignment = left_align; cell.border = thin_border; cell.fill = row_fill

    # Description
    cell = ws.cell(row=r, column=5, value=item['desc'])
    cell.font = normal_font; cell.alignment = left_align; cell.border = thin_border; cell.fill = row_fill

    # QTY
    cell = ws.cell(row=r, column=6, value=item['qty'])
    cell.font = normal_font; cell.alignment = center; cell.border = thin_border; cell.fill = row_fill
    cell.number_format = '0.00'

    # UNIT
    cell = ws.cell(row=r, column=7, value=item['unit'])
    cell.font = normal_font; cell.alignment = center; cell.border = thin_border; cell.fill = row_fill

    # U_PRICE
    cell = ws.cell(row=r, column=8, value=item['price'])
    cell.font = normal_font; cell.alignment = right_align; cell.border = thin_border; cell.fill = row_fill
    cell.number_format = '#,##0.00'

    # AMOUNT
    cell = ws.cell(row=r, column=9, value=item['amt'])
    cell.font = normal_font; cell.alignment = right_align; cell.border = thin_border; cell.fill = row_fill
    cell.number_format = '#,##0.00'

# === Per-invoice subtotal rows ===
current_row = row + len(all_rows)

# Insert subtotal for each invoice group
for inv in invoices:
    r = current_row
    ws.merge_cells(f'A{r}:G{r}')
    label = f"{inv['date']} ({inv['invoice_no']})"
    if inv['note']:
        label += f" [{inv['note']}]"

    cell = ws.cell(row=r, column=1, value=f"SUB-TOTAL  {label}")
    cell.font = header_font; cell.alignment = right_align
    cell.border = thin_border; cell.fill = subtotal_fill
    for c in range(2, 8):
        ws.cell(row=r, column=c).border = thin_border
        ws.cell(row=r, column=c).fill = subtotal_fill

    cell = ws.cell(row=r, column=8, value='')
    cell.border = thin_border; cell.fill = subtotal_fill

    cell = ws.cell(row=r, column=9, value=inv['subtotal'])
    cell.font = header_font; cell.alignment = right_align
    cell.border = thin_border; cell.number_format = '$#,##0.00'
    cell.fill = subtotal_fill
    current_row += 1

    # GST row
    r = current_row
    ws.merge_cells(f'A{r}:G{r}')
    cell = ws.cell(row=r, column=1, value=f"GST 9%  {label}")
    cell.font = header_font; cell.alignment = right_align
    cell.border = thin_border; cell.fill = subtotal_fill
    for c in range(2, 8):
        ws.cell(row=r, column=c).border = thin_border
        ws.cell(row=r, column=c).fill = subtotal_fill
    cell = ws.cell(row=r, column=8, value='')
    cell.border = thin_border; cell.fill = subtotal_fill
    cell = ws.cell(row=r, column=9, value=inv['gst'])
    cell.font = header_font; cell.alignment = right_align
    cell.border = thin_border; cell.number_format = '$#,##0.00'
    cell.fill = subtotal_fill
    current_row += 1

    # Total row
    r = current_row
    ws.merge_cells(f'A{r}:G{r}')
    cell = ws.cell(row=r, column=1, value=f"TOTAL  {label}")
    cell.font = Font(name='Arial', size=11, bold=True); cell.alignment = right_align
    cell.border = thin_border; cell.fill = total_fill
    for c in range(2, 8):
        ws.cell(row=r, column=c).border = thin_border
        ws.cell(row=r, column=c).fill = total_fill
    cell = ws.cell(row=r, column=8, value='')
    cell.border = thin_border; cell.fill = total_fill
    cell = ws.cell(row=r, column=9, value=inv['total'])
    cell.font = Font(name='Arial', size=11, bold=True); cell.alignment = right_align
    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='double'))
    cell.number_format = '$#,##0.00'; cell.fill = total_fill
    current_row += 1

    # Blank separator
    current_row += 1

# === Grand Total ===
r = current_row
ws.merge_cells(f'A{r}:G{r}')
cell = ws.cell(row=r, column=1, value='总合计 GRAND TOTAL')
cell.font = Font(name='Arial', size=12, bold=True); cell.alignment = right_align
cell.border = thin_border; cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
for c in range(2, 8):
    ws.cell(row=r, column=c).border = thin_border
    ws.cell(row=r, column=c).fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
cell = ws.cell(row=r, column=8, value='')
cell.border = thin_border
cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

grand_total = round(sum(inv['total'] for inv in invoices), 2)
cell = ws.cell(row=r, column=9, value=grand_total)
cell.font = Font(name='Arial', size=12, bold=True); cell.alignment = right_align
cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                     top=Side(style='thin'), bottom=Side(style='double'))
cell.number_format = '$#,##0.00'
cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

# Print setup
ws.page_setup.orientation = 'landscape'
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1

# Freeze header
ws.freeze_panes = 'A5'

# Auto filter
ws.auto_filter.ref = f'A4:I{row + len(all_rows) - 1}'

output_path = '/home/user/JoeyZzz/invoice_summary.xlsx'
wb.save(output_path)
print(f'Summary saved to {output_path}')
