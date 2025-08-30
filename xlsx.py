from openpyxl import load_workbook, Workbook

data={
    'lcNum' : '002LCNB250350001',
    'referenceInvoiceNum' : 'EC-202501-STEG-LC-COND-L6-B2',
    'referenceInvoiceDate' : '31-01-2025',
    'invoiceNum' : 'INV-STEG-LC-2025-001',
    'invoiceDate' : '19-02-2025',
    'clientName' : 'STEG INTERNATIONAL SERVICES',
    'tinNum' : '122-855-929',
    'vrnNum' : '40-018812-P',
    'item1' : {
        'SN': '1',
        'description' : 'LV ABC CABLE 4*95 sqmm',
        'unit':'KM',
        'qty': 78,
        'unit-price': 14295000.00
    },
    'item2' : {
        'SN': '2',
        'description' : 'LV ABC CABLE 4*50 sqmm',
        'unit':'KM',
        'qty': 30,
        'unit-price': 8074000.00
    },
    'amountInWords' : 'not programed'
}
columns = ['C6',
           'C7',
           'D8',
           'H10',
           'H11',
           'B12',
           'G12',
           'G13',
           'A18',
           'B18',
           'F18',
           'H18',
           'H18',
           'C22'
           ]

file_path = 'xlsx/test1.xlsx'
wb = load_workbook(file_path)
ws = wb['Sheet1']
ws.merge_cells('B18:E18')
ws.insert_rows(18)
# for i, (key, value) in enumerate(data.items()):
#     if i >= 8:
#         break
#     ws[columns[i]] = value

wb.save(file_path)
