# 调用库
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Border, Side
from openpyxl.styles import Alignment

#准备数据
#从数据源获取的其他信息
other_data={
    'C4' : '002LCNB250350001',
    'D5' : 'EC-202501-STEG-LC-COND-L6-B2 DATED 31-01-2025',
    'H9' : 'INV-STEG-LC-2025-001',
    'H10' : '19-02-2025',
    'B11' : 'STEG INTERNATIONAL SERVICES',
    'G11' : '122-855-929',
    'G12' : '40-018812-P',
}
#从数据源获取的产品信息
items = {
    'item1' : {
        'SN': 1,
        'description' : 'LV ABC CABLE 4*95 sqmm',
        'unit':'KM',
        'qty': 78,
        'unit-price': 14295000.00
    },
    'item2' : {
        'SN': 2,
        'description' : 'LV ABC CABLE 4*50 sqmm',
        'unit':'KM',
        'qty': 30,
        'unit-price': 8074000.00
    },
    'item3': {
        'SN': 3,
        'description': 'LV ABC CABLE 4*25 sqmm',
        'unit': 'KM',
        'qty': 30,
        'unit-price': 123.00
    }
}




#1.写入其他数据
def input_data_others(other_data,other_cells):
    pass
#2. 根据len(details_data)确认需要插入的行数,并插入行
def insert_rows(starting_row,rows):
    pass
#3. 获取插入的所有单元格对象,return一个包含所有单元格对象的列表
def get_cells_obj(starting_row,rows):
    pass
#4. 将列表中单元格对象分类:需要填入details_data的列表,需要填入公式的列表
def slic_list(original_list):
    pass
#5. 处理detailed_data, 返回一个2元列表
def get_values_from_dic(details_data):
    pass
#6. 为单元格填入数据
def insert_value(cells_obj_list,value_list):
    pass
#7. 为单元格添加公式
def insert_formula(formula_obj_list,formula):
    pass
#8. 为单元格添加边框
def add_border(cells_obj_list):
    pass