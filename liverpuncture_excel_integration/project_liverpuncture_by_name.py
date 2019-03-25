#coding: utf-8
import openpyxl
from openpyxl.styles import colors
from openpyxl.styles import PatternFill

# global parameter
patientADict = {}
patientBDict = {}

# 将单元格内变量尽量以 float 类型返回
def get_value(text):
    try:
        return float(text)
    except ValueError:
        return text
    except TypeError:
        return text

# 将住院号写入字典并返回
def make_dict(filename):
    wb = openpyxl.load_workbook(filename)
    sheet = wb[wb.sheetnames[0]]
    new_dict = {}
    max_row = sheet.max_row

    for i in range(1, max_row+1):
        item = sheet.cell(i, 2).value
        new_dict[item] = i
    return new_dict

def read_excel(filename):
    wb = openpyxl.load_workbook(filename)
    sheet = wb[wb.sheetnames[0]]

    for row in sheet.rows:
        for cell in row:
            print(cell.value, "\t", end="")
        print()

def backup_excel(src_name, backup_name):
    wb2 = openpyxl.Workbook()
    wb2.save(backup_name)
    print('新建成功')

    #读取数据
    wb1 = openpyxl.load_workbook(src_name)
    wb2 = openpyxl.load_workbook(backup_name)
    sheet1 = wb1[wb1.sheetnames[0]]
    sheet2 = wb2[wb2.sheetnames[0]]

    max_row = sheet1.max_row        #最大行数
    max_column = sheet1.max_column  #最大列数

    for m in range(1, max_row+1):
        for n in range(1, max_column+1):
            cell1 = sheet1.cell(n, m).value #获取data单元格数据
            sheet2.cell(n, m).value = cell1 #赋值到test单元格

    wb2.save(backup_name)   #保存数据
    wb1.close()
    wb2.close()

# src_name: 原始A表格
# ref_name: 参考表格B，A中没有B中有的信息会被添加到新excel中
# dst_name: 最终生成的新excel名称
def insert_excel(src_name, ref_name, dst_name):
    #pfill_empty = PatternFill(start_color ='FFFF00', 
                              #end_color = 'FFFF00', 
                              #fill_type = 'solid')  # 以黄色填充A表空白单元格
    #pfill_cofil = PatternFill(start_color ='FF4500',
                              #end_color = 'FF4500', 
                              #fill_type = 'solid')  # 以红色填充AB表冲突单元格
    wb_d = openpyxl.Workbook()
    wb_d.save(dst_name)
    print('新建成功')

    #读取数据
    wb_s = openpyxl.load_workbook(src_name)
    wb_r = openpyxl.load_workbook(ref_name)
    wb_d = openpyxl.load_workbook(dst_name)
    sheet_s = wb_s[wb_s.sheetnames[0]]  # src sheet
    sheet_r = wb_r[wb_r.sheetnames[0]]  # ref sheet
    sheet_d = wb_d[wb_d.sheetnames[0]]  # dst sheet

    max_row = sheet_s.max_row       #最大行数
    max_col = sheet_s.max_column    #最大列数

    for m in range(1, max_row+1):
        patient_name = sheet_r.cell(m, 2).value
        if patient_name in patientADict:
            for n in range(1, max_col+1):
                new_row = patientADict[patient_name]
                cell_s = sheet_s.cell(new_row, n).value
                sheet_d.cell(m, n).value = cell_s
        else:
            for n in range(1, max_col+1):
                cell_r = sheet_r.cell(m, n).value
                sheet_d.cell(m, n).value = cell_r

    wb_d.save(dst_name)   #保存数据
    wb_s.close()
    wb_d.close()

if __name__ == "__main__":
    fileAName = "../0325查找数据计划-gua/数据库sample.xlsx"
    fileBName = "../0325查找数据计划-gua/需找到人名sample.xlsx"
    backAName = "../0325查找数据计划-gua/表A-test.xlsx"
    patientADict = make_dict(fileAName)
    #patientBDict = make_dict(fileBName)
    #print(patientBDict)
    insert_excel(fileAName, fileBName, backAName)
