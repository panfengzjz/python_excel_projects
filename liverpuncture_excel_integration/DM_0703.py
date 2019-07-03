# -*- coding: utf-8 -*-
import pandas as pd
from random import shuffle

filename = u'曲阳社区最终筛选名单-1128整理版.xlsx'
new_name = 'new.xlsx'

if 1:
    #保存每个类别符合要求的行号
    male_35_44 = []
    male_45_54 = []
    male_55_64 = []
    male_65_74 = []
    female_35_44 = []
    female_45_54 = []
    female_55_64 = []
    female_65_74 = []
    
    #每个类别允许标记的最大个数
    num_male_35_44 = 6
    num_male_45_54 = 16
    num_male_55_64 = 54
    num_male_65_74 = 60
    num_female_35_44 = 6
    num_female_45_54 = 18
    num_female_55_64 = 66
    num_female_65_74 = 65

    #每个类别对应的备注号
    dic = {'male_35_44': 1,
           'male_45_54': 2,
           'male_55_64': 3,
           'male_65_74': 4,
           'female_35_44': 5,
           'female_45_54': 6,
           'female_55_64': 7,
           'female_65_74': 8,
           }

def insert_result(lst_name, rem_col):
    index_lst = []    
    lst = eval(lst_name)
    shuffle(lst)
    for i in range(min(len(lst), eval("num_%s" %lst_name))):
        rem_col[lst[i]] = dic[lst_name]

def main():
    src_data = pd.read_excel(filename)
    gender_col = src_data['性别']
    age_col = src_data['年龄']

    for i in range(len(gender_col)):
        if (gender_col[i] == '男'):
            age = age_col[i]
            if (35 <= age <= 44):
                male_35_44.append(i)
            elif (45 <= age <= 54):
                male_45_54.append(i)
            elif (55 <= age <= 64):
                male_55_64.append(i)
            elif (65 <= age <= 74):
                male_65_74.append(i)
        elif (gender_col[i] == '女'):
            age = age_col[i]
            if (35 <= age <= 44):
                female_35_44.append(i)
            elif (45 <= age <= 54):
                female_45_54.append(i)
            elif (55 <= age <= 64):
                female_55_64.append(i)
            elif (65 <= age <= 74):
                female_65_74.append(i)
    remark_col = src_data['备注'].copy()

    insert_result('male_35_44', remark_col)
    insert_result('male_45_54', remark_col)
    insert_result('male_55_64', remark_col)
    insert_result('male_65_74', remark_col)
    insert_result('female_35_44', remark_col)
    insert_result('female_45_54', remark_col)
    insert_result('female_55_64', remark_col)
    insert_result('female_65_74', remark_col)

    src_data['备注'] = remark_col
    src_data.to_excel(new_name)

if __name__ == "__main__":
    main()
    print( "parse ok")