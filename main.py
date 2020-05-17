import pandas as pd
import difflib


def string_similar(s1, s2):
    return difflib.SequenceMatcher(None, s1, s2).quick_ratio()


file_name = '“以艺交心”活动报名表（收集结果）.xlsx'
sheet = pd.read_excel(file_name)
number = 0
sheet_values = sheet.values
while number < len(sheet_values):
    if pd.isna(sheet_values[number][11]):
        print(sheet_values[number])
        this_term = sheet_values[number]
        pai_m = {}
        number2 = 0
        while number2 < len(sheet_values):
            term = sheet_values[number2]
            if pd.isna(term[11]):
                similar = string_similar(this_term[9] + this_term[8], term[9] + term[8])
                pai_m[number2] = similar
            number2 += 1
        pai_m1 = sorted(pai_m.items(), key=lambda item: item[1], reverse=True)
        del pai_m1[0]
        for index_, info in enumerate(pai_m1[0:5]):
            print('****************************')
            print(index_, info[1])
            print(sheet_values[info[0]][2:])
        couple_num = input('请输入配对序号:')
        if int(couple_num) == 11:
            pass
        else:
            number_list = pai_m1[int(couple_num)][0]
            sheet_values[number_list][11] = this_term[2]
            sheet_values[number][11] = sheet_values[number_list][2]
            print(sheet_values[number])
            print(sheet_values[number_list])
    print('///////////////////')
    number += 1
data = pd.DataFrame(sheet_values)
data.to_excel("data_pai.xls")
