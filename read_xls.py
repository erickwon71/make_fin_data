# -*- coding: utf-8 -*-

import os
import re
import unicodedata
import pandas as pd

# 연결 데이터가 있는지 확인
def check_consolidated_fs(src_file):
    ret_value = False
    
    xl = pd.ExcelFile(src_file)
    #print(xl.sheet_names)
    for sh_name in xl.sheet_names:
        if sh_name == '연결 재무상태표' or sh_name == '연결 포괄손익계산서' or sh_name == '연결 자본변동표' or sh_name == '연결 현금흐름표':
            ret_value = True

    return ret_value    

def check_xls(src_file):
    
    if os.path.isfile(src_file):
        if os.path.exists(src_file):
            return True
        else:
            print(src_file, "is not exist.")
            return False
    else:
        print(src_file, "is not a file.")
        return False

def check_sheet(src_file, sheet_name):
    xl = pd.ExcelFile(src_file)
    #print(xl.sheet_names)
    for sh in xl.sheet_names:
        if sh == sheet_name:
            return True
    
    return False

def read_xls_sheet(src_file, sh_name):
    print(src_file, sh_name)
    
    # file 있는지 확인
    if check_xls(src_file) == False:
        return    
    
    # engine type check
    # engine = openpyxl -> xlsx file사용할때
    # default로는 xlrd (xls file일때)
    ext = os.path.splitext(src_file)[1]
    if ext == '.xlsx':
        engin_name = 'openpyxl'
    else:
        engin_name = 'xlrd'
    
    df = pd.read_excel(src_file, sheet_name=sh_name, names=['item', 'value'], usecols='A:B', engine=engin_name)
    return df

# 전체 data의 원 단위를 맞추기 위해 file에서 사용한 단위를 확인하고 
# 전체 data의 원 단위는 억원으로 맞추려고 함
# 이 함수로 얻은 Unit을 data/unit 과 같이 사용
def get_won_unit(df):
    won_unit = 100000000
    
    tmp_str = df.loc[4, :].to_list()[0]
    won = re.split('[\s:\(\)]', tmp_str)[4]
    print('단위 :', won)
    if won == '백만원':
        won_unit = 100
    elif won == '천만원':
        won_unit = 10
    elif won == '억원':
        won_unit = 1
    
    return won_unit

def get_value_of_item(df, req_list, unit):
    item_val_dict = {}

    for req_item in req_list:
        for item, value in zip(df.loc[7:, 'item'], df.loc[7:, 'value']):
            item_name = item.replace(" ","")
            #print(req_item, item_name)
            if req_item == item_name:
                if str(type(value)) == "<class 'str'>":
                    #print(item_name, 'str', len(value))
                    item_val_dict[item_name] = value
                else:
                    #print(item_name, 'int', value)
                    item_val_dict[item_name] = value/unit
    
    #print(item_val_dict)
    
    return item_val_dict

if __name__=="__main__":
    print(__name__)
    
    src_file = './findata/' + '[가비아]사업보고서_재무제표(2022.03.22)_ko.xls'
    
    if check_consolidated_fs(src_file):
        sheet_list = ['기본정보', '연결 재무상태표', '연결 포괄손익계산서', '연결 자본변동표', '연결 현금흐름표', '재무상태표', '포괄손익계산서', '자본변동표', '현금흐름표']
        print('연결 재무 데이터가 있음')
    else:
        sheet_list = ['기본정보', '재무상태표', '포괄손익계산서', '자본변동표', '현금흐름표']
        print('별도 재무 데이터만 있음')
    
    con_sf = read_xls_sheet(src_file, '연결 재무상태표')

    unit = get_won_unit(con_sf)

    item_list = ['유동자산', '비유동자산', '유동부채', '비유동부채']
    val_dict = get_value_of_item(con_sf, item_list, unit)
    for k, v in val_dict.items():
        print("{} {}".format(k, v))

 # dataframe에 있는 개별 data 가져오는 방법,
 #    - loc
 #    - to_list() 로 Object를 list로 바꾸고
 #    - replace() 로 문자열에 있는 공백 삭제
 #   item_name = df.loc[7:7, 'item'].to_list()[0].replace(" ","")
 #   item_value = df.loc[val_loc:val_loc, 'value'].to_list()[0]
 #   if item_name == '유동자산':
 #       print(item_name, len(item_name))
 #       item_value = con_sf.loc[7:7, 'value'].to_list()[0]
 #       print(item_value)
    