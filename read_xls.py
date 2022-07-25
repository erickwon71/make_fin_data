# -*- coding: utf-8 -*-

import os
import re
import unicodedata
import pandas as pd

from make_acct_list import *

# return 별도, 연결 list
def get_sheet_list(src_file):
    sh_list = []
    consh_list = []
    
    try:    
        xl = pd.ExcelFile(src_file)
        #print(xl.sheet_names)
        for sh in xl.sheet_names:
            if sh == '연결 재무상태표':
                consh_list.append(sh)
            elif sh == '연결 손익계산서' or sh == '연결 포괄손익계산서':
                consh_list.append(sh)
            elif sh == '연결 현금흐름표':
                consh_list.append(sh)
            elif sh == '재무상태표' or sh == '대차대조표':
                sh_list.append(sh)
            elif sh == '손익계산서' or sh == '포괄손익계산서':
                sh_list.append(sh)
            elif sh == '현금흐름표':
                sh_list.append(sh)
    except FileNotFoundError as e:
        print(e, src_file, " is not found.")
    except ValueError as e:
        print(e)
    
    return sh_list, consh_list

def check_consolidated(sh_name):
    #print('check_consolidated')

    if '연경' in sh_name:
        return True
    
    return False

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
    #print(src_file, sh_name)
    df = None
    # file 있는지 확인
    if check_xls(src_file) == False:
        return df
    
    # engine type check
    # engine = openpyxl -> xlsx file사용할때
    # default로는 xlrd (xls file일때)
    ext = os.path.splitext(src_file)[1]
    if ext == '.xlsx':
        engin_name = 'openpyxl'
    else:
        engin_name = 'xlrd'
    
    try:
        df = pd.read_excel(src_file, sheet_name=sh_name, names=['item', 'value'], usecols='A:B', engine=engin_name)
    except FileNotFoundError as e:
        print(e, src_file, " is not found.")
        print(e)
    except ValueError as e:
        print(e)

    return df

# 전체 data의 원 단위를 맞추기 위해 file에서 사용한 단위를 확인하고 
# 전체 data의 원 단위는 억원으로 맞추려고 함
# 이 함수로 얻은 Unit을 data/unit 과 같이 사용
def get_won_unit(df):
    # print('get_won_unit')
    won_unit = 100000000
    won = '원'
    
    # (단위 : 원) 이 문자열을 찾기 (대충 10줄 내에서 찾을 수 있을 것으로 예상하고)
    for r in range(0, 9):
        for c in range(0, 1):
            if type(df.iloc[r, c]) != str:
                continue
            tmp_str = df.iloc[r, c]
            if '단위' in tmp_str:
                # print(tmp_str)
                for sp_str in re.split('[\s:\(\)]', tmp_str):
                    if len(sp_str) > 0:
                        if sp_str == '단위': continue
                        won = sp_str
                break
    
    # print('단위 :', won)
    if won ==     '원': won_unit = 100000000
    elif won ==  '십원': won_unit = 10000000
    elif won ==  '백원': won_unit = 1000000
    elif won ==  '천원': won_unit = 100000
    elif won ==  '만원': won_unit = 10000
    elif won == '십만원': won_unit = 1000
    elif won == '백만원': won_unit = 100
    elif won == '천만원': won_unit = 10
    elif won == '억원':  won_unit = 1
    
    return won_unit

# req_list에 있는 계정을 찾지 못하면 '0'값으로 채울 수 있도록
def get_value_of_item(df, fs_type, req_list, unit):
    item_val_dict = {}

    for req_acct in req_list:
        item_val_dict[req_acct] = 0
        for item, value in zip(df.loc[6:, 'item'], df.loc[6:, 'value']):
            if type(item) is not int and type(item) is not float:
                item_name = item.replace(" ","")
                if req_acct in item_name:
                    # print(req_acct, item, value)
                    if str(type(value)) == "<class 'str'>":
                        #print(item_name, 'str', len(value))
                        item_val_dict[req_acct] = value
                    else:
                        #print(item_name, 'int', value)
                        item_val_dict[req_acct] = value/unit
                    break        
        # print(req_acct, item, value)
       
    return item_val_dict

if __name__=="__main__":
    print(__name__)
    
    src_file = './findata/' + '[가비아]분기보고서_재무제표(2007.11.14)_ko.xls'

    sf_list = make_item_info('./resource/'+'sf.txt')
    sc_list = make_item_info('./resource/'+'sc.txt')
    si_list = make_item_info('./resource/'+'si.txt')
    
    sh_list, consh_list = get_sheet_list(src_file)
    print(sh_list)
    
    #    
    for sh in sh_list:
        con_sf = read_xls_sheet(src_file, sh)
        unit = get_won_unit(con_sf)

        if sh == '연결 재무상태표' or sh == '재무상태표' or sh == '대차대조표':
            sh_type = '재무'
            item_list = sf_list
        elif sh == '연결 손익계산서' or sh == '연결 포괄손익계산서' or sh == '손익계산서' or sh == '포괄손익계산서':
            sh_type = '손익'
            item_list = si_list
        else: #sh == '연결 현금흐름표' or sh == '현금흐름표':
            sh_type = '현금'
            item_list = sc_list

        val_dict = get_value_of_item(con_sf, sh_type, item_list, unit)
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
    