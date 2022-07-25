# -*- coding: utf-8 -*-

import os

from make_flist import *
from make_acct_list import *
from read_xls import *

def main(corp_name, res_folder, in_folder, out_folder):
    print(corp_name, in_folder, out_folder)
    
    #step 1. 재무 xls file list 만들기
    src_file_dict = make_corp_f_dict(corp_name, in_folder)
    
    #step 2. sf, si, sc list 만들기 
    sf_list = make_item_info(res_folder+'sf.txt')
    sc_list = make_item_info(res_folder+'sc.txt')
    si_list = make_item_info(res_folder+'si.txt')

    #step 2. file에서 data 추출
    for period in src_file_dict:
        src_file = in_folder+src_file_dict[period]
        print(period, src_file)

        sh_list, consh_list = get_sheet_list(src_file)

        # 별도
        for sh in sh_list:
            con_df = read_xls_sheet(src_file, sh)
            unit = get_won_unit(con_df)

            if sh == '연결 재무상태표' or sh == '재무상태표' or sh == '대차대조표':
                sh_type = '재무'
                req_accts = sf_list
            elif sh == '연결 손익계산서' or sh == '연결 포괄손익계산서' or sh == '손익계산서' or sh == '포괄손익계산서':
                sh_type = '손익'
                req_accts = si_list
            else: #sh == '연결 현금흐름표' or sh == '현금흐름표':
                sh_type = '현금'
                req_accts = sc_list

            print(sh_type)
            val_dict = get_value_of_item(con_df, sh_type, req_accts, unit)
            for k, v in val_dict.items():
                print("{} {}".format(k, v))
        break

        # # 연결
        # for sh in consh_list:
        #     con_df = read_xls_sheet(src_file, sh)
        #     unit = get_won_unit(con_df)

        #     if sh == '연결 재무상태표' or sh == '재무상태표' or sh == '대차대조표':
        #         sh_type = '재무'
        #         req_accts = sf_list
        #     elif sh == '연결 손익계산서' or sh == '연결 포괄손익계산서' or sh == '손익계산서' or sh == '포괄손익계산서':
        #         sh_type = '손익'
        #         req_accts = si_list
        #     else: #sh == '연결 현금흐름표' or sh == '현금흐름표':
        #         sh_type = '현금'
        #         req_accts = sc_list

        #     val_dict = get_value_of_item(con_df, sh_type, req_accts, unit)
        #     for k, v in val_dict.items():
        #         print("{} {}".format(k, v))


    

if __name__=="__main__":
    print(__name__)
    main('가비아', './resource/', './findata/', './result/')