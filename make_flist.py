# -*- coding: utf-8 -*-

import os
import unicodedata
import datetime

def make_report_str(d_data, filename):
    date_data = datetime.datetime.strptime(d_data, '%Y.%m.%d')    
    y = date_data.year
    m = date_data.month
    
    if '사업보고서' in filename:
        y = y - 1
        m_str = 'Q4'
    elif '반기보고서' in filename:
        m_str = 'Q2'
    else: #분기보고서  
        if m >= 4 and m <= 6:
            m_str = 'Q1'
        else: # m >= 9 and m <= 11:
            m_str = "Q3"
    
    rep_str = '%d%s' %(y, m_str)
    #print(rep_str, filename)
    
    return rep_str

def make_corp_flist(corp_name_req, finpath):
    print(corp_name_req, finpath)
    
    file_list = os.listdir(finpath)
    file_list_xls = [file for file in file_list if file.endswith(".xls")]

    # make dict flist for corp_name
    # unicodedata for 한글자모합치기
    flist_get = {}
    for xlss in file_list_xls:
        fname = unicodedata.normalize('NFC', os.path.splitext(os.path.basename(xlss))[0])
        corp_name_found = fname.split(']')[0].split('[')[1]
        #print(corp_name_found, len(corp_name_found))
        if corp_name_req == corp_name_found:
            dat = fname.split('(')[1].split(')')[0]
            flist_get[dat] = fname

    # convert date to str (ex. 2020.05.11 to 2020Q1)
    # flist_tmp = {'2020Q4', '~~~.xls', '2021Q1', 'xxx.xls'}
    flist_tmp = {}
    for d in flist_get:
        rep_str = make_report_str(d, flist_get[d])
        flist_tmp[rep_str] = flist_get[d]
    
    # sort flist
    flist_sorted = {}
    for d in sorted(set(flist_tmp)):
        flist_sorted[d] = flist_tmp[d]
        #print(d, flist_sorted[d])

    return flist_sorted

if __name__=="__main__":
    print(__name__)
    fin_xls_list = make_corp_flist('가비아', './findata/')