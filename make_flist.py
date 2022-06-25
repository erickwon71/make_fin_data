# -*- coding: utf-8 -*-

import os
import unicodedata
import datetime

def make_report_str(date_data):
    y = date_data.year
    m = date_data.month
    
    if m >= 4 and m <= 5:
        m_str = 'Q1'
    elif m >= 6 and m <= 8:
        m_str = "Q2"
    elif m >= 9 and m <= 11:
        m_str = "Q3"
    else:
        y = y - 1
        m_str = "Q4"
    
    rep_str = '%d%s' %(y, m_str)
    
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
            flist_get[dat] = xlss
    
    # sort flist 
    flist_sorted = {}
    for d in sorted(set(flist_get)):
        flist_sorted[d] = flist_get[d]
        #print(d, flist_sorted[d])

    # convert date to str (ex. 2020.05.11 to 2020Q1)
    # flist_ret = {'2020Q4', '~~~.xls', '2021Q1', 'xxx.xls'}
    flist_ret = {}
    for d in flist_sorted:
        date_data = datetime.datetime.strptime(d, '%Y.%m.%d')
        rep_str = make_report_str(date_data)
        flist_ret[rep_str] = flist_sorted[d]
        #print(d, date_data, date_data.year, date_data.month, rep_str)

    return flist_ret

if __name__=="__main__":
    print(__name__)
    print(make_corp_flist('가비아', './findata/'))