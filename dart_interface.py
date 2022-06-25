# -*- coding: utf-8 -*-

import os
import requests
import zipfile
import re
import time
from io import BytesIO
import xml.etree.cElementTree as ET

API_KEY = '8a1b054bb10967f4c2a5840fb583d0b2aaf1d84e'
OPENDARTAPI_URL = 'https://opendart.fss.or.kr/api/'
CORPCODE_FOLDER = './resource/'
FINANCE_FOLDER = './findata/'
CORPCODE_FILE = 'CORPCODE.xml'

def dart_download_corpcode_file():
    if os.path.isfile(CORPCODE_FOLDER+CORPCODE_FILE) == True:
        print(CORPCODE_FOLDER+CORPCODE_FILE +' exists already.')
        return
    else:
        if not os.path.isdir(CORPCODE_FOLDER):
            os.mkdir(CORPCODE_FOLDER)

    api_url = OPENDARTAPI_URL + 'corpCode.xml'
    res = requests.get(api_url, params={'crtfc_key': API_KEY})
    #utils_check_api_xml_return(res)

    zfile = zipfile.ZipFile(BytesIO(res.content))
    filename = zfile.namelist()[0]
    print("파일명 : " + filename)

    zfile.extractall(CORPCODE_FOLDER)
    if os.path.isfile(filename):
        os.remove(filename)  # 원본 압축파일 삭제

def dart_get_corp_code(name):
    if os.path.isfile(CORPCODE_FOLDER+CORPCODE_FILE) == False:
        print(CORPCODE_FILE +' is not ready.')
        dart_download_corpcode_file()

    tree = ET.parse(CORPCODE_FOLDER+CORPCODE_FILE)  # CORPCODE.xml을 파싱하여 tree에 저장
    root = tree.getroot()

    req_corp_name_len = len(name)
    if req_corp_name_len != 0:
        for company in root.findall('list'):
            company_name = company.find('corp_name').text
            if len(company_name) > req_corp_name_len:
                continue
            if name in company_name:
               return company.find('corp_code').text

    return 'NaN'

# referenced from this blog
# https://blog.naver.com/PostView.nhn?blogId=sleep0615&logNo=222186996431
def dart_get_rcept_dcm_no(corp_code):
    ## 사업보고서
    report = 'A001'
    api_url = OPENDARTAPI_URL + 'list.json?crtfc_key={}&corp_code={}&bgn_de=19900101&pblntf_detail_ty={}&page_no=3&page_count=100'.format(API_KEY, corp_code,report)
    #print(api_url)
    resp = requests.get(api_url)
    webpage = resp.content.decode('utf-8')
    #print(webpage)
    period_list = re.findall(r'report_nm":"(.*?)"', webpage)
    rcept_no_list = re.findall(r'rcept_no":"(.*?)"', webpage)

    # to avoid request abort
    time.sleep(1)

    ## 반기보고서
    report = 'A002'
    api_url = OPENDARTAPI_URL + 'list.json?crtfc_key={}&corp_code={}&bgn_de=19900101&pblntf_detail_ty={}&page_no=3&page_count=100'.format(API_KEY, corp_code,report)
    #print(api_url)
    resp = requests.get(api_url)
    webpage = resp.content.decode('utf-8')
    #print(webpage)
    period_list += re.findall(r'report_nm":"(.*?)"', webpage)
    rcept_no_list += re.findall(r'rcept_no":"(.*?)"', webpage)

    # to avoid request abort
    time.sleep(1)
    
    ## 분기보고서
    report = 'A003'    
    api_url = OPENDARTAPI_URL + 'list.json?crtfc_key={}&corp_code={}&bgn_de=19900101&pblntf_detail_ty={}&page_no=3&page_count=100'.format(API_KEY, corp_code,report)
    #print(api_url)
    resp = requests.get(api_url)
    webpage = resp.content.decode('utf-8')
    #print(webpage)
    period_list += re.findall(r'report_nm":"(.*?)"', webpage)
    rcept_no_list += re.findall(r'rcept_no":"(.*?)"', webpage)

    #print(period_list)
    #print(rcept_no_list)
    dict(zip(period_list, rcept_no_list))

    #dcm number
    dcm_no_list = []
    for rcept_no in rcept_no_list:
        resp_dcm = requests.get("http://dart.fss.or.kr/dsaf001/main.do?rcpNo={}".format(rcept_no))
        webpage = resp_dcm.content.decode('utf-8')
        dcm_no = re.findall(r"{}', '(.*?)',".format(rcept_no), webpage)[0]
        dcm_no_list.append(dcm_no)
        time.sleep(1)
    
    #print(dcm_no_list)
    
    return period_list, rcept_no_list, dcm_no_list

def dart_download_xlr(period_list, rcept_no_list, dcm_no_list):
    for p, r, d in zip(period_list, rcept_no_list, dcm_no_list):
        try:
            url_down = "https://dart.fss.or.kr/pdf/download/excel.do?rcp_no={}&dcm_no={}&lang=ko".format(r, d)
            print(url_down)
            user_agent = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.115 Safari/537.36"
            resp_down = requests.get(url_down, headers={"user-agent" : user_agent}, verify=False)
            time.sleep(1)
        except:
            pass

def dart_download_xls_list(period_list, rcept_no_list, dcm_no_list):
    for p, r, d in zip(period_list, rcept_no_list, dcm_no_list):
        try:
            url_down = "https://dart.fss.or.kr/pdf/download/excel.do?rcp_no={}&dcm_no={}&lang=ko".format(r, d)
            print(url_down)
        except:
            pass


if __name__=="__main__":
    print(__name__)
    corp_code = dart_get_corp_code('가비아')
    print(corp_code)
    per_list, rcept_list, dcm_list = dart_get_rcept_dcm_no(corp_code)
    print(per_list, rcept_list, dcm_list)
    dart_download_xls_list(per_list, rcept_list, dcm_list)