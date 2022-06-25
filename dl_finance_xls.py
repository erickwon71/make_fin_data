# -*- coding: utf-8 -*-

from dart_interface import *

#
# 아직 직접 xls file을 download할 수 있는 방법을 못찾아서 download link url을 Print하면 이것으로 download
def download_xls_files(corp_name):
    corp_code = dart_get_corp_code('가비아')
    print(corp_code)
    per_list, rcept_list, dcm_list = dart_get_rcept_dcm_no(corp_code)
    print(per_list, rcept_list, dcm_list)
    dart_download_xls_list(per_list, rcept_list, dcm_list)

if __name__=="__main__":
    print(__name__)
    download_xls_files('가비아')