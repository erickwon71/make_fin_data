# -*- coding: utf-8 -*-

import os

from make_flist import *

def main(corp_name, in_folder, out_folder):
    print(corp_name, in_folder, out_folder)
    
    src_list = make_corp_flist(corp_name, in_folder)
    
    for item in src_list:
        print(item, src_list[item])

if __name__=="__main__":
    print(__name__)
    main('가비아', './findata/', './result/')