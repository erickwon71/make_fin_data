# -*- coding: utf-8 -*-

import os
import re
import unicodedata
import pandas as pd


def make_item_info(txt_f):
    item_list = []
    
    #print(txt_f)
    f = open(txt_f, 'r')
    while True:
        line = f.readline().strip()
        if not line:
            break
        #print(line)
        item_list.append(line)
    f.close()
  
    return item_list

if __name__=="__main__":
    print(__name__)
    
    txt_file = './resource/' + 'sf.txt'
    
    sf_list = make_item_info(txt_file)
    for item in sf_list:
        print(item)
