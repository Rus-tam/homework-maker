# -*- coding: utf-8 -*-
"""
Created on Fri Oct 21 22:07:19 2022

@author: Rustam
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font

class Writer:
    def __init__(self):
        pass
    
    @staticmethod
    def make_excel(info1, info2):
        writer = pd.ExcelWriter('data.xlsx', engine='openpyxl')
        workbook = writer.book
        
        df1 = pd.DataFrame(info1)
        df2 = pd.DataFrame(info2)
        
        df1.to_excel(writer, sheet_name='Heat Exchangers', startrow=0, startcol=0, index=False)
        df2.to_excel(writer, sheet_name='Stream Info', startrow=0, startcol=0, index=False)
        
        workbook.save('data.xlsx')
        
        