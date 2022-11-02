# -*- coding: utf-8 -*-
"""
Created on Wed Sep 21 21:28:44 2022

@author: Rustam
"""

import win32com.client as win32

class Connector:
    def case(file_path):
        print('---- Подключаюсь к хайсису ----')
        print(' ')
        app = win32.DispatchEx('HYSYS.Application')
        case = app.SimulationCases.Open(file_path)
        case.Visible = 1
        
        return case