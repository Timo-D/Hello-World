# -*- coding: utf-8 -*-
"""
Created on Tue Apr 11 13:20:12 2023

@author: Timo
"""
import random

file_path = 'C:/Users/Timo/Desktop/Python_Excel/zufallsdaten.xlsx'
col1_name = 'Spalte 1'
col2_name = 'Spalte 2'
output_col_name = 'Spalte 20'
new_file_path = 'C:/Users/Timo/Desktop/Python_Excel/output'
target_number = str(random.randint(0, 9))

new_file = modifyExcelSheet(file_path, col1_name, col2_name, output_col_name, new_file_path, target_number)
    
    