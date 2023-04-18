# -*- coding: utf-8 -*-
"""
Created on Tue Apr 11 13:19:57 2023

@author: Timo
"""

import os
import pandas as pd

def modifyExcelSheet(file_path, col1_name, col2_name, output_col_name, new_file_path, target_number):
    # Laden Sie die Excel-Datei in einen Pandas DataFrame
    df = pd.read_excel(file_path)
    
    # Multiplizieren Sie col1 mit col2 und speichern Sie das Ergebnis in output_col, aber nur wenn col1 target_number enthält
    df[output_col_name] = df.apply(lambda row: row[col1_name] * row[col2_name] if str(target_number) in str(row[col1_name]) else "", axis=1)
    
    # Entfernen Sie Zeilen, die keine Multiplikation durchgeführt haben
    df.dropna(subset=[output_col_name], inplace=True)
    
    # Erstellen Sie den neuen Speicherpfad mit einem hochzählenden Zähler
    base_path, file_name = os.path.split(os.path.basename(file_path))
    file_name_without_ext, file_ext = os.path.splitext(file_name)
    new_file = os.path.join(new_file_path, base_path + '_modified' + file_ext)
    i = 0
    while os.path.exists(new_file):
        i += 1
        new_file_name = f"{file_name_without_ext}_{i}{file_ext}"
        new_file = os.path.join(new_file_path, new_file_name)
    
    # Speichern Sie den DataFrame in eine neue Excel-Datei
    with pd.ExcelWriter(new_file) as writer:
        df.to_excel(writer, index=False)
    
    return new_file
