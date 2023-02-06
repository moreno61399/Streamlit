# -*- coding: utf-8 -*-
"""
Created on Mon Oct 19 11:56:46 2020
@author: Aitor.vidart
"""


import pandas as pd
import numpy as np
from km import open_files_km as openFile


def kameliste(date_variable):
    
    ''' the date introduced by parameter is the filter to
    return a data frame with corresponding values 
    (EOP values and values that are within a range 
     with the parameter reference)'''
    
    file_path = openFile.get_gui_filenames()
    
    df_concatenar = pd.DataFrame()
    df_kameliste = pd.DataFrame()
    
    for lista in file_path:
        for path in lista:
            df_concatenar = pd.read_excel(path)
            df_kameliste = pd.concat([df_kameliste,df_concatenar])
            
    df_kameliste = df_kameliste.iloc[:,[13,14,16,17]].dropna(how='all')
    df_kameliste = df_kameliste.replace(to_replace=['EOP'],value=0).fillna(0)
    df_kameliste = df_kameliste.replace(['-','x'],np.nan)
    df_kameliste = df_kameliste.dropna()
    df_kameliste.columns = ['Einsatz','Entfall','Kunden-teilenummer','interne Teilenummer']
    df_final=df_kameliste[((df_kameliste['Einsatz'].astype(int) <= date_variable) & (df_kameliste['Entfall'].astype(int) >= date_variable)) | ((df_kameliste['Einsatz'].astype(int) <= date_variable) & (df_kameliste['Entfall'].astype(int) == 0))] 

    return df_final


