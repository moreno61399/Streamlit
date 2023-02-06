# -*- coding: utf-8 -*-
"""
Created on Mon Oct 19 11:56:46 2020
@author: Aitor.vidart
"""


import pandas as pd
import numpy as np
import os
from tkinter import filedialog


#concatena todos los módulos activos de diferentes KM_Listes
def kameliste(lista_KMs, Status):

    ''' the date introduced by parameter is the filter to
    return a data frame with corresponding values 
    (EOP values and values that are within a range 
     with the parameter reference)'''

    df_final = pd.DataFrame()
    
    for path in lista_KMs:
        try:
            df_KM = pd.read_excel(path,sheet_name = 'KM_Liste')
        except:
            df_KM = pd.read_excel(path,sheet_name=0)
        df_KM_aktuell = filter_by_Status(df_KM,Status)
        df_moduls_aktuell = get_moduls_aktuell(df_KM_aktuell)
        df_moduls_aktuell['Bauraum']=get_BauRaum(path) + '_' + get_HandDrive(path)
        df_final = pd.concat([df_final,df_moduls_aktuell])

    return df_final

#genera un dataframe con los módulos válidos dado un estatus técnico
def filter_by_Status(df_KM,Status):
    
    df_KM_aktuell = df_KM.dropna(how='all')
# =============================================================================
#     df_KM_aktuell = df_KM.iloc[:,[13,14,16,17]].dropna(how='all')
# =============================================================================
    df_KM_aktuell = df_KM_aktuell.rename(columns={df_KM_aktuell.columns[13]:'Einsatz',df_KM_aktuell.columns[14]:'Entfall', df_KM_aktuell.columns[16]:'MV extern', df_KM_aktuell.columns[17]:'MV intern'})
# =============================================================================
#     df_KM_aktuell.columns[13,14,16,17] = ['Einsatz','Entfall','Kunden-teilenummer','interne Teilenummer']
# =============================================================================
    df_KM_aktuell = df_KM_aktuell.replace(r"^\s*$",np.nan,regex=True)
    df_KM_aktuell['Entfall'] = df_KM_aktuell['Entfall'].replace(to_replace=['EOP'],value=0).fillna(0)
    df_KM_aktuell['Einsatz'] = df_KM_aktuell['Einsatz'].replace(['-','x'],np.nan)
    df_KM_aktuell['Entfall'] = df_KM_aktuell['Entfall'].replace(['-','x'],np.nan)
    df_KM_aktuell['Einsatz'] = df_KM_aktuell['Einsatz'].fillna(0)
    df_KM_aktuell['Entfall'] = df_KM_aktuell['Entfall'].fillna(9999)
    df_KM_aktuell.loc[:,'Einsatz'] = df_KM_aktuell.loc[:,'Einsatz'].astype(int)
    df_KM_aktuell.loc[:,'Entfall'] = df_KM_aktuell.loc[:,'Entfall'].astype(int)
    df_KM_aktuell=df_KM_aktuell[((df_KM_aktuell['Einsatz'] <= Status) & (df_KM_aktuell['Entfall'] >= Status)) | ((df_KM_aktuell['Einsatz'] <= Status) & (df_KM_aktuell['Entfall'] == 0))] 
    
# =============================================================================
#     df_KM_aktuell=df_KM_aktuell.iloc[2:,:]
# =============================================================================

    return df_KM_aktuell

def get_moduls_aktuell(df_KM_aktuell):
    
    df_moduls_aktuell = df_KM_aktuell.iloc[:,[16,17]].reset_index(drop=True)
# =============================================================================
#     df_moduls_aktuell = df_moduls_aktuell.rename(columns={df_KM_aktuell.columns[0]:'MV extern'})
# =============================================================================
    return df_moduls_aktuell


def get_BauRaum(path):
    options_INRA = ['INRA','IR', 'INNENRAUM','INNEN']
    options_MORA = ['MORA','MR','MOTORRAUM']
    options_COCKPIT = ['COCKPIT']
    options_FGSR = ['FGSR','FGR','FGST','FAHRGASTRAUM']
    options_Fahrwerk = ['Fahrwerk','FAHRWERK']
    options_Tueren = ['Tueren','turen','TUEREN','TUREN']
    options_Vorderwagen = ['Vorderwagen','VORDERWAGEN','VDW']
    
    if any(x in path.upper() for x in options_INRA):
        BauRaum = 'INRA'
    elif any(x in path.upper() for x in options_MORA):
        BauRaum = 'MORA'
    elif any(x in path.upper() for x in options_COCKPIT):
        BauRaum = 'COCKPIT'
    elif any(x in path.upper() for x in options_FGSR):
        BauRaum = 'FGSR'
    elif any(x in path.upper() for x in options_Fahrwerk):
        BauRaum = 'FAHRWERK'
    elif any(x in path.upper() for x in options_Tueren):
        BauRaum = 'TUEREN'
    elif any(x in path.upper() for x in options_Vorderwagen):
        BauRaum = 'VORDERWAGEN'
    else:
        BauRaum = 'not found'
    
    return BauRaum

def get_HandDrive(path):
    OPTIONS_LL = ['LL','L0L','LOL']
    OPTIONS_RL = ['RL','L0R','LOR','LR','ROL']
    
    if (any(x in path.upper() for x in OPTIONS_LL) and any(x in path.upper() for x in OPTIONS_RL)):
        HandDrive = 'ALL'
    
    elif any(x in path.upper() for x in OPTIONS_LL):
        HandDrive = 'LL'
    elif any(x in path.upper() for x in OPTIONS_RL):
        HandDrive = 'MORA'
    else:
        HandDrive = 'not found'
    
    return HandDrive


def get_moduls_aktuell_with_IBG (df_KM_aktuell):
    
    df_moduls_aktuell = get_moduls_aktuell(df_KM_aktuell)
    df_IBG_aktuell= df_KM_aktuell.iloc[:,28:].reset_index(drop=True)
    
    df_moduls_aktuell_with_IBG = pd.concat([df_moduls_aktuell,df_IBG_aktuell],axis=1)
    
    
# =============================================================================
#     df_intern = df_KM.iloc[:,17]
#     #df_IBG = df_KM.iloc[:,28:].dropna(how='all',axis=1)
#     df_IBG = df_KM.iloc[:,28:]
#     #df_concat = pd.concat([df_intern,df_IBG],axis=1).dropna(how='all')
#     df_concat = pd.concat([df_intern,df_IBG],axis=1)
#     df_concat = df_concat.rename(columns={df_concat.columns[0]:'interne Teilenummer'})
#     
#     df_KM_aktuell_with_IBG = df_KM_aktuell.merge(df_concat,on='interne Teilenummer',how='left')
# =============================================================================

    return df_moduls_aktuell_with_IBG

def load_km_file(path=None):
    '''
    Permite al usuario obtener el path del file donde se encuentra la km, si se 
    le pasa por defecto el path, devuelve el mismo dado como entrada.
    
    Parameters
    ----------
    path : string, default=None
    
    Attributes
    ----------
  
    Returns
    -------
    path : String
        Cadena de caracteres que representa el path del file 
    '''
    if path is None:
        try:
            path = filedialog.askopenfilename(title="Select the km file",\
                                              filetypes=[("Excel files","*.xls?")])
        except Exception as error:
            print("Error loading the orders file: <{}>".format(error))
    
    return path


def load_km_folder(path=None):
    '''
    Permite al usuario obtener el directorio donde se encuantran las km.
    Busca en el directorio los file que sean km y los clasifica en 'LL' o 'RL',
    empleando 'options_LL' y 'options_RL' respectivamente, para ello verifica si 
    estas opciones aparecen en la cadena de caracteres que representa el path del las km.     
     
    Parameters
    ----------
    path : string, default=None
    
    Attributes
    ----------
    options_LL : array de String
        Contine las opciones que son válidas para LL
    options_RL: array de String
        Contine las opciones que son válidas para RL
  
    Returns
    -------
    lol_files : array
        lista de km que su guía es LL             
    
    lor_files : array
        lista de km que su guía es RL  
    '''
    
    if path is None:
    
        try:
            path = filedialog.askdirectory(title="Select the folder where the km files are located")
        except Exception as error:
            print("Error loading the orders folder: <{}>".format(error))
    
    lol_files = []
    lor_files = []
    
    options_LL = ['LL','L0L','LOL', 'ALL']
    options_RL = ['RL','L0R','LOR','LR', 'ROL', 'ALL']
    
    files = os.listdir(path)
    
    for file in files:
        file_upp = file.upper()
        if any(x in file_upp for x in options_LL):
            lol_files.append(os.path.join(path, file))
       
        if any(x in file_upp for x in options_RL):
            lor_files.append(os.path.join(path, file))
            
    return lol_files, lor_files
