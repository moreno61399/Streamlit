# -*- coding: utf-8 -*-
"""
Created on Mon Oct 19 11:56:46 2020
@author: Aitor.vidart
"""


import pandas as pd
import numpy as np
import os
from tkinter import filedialog
from tkinter import *

def get_Sesam_Doc(km_path):
    
    file_name = os.path.basename(km_path)
    
    Position = file_name.find('D')
    
    if Position ==-1:
        sesam_doc = 'DXXXXXX'
    else:
        sesam_doc = file_name[Position:Position+8]
    
    return sesam_doc


def load_df_KM (path):
    
    try:
        df_KM = pd.read_excel(path,sheet_name = 'KM_Liste', engine ='openpyxl')
    except:
        df_KM = pd.read_excel(path,sheet_name=0, engine ='openpyxl')
    
    return df_KM

def filter_by_HandDrive(km_folder_path, guide='ALL'):
    

    list_km_files = os.listdir(km_folder_path)

    list_km_files_filtered = []
    
    for km_name in list_km_files:
        
        if guide =='ALL':
            list_km_files_filtered.append(os.path.join(km_folder_path, km_name))

        elif guide == 'LL':
            if any(x in km_name for x in ['LL','L0L','LOL', 'ALL']):
                list_km_files_filtered.append(os.path.join(km_folder_path, km_name))
        elif guide == 'RL':
            if any(x in km_name for x in ['RL','L0R','LOR','LR', 'ROL', 'ALL']):
                list_km_files_filtered.append(os.path.join(km_folder_path, km_name))

    return list_km_files_filtered


#concatena todos los módulos activos de diferentes KM_Listes
def get_dict_Bauraum_df_KMs(list_km_files_filtered, Status=None):

    ''' the date introduced by parameter is the filter to
    return a data frame with corresponding values 
    (EOP values and values that are within a range 
     with the parameter reference)'''

    dict_km_df_KM = dict()
    for km_path in list_km_files_filtered:
        df_KM = load_df_KM(km_path)
        if Status!=None:
            df_KM_aktuell = filter_by_Status(df_KM,Status)
        else:
            df_KM_aktuell = df_KM
            
        Sesam_doc = get_Sesam_Doc(km_path)
        BauRaum = get_BauRaum(km_path)
        HandDrive = get_HandDrive(km_path)
        dict_km_df_KM[Sesam_doc + '_' + HandDrive + '_' + BauRaum] = df_KM_aktuell

    return dict_km_df_KM

def get_df_moduls_aktuell_all_KMs(dict_km_df_KM):
    
    df_moduls_aktuell_all_KMs = pd.DataFrame()
    
    for km_type,df_km in dict_km_df_KM.items():
           
        km_Sesam_doc = km_type.split('_')[0]
        km_HandDrive = km_type.split('_')[1]
        km_Bauraum = km_type.split('_')[2]
                
        df_KM_moduls_aktuell = get_moduls_aktuell(df_km)
        
        df_KM_moduls_aktuell['Sesam Doc']=km_Sesam_doc        
        df_KM_moduls_aktuell['BauRaum']=km_Bauraum
        df_KM_moduls_aktuell['HandDrive']=km_HandDrive        
        
        df_moduls_aktuell_all_KMs = pd.concat([df_moduls_aktuell_all_KMs,df_KM_moduls_aktuell])
    
    df_moduls_aktuell_all_KMs = df_moduls_aktuell_all_KMs.rename(columns={df_moduls_aktuell_all_KMs.columns[0]:'ModuleNoOem',df_moduls_aktuell_all_KMs.columns[1]:'ModuleNoSEBN'})

    return df_moduls_aktuell_all_KMs

def get_df_Baugruppe_all_KMs(dict_km_df_KM):
    
    df_Baugruppe_all_KMs = pd.DataFrame()
    
    for km_type,df_KM in dict_km_df_KM.items():
        
        Sesam_doc = km_type.split('_')[0]
        HandDrive = km_type.split('_')[1]
        BauRaum = km_type.split('_')[2]

        if HandDrive =='ALL':
            HandDrive='LL/RL'
        
        df_of_interest = df_KM.iloc[:,[17,24,25,26]].dropna(how='all').reset_index(drop=True)

        df_of_interest.iloc[:,1].where(df_of_interest.iloc[:,1]=='-','A' + df_of_interest.iloc[:,0].str[1:],inplace=True)
        df_of_interest.iloc[:,2].where(df_of_interest.iloc[:,2]=='-','B' + df_of_interest.iloc[:,0].str[1:],inplace=True)
        df_of_interest.iloc[:,3].where(df_of_interest.iloc[:,3]=='-','M' + df_of_interest.iloc[:,0].str[1:],inplace=True)
        
        df_KM_Baugruppe = df_of_interest.melt(id_vars =['MV intern'], value_name = 'Variant').drop(['variable'],axis=1)
        df_KM_Baugruppe.insert(0,'Lenkerart',HandDrive)
        df_KM_Baugruppe.insert(1,'Bauraum',BauRaum)
        df_KM_Baugruppe.insert(2,'Zeichnung',Sesam_doc)

        df_Baugruppe_all_KMs = pd.concat([df_Baugruppe_all_KMs,df_KM_Baugruppe])
        
    df_Baugruppe_all_KMs = df_Baugruppe_all_KMs[df_Baugruppe_all_KMs['Variant']!='-']
    df_Baugruppe_all_KMs=df_Baugruppe_all_KMs.rename(columns={'MV intern':0})
    df_Baugruppe_all_KMs=df_Baugruppe_all_KMs[['Lenkerart','Bauraum','Zeichnung','Variant',0]]
        
    return df_Baugruppe_all_KMs


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
    df_KM_aktuell['Entfall'] = df_KM_aktuell['Entfall'].replace('EOP',9999)
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
    
    df_moduls_aktuell = df_KM_aktuell.iloc[:,[16,17]].reset_index(drop=True).dropna(how='all')
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
    options_Vorderwagen = ['Vorderwagen','VORDERWAGEN', 'VoWa','VDW']
    
    file_path = os.path.basename(path)
    
    if any(x in file_path.upper() for x in options_INRA):
        BauRaum = 'INRA'
    elif any(x in file_path.upper() for x in options_MORA):
        BauRaum = 'MORA'
    elif any(x in file_path.upper() for x in options_COCKPIT):
        BauRaum = 'COCKPIT'
    elif any(x in file_path.upper() for x in options_FGSR):
        BauRaum = 'FGST'
    elif any(x in file_path.upper() for x in options_Fahrwerk):
        BauRaum = 'FAHRW'
    elif any(x in file_path.upper() for x in options_Tueren):
        BauRaum = 'TUEREN'
    elif any(x in file_path.upper() for x in options_Vorderwagen):
        BauRaum = 'VWAGEN'
    else:
        print('Caution! Bauraum for',path.split('/')[-1],'not found')
        BauRaum = 'not found'
    
    return BauRaum


def get_HandDrive(path):
    
    OPTIONS_LL = ['LL','L0L','LOL']
    OPTIONS_RL = ['RL','L0R','LOR','LR','ROL']
    
    
    file_path = os.path.basename(path)
    
    if any(x in file_path.upper() for x in OPTIONS_LL) and any(x in file_path.upper() for x in OPTIONS_RL):
        HandDrive = 'LL/RL'
    elif any(x in file_path.upper() for x in OPTIONS_LL):
        HandDrive = 'LL'
    elif any(x in file_path.upper() for x in OPTIONS_RL):
        HandDrive = 'RL'
    else:
        print('Caution! Hand Drive for',path.split('/')[-1],'not found')
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

def load_km_file(km_file_path=None):
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
    if km_file_path is None:
        try:
            km_file_path = filedialog.askopenfilename(title="Select the km file",\
                                              filetypes=[("Excel files","*.xls?")])
        except Exception as error:
            print("Error loading the orders file: <{}>".format(error))
    
    return km_file_path


def load_folder(km_folder_path=None):
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
    
    root = Tk()
    root.withdraw()
    
    if km_folder_path is None:
    
        try:
            km_folder_path = filedialog.askdirectory(title="Select the folder where the km files are located")
        except Exception as error:
            print("Error loading the orders folder: <{}>".format(error))
    
    return km_folder_path
    
def filter_by_HandDrive(km_folder_path, guide):
 

    list_km_files = os.listdir(km_folder_path)
    
    list_km_files_filtered = []
    
    for km_name in list_km_files:
    
        if guide =='ALL':
            list_km_files_filtered.append(os.path.join(km_folder_path, km_name))

        elif guide == 'LL':
            if any(x in km_name for x in ['LL','L0L','LOL', 'ALL']):
                list_km_files_filtered.append(os.path.join(km_folder_path, km_name))
        elif guide == 'RL':
            if any(x in km_name for x in ['RL','L0R','LOR','LR', 'ROL', 'ALL']):
                list_km_files_filtered.append(os.path.join(km_folder_path, km_name))

    return list_km_files_filtered