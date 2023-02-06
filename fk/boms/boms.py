# -*- coding: utf-8 -*-
"""
Created on Tue Jun  1 08:48:13 2021

@author: jesus.roldan
"""

import pandas as pd


def delete_red_striked_boms(ws):
    
    '''
        Removes all the rows/columns with red and striked values
    '''
    hoja = ws.title
    
    #en primer lugar buscamos la última linea y eliminamos en un paso todas las filas de ahi en adelante
    maximum_row = max ((c.row for c in ws['B'] if c.value is not None))
    ws.delete_rows(maximum_row + 1, ws.max_row)
    
    #En boms solo recorremos filas basándonos en la columna B (Modul family)
    cont = 0
    for cell in ws['B']:
        
        fila = cell.row
        
        if cell.value!=None:
            if cell.font.strike == True:
                if cell.font.color == None or (cell.font.color and cell.font.color.rgb !='FFFF0000'):   #el texto está tachado y tiene color rojo
                    print('Cuidado! Fila ' + str(fila + cont) + ' en hoja ' + str(hoja) + ' está tachada pero no en rojo y ha sido eliminada!')
                ws.delete_rows(fila)
                cont = cont+1
    return ws

def check_Baugruppe(df_bom):
    
    #seleccionamos las columnas de interés
    list_cols = [x for x in df_bom.columns if str(x).find('Baugruppe')!=-1 or str(x).find('Verbauort')!=-1 or str(x).find('Kurzname')!=-1]
    
    Verbau = list_cols[0]
    Bau = list_cols[1]
    Kurz = list_cols[2]
    
    #creamos el dataframe solo con estas columnas
    df = df_bom[list_cols].dropna(how='all').drop_duplicates().reset_index(drop=True)
    
    df = df[(~df[Kurz].str.contains('(g)', na = False)) & (~df[Kurz].str.contains(Kurz,na=False))]
    
    df_wrong_Verbauort_1 = df[~(df[Bau].str.contains('A', na=False)) & (df[Verbau].str.contains('VT', na=False) | df[Verbau].str.contains('PIM', na=False))]

    df_wrong_Verbauort_2 = df[~(df[Bau].str.contains('B', na=False)) & (df[Verbau].str.contains('KSK', na=False))]

    df_wrong_Verbauort_3 = df[~(df[Bau].str.contains('M', na=False)) & (df[Verbau].str.contains('KM', na=False))]
    
    Report_Baugruppe_boms = pd.concat([df_wrong_Verbauort_1,df_wrong_Verbauort_2,df_wrong_Verbauort_3],ignore_index=True).drop_duplicates().sort_values(by=[Kurz,Bau])

    Report_Baugruppe_boms = Report_Baugruppe_boms[[Kurz,Verbau,Bau]]
    
    return Report_Baugruppe_boms

