# -*- coding: utf-8 -*-
"""
Created on Tue Jan 12 11:18:13 2021

@author: jesus.roldan
"""
import pandas as pd

def remove_rows_with_string(df,string):
    list_of_rows = get_rows_where_string(df,string)
    if len(list_of_rows)!=0:
        df =df.drop(labels=list_of_rows)
    return df

def get_rows_where_string(df,string):
    
    set_of_rows =set()
    
    for col in range(0,len(df.columns)):
        
        fila=df[df.iloc[:,col]==string].index.values
        if len(fila)!=0:
            set_of_rows.add(fila[0])
    list_of_rows = list(set_of_rows)
    list_of_rows.sort()
    
    return list_of_rows

def get_cols_where_string(df,string):
    
    set_of_cols =set()
    
    for col in range(0,len(df.columns)):
        
        if string in df.iloc[:,col].tolist():
            set_of_cols.add(col)
    list_of_cols = list(set_of_cols)
    list_of_cols.sort()
    return list_of_cols


def find_index_column(df,name_column):
    '''Buscamos la posicion de una columna en concreto'''
    df_headers_2 = pd.DataFrame(df.columns)
    tupla_posicion = buscar_string(df_headers_2,name_column)
    posicion_column = tupla_posicion[0][0]
    return posicion_column


def buscar_string(df,string):
    
    lista_tuplas=[]
    for col in range(0,len(df.columns)):
        
        indices=df[df.iloc[:,col]==string].index.values
        
        for  i in indices:
            lista_tuplas.append((i,col))
            
            
    return lista_tuplas


def buscar_substring(df,substring):
    
    lista_tuplas=[]
    for col in range(0,len(df.columns)):
        
        indices=df[df.iloc[:,col].astype(str).str.contains(substring)].index.values
        
        for  i in indices:
            lista_tuplas.append((i,col))
            
            
    return lista_tuplas


def extract_ca(df, lista):
    '''
        Taking the vt name as key, the vt DataFrame and a list of vt column names,
        it returns a DataFrame with those columns.
    '''

    df = df[lista].reset_index(drop=True)
    
    return df

#convertimos un dataframe de 2 columnas en un diccionario
def convert_df_to_dict(df):

    columns = df.columns

    df2 =df.copy()

    df2.loc[:,columns[1]]=df2.loc[:,columns[1]].map(lambda x: list(x.split(',')))


# =============================================================================
#     df[columns[1]].map(lambda x: [x])
# =============================================================================
    
    
# =============================================================================
#     for row in range(0,len(df[columns[1]])):
#         df.loc[row,columns[1]] = [df.loc[row,columns[1]]]
# =============================================================================
    
    df2 = df2.groupby(columns[0]).agg({columns[1]: 'sum'})
    
    dict = df2.T.to_dict('records')[0]
    #dict_modul_Num_aux2 = df.set_index('MV intern')['Nummerierung'].to_dict()
    
    return dict

def convert_df_to_wb(df):
    
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl import Workbook
    
    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    
    return wb,ws

def add_df_to_wb(wb, df, sheet_Name, position_in_wb):
    
    from openpyxl.utils.dataframe import dataframe_to_rows

    ws = wb.create_sheet(sheet_Name, position_in_wb)

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    
    return ws, wb


def unmangleCols2(df):
    cols=pd.Series(df.columns)
    cols[cols.astype(str).str.isnumeric()] = cols[cols.astype(str).str.isnumeric()].astype(str)
    for dup in cols[cols.duplicated()].unique():
        cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) for i in range(1,sum(cols == dup)+1)]
    df.columns = cols
    return df



