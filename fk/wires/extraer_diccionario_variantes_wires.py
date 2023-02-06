# -*- coding: utf-8 -*-
"""
Created on Fri Jul 16 11:22:14 2021

@author: eduardo.moreno
"""

import pandas as pd
import numpy as np



def dict_variante_wires(df_variante_wires):
    
    
    #Primera fila a header
    df_variante_wires.columns = df_variante_wires.iloc[0,:]
    df_variante_wires.drop(df_variante_wires.index[0],inplace=True)
        
        #Creamos un diccionario para con clave la columna variante y valor lista de wires
        
    dict_variante_wires_final={}
        
        #prueba=df_variante_wires.columns[60:]
        
        #columna='V10'
        #recorremos la lista de columnas y filtramos aquellas que tengan len=2 y empiezen por V
    for columna in df_variante_wires.columns:
            #columna='V1'
        if len(str(columna).strip())<=3 and str.lower(str(columna).strip())[0]=='v' and str.lower(str(columna))!='von':
            df_variante_wires = unmangleCols2(df_variante_wires)

            lista_columnas= list(df_variante_wires.columns)       
            for index, value in enumerate(lista_columnas):
                if value == 'Ltg-Nr._1':
                    
                    lista_columnas[index] = 'Ltg-Nr.'
                    
            df_variante_wires.columns= lista_columnas   
            
            Columnas
            
            'Ltg-Nr.' in df_variante_wires.columns
                
            lista_wires=df_variante_wires['Ltg-Nr.'][df_variante_wires[columna]=='x'].tolist()
            dict_variante_wires_final[columna]=lista_wires

    
    return dict_variante_wires_final



def wirelist_get_family(df_wirelist,familia):
    
    saltar=False
    #familia='226'
    #filtramos la familia del df
    df_wirelist_familia=df_wirelist[df_wirelist['Modulfamilie Zeichnung']==familia]
    
    #dividimos en dos data_frames
    try:
        fila_separacion=np.where(df_wirelist_familia['Ltg-Nr.']=='Ltg-Nr.')[0][0]
    
    except:
        saltar=True
        fila_separacion=1
        
    
    df_1=df_wirelist_familia.iloc[:fila_separacion,:]
    df_variante_wires=df_wirelist_familia.iloc[fila_separacion:,:]
    
    
    return df_1,df_variante_wires,saltar



def df_variante_externalmodul(df_1):
    
    #Eliminamos la primera fila
    df_1.drop(df_1.index[0],inplace=True)
    
    df_variante_externalmodul=df_1[['Ltg-Nr.','Verbindung']]
    
    return df_variante_externalmodul



def unmangleCols2(df):
    cols=pd.Series(df.columns)
    cols[cols.astype(str).str.isnumeric()] = cols[cols.astype(str).str.isnumeric()].astype(str)
    for dup in cols[cols.duplicated()].unique(): 
        cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) for i in range(1,sum(cols == dup)+1)]
    df.columns = cols
    return df





ruta="C:/Users/eduardo.moreno/Desktop/Ordenador Nuevo Edu/Antenas/D989193-7-F_Konzept_VW276_Inra_LOL_KW22_25_TAB016622_K_ZD20201112_PVS_IC_007.xlsx"


df_wirelist=pd.read_excel(ruta,sheet_name='Wirelist',header=10,converters={'Modulfamilie Zeichnung':str})

dict_variante_wires_familia={}

familias=df_wirelist['Modulfamilie Zeichnung'].dropna().drop_duplicates()


saltar=False

df_concatenado=pd.DataFrame()

for familia in familias:
    
    
    #familia="000"
    
    df_variantes_externalmodul,df_variante_wires,saltar=wirelist_get_family(df_wirelist,familia)
    dict_variante_wires_familia=dict_variante_wires(df_variante_wires)
    df=pd.DataFrame.from_dict(dict_variante_wires_familia, orient='index').reset_index()
    df_bis=df_variante_externalmodul(df_variantes_externalmodul)
    
    df_merge=pd.merge(df_bis,df,how="inner",left_on=("Ltg-Nr."),right_on=("index"))
    
    df_concatenado=pd.concat([df_concatenado,df_merge])
    
    

excel=df_concatenado.to_excel("Modulos_wires.xlsx")

    
    
    
    






