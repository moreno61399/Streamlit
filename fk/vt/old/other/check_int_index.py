
# -*- coding: utf-8 -*-
"""
Created on Tue Sep 29 15:24:37 2020
@author: jesus.roldan
"""
#importamos los paquetes a usar
import tkinter.filedialog
import pandas as pd
import openpyxl

def check_int_index():
    try:
#pedimos al usuario el archivo a trabajar
        path = tkinter.filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
#creamos el wokbook con el archivo seleccionado        
        wb = openpyxl.load_workbook(path)
#guardamos todas las hojas del archivo excel        
        hojas = wb.sheetnames
#creo un DataFrame vacio para utilizarlo mas adelante        
        df_final = pd.DataFrame()
#recorremos las hojas inviduales que cumplas la condición especificada        
        for name in hojas:
             if name[:2].lower() == 'VT' or name[:2].lower() =='PIM' or name[:2].lower() == 'St':
#leemos los archivos excel con pandas                
                df = pd.read_excel(path,sheet_name = name)
#recorro las 5 pirmeras columnas                
                for col in df.columns[:5]:
#Pregunto con una condificon cuantas columnas llamadas 'MV intern' se encuentran en la hoja excel
                    if len(df[df[col].astype(str).str.contains('MV intern')])>=1:
#Hago una busqueda para que me devuelvan el index de la columna llamada 'MV intern'                        
                        row = df[df[col].astype(str).str.contains('MV intern')].index[0]
#Creo un DataFrame nuevo desde la posicion(index devuelta anteriormente) hasta el final                 
                df2= df.iloc[row:,:]
#creo nuevo DataFrame con los titulos                 
                headers = df2.iloc[0]
#el DataFrame df2 completa el df_of_interest modificando los encabezados del nuevo DataFrame
                df_of_interest  = pd.DataFrame(df2.values[1:], columns=headers)
#creo un nuevo DataFrame llamdo df_of_comparar con las columnas 'MV intern','int. Index'              
                df_of_comparar = pd.DataFrame(df_of_interest.loc[:,['MV intern','int. Index']])
#en la columna MV intern me quedo con su ultimo caracter                
                df_of_comparar['MV intern'] = df_of_comparar['MV intern'].str[-1:]
#borro todos los NAN, espacios en blanco por que sino no me los imprime en el resultado 
                df_of_comparar = df_of_comparar.loc[:,['MV intern','int. Index']].dropna(how='all')
#creo un nuevo DataFrame llamado solución el cual compara los elementos de las dos tablas (MV intern,int. Index) y se rellena de los valores que son diferentes                
                solucion = pd.DataFrame(df_of_comparar[df_of_comparar['MV intern'] !=  df_of_comparar['int. Index']])
#Añado una nueva columna llamada VT al DataFrame 'solución' en dicha columna se añaden los nombres de las hojas que hayan cumplido las condiciones puestas hasta ahora                  
                solucion['VT'] = name
#Concateno dos DataFrame, uno df_final que esta vacio y el resultado final que esta en solución. Este proceso se hace para que no se sobrescriban los resultados ya que estamos dentro de un for                
                df_final = pd.concat([df_final,solucion])
#Para terminar, se borran todos los NAN que se hayan podido colar y se reemplaza el DataFrame con NAN por el sin NAN               
                df_final.dropna(how='all',inplace=True)
#Exportamos el DataFrame a una hoja excel y eliminamos la columa de index              
                df_final.to_excel('Solucion.xlsx',index=False)

       
        
    except:
       print("Error personalizado")

def main():
    

      check_int_index()

if __name__ == "__main__":
    main()
