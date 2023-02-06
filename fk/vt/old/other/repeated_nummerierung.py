
import pandas as pd
import openpyxl
import tkinter.filedialog
from tkinter import *

'''
    @author: Aritz
'''
def duplicated_nummerierung_values():
    
 try:
     
    FILE = tkinter.filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])

    print('Procesando . . .')
    required_columns = [2, 3, 4, 8]                                        # Usamos las columnas C (2), D (3), I (8)
    
    wb = openpyxl.load_workbook(FILE)                                   # WorkBook
    sheets = wb.sheetnames
    
    output_df = pd.DataFrame()                                          # DataFrame que generaremos al final
    
    for sheet_name in sheets:
        if sheet_name.lower().startswith('vt') or sheet_name.lower().startswith('pim') or sheet_name.lower().startswith('st'):
            
            print('Procesando hoja: {}'.format(sheet_name))
    
            df = pd.read_excel(FILE, sheet_name, usecols=required_columns)
            df = df.dropna()   
            df = df.reset_index(drop=True)                                 
            df = df.drop(df.index[0])
    
            df.columns = ['vt', 'nummerierung', 'mvi', 'neu']                          # <neu> como abrebiatura de <Infobaugruppe NEU / aktuell>
            df = df.reset_index(drop=True)
            df['nummerierung'] = df['nummerierung'].astype(str)
    
            duplicated_rows = df[df.duplicated(subset=['nummerierung'])]        # Procesamos el <df> para quedarnos solo con los duplicados...
            duplicated_nummeriers = duplicated_rows['nummerierung'].to_list()
            index_of_duplicated = [index for index, value in df['nummerierung'].items() if value in duplicated_nummeriers]
            
            dpdf = df.iloc[index_of_duplicated].sort_values('nummerierung').reset_index(drop=True)  # DataFrame con valores duplicados
    
            nummerierung = dpdf['nummerierung'].unique().tolist()               
    
            for num in nummerierung:
                df_aux = dpdf[dpdf['nummerierung']==num]                        # Esta DF auxiliar contiene el dpdf subdividido por cada valor nummerierung.
                df_aux2 = df_aux.drop_duplicates(subset='neu')                  # Solo nos interesan los nummerier repetidos pero con valor <neu> DISTINTO.
                if len(df_aux2) > 1:
                    output_df = pd.concat([output_df, df_aux])
    
    
    if len(output_df) > 0:
        print('Contenido generado en: output.xlsx')
        output_df.columns = ['VT', 'Nummerierung', 'MV Intern', 'Infobaugruppe NEU / aktuell']
        output_df.to_excel('output.xlsx',index=False)
    else:
        print('No hay valores duplicados distintos')

 except:
       print("Error personalizado")
       
def main():
    
    duplicated_nummerierung_values()
    
if __name__== "__main__":
    main()
