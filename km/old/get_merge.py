

# -*- coding: utf-8 -*-
"""
Created on Wed Oct 21 10:15:59 2020
@author: Aitor.vidart
"""

import os,sys,inspect
current_dir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir) 

import km.km as km
from orders import orders as order


     
def get_sbnModule(lista_KMs, status,file_path=None):


  '''Get DataFrame with 3 colums: ModuleNoOem,ModuleNoSEBN
     and CarID. Merge between kameliste status and Orders
     left or rigth / dataframe or dict'''
    
  #order.load_orders_file(file_path)
  df_txt = order.leer_pedido(file_path)
  km_1 = km.kameliste(lista_KMs, status)
  km_1.columns = ['ModuleNoOem','estructuras','Bauraum']
  

  
  df_output = df_txt.merge(km_1,on='ModuleNoOem',how='left')
  
  return df_output


