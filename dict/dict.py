# -*- coding: utf-8 -*-
"""
Created on Fri Jan 29 14:45:03 2021

@author: jesus.roldan
"""
#permite actualizar un diccionario con otro donde hay elementos que ya existen y otros nuevos
def dict_update(dict_1,dict_2):
    
    for key,value in dict_2.items():
        if key in dict_1:
            dict_1[key].extend(value)
        else:
            dict_1[key]=value
    
    return dict_1