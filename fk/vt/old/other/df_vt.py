# -*- coding: utf-8 -*-
"""
Created on Tue Oct 13 09:25:19 2020
**********************************************************************************************************************************
*  The purpouse of this file (python module), is to be called
*  from another one having the clean_fk.get_dict_fk() dictionary
*  in a variable. For example,
*  imagin that we import this file (python module) into another,
*  we could do the following:
*
*  import clean_fk
*  import vt                                       # Could be another path (take it just as an example) 
*
*  my_dict_fk = clean_fk.get_dict_fk()             # Dictionary with key (VT name) and (VT DataFrame) as value
*
*  for vt_name in my_dict_fk:
*            variant_nummerier_dict = vt.get_variants_nummerierungs_dict(my_dict_fk[vt_name]) # vt DataFrame as a parameter
*            #  the same with the other methods...
*
*  Hope you enjoy it :p
************************************************************************************************************************************
"""

            ########## IMPORT PARENT MODULE ##########
import os,sys,inspect
current_dir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir) 
######################################################################

import clean_fk
import pandas as pd


###### Para testear
#dict_fk = clean_fk.get_dict_fk()
 


def extract_ca(key, df_vt, lista):
    '''
        Taking the vt name as key, the vt DataFrame and a list of vt column names,
        it returns a DataFrame with those columns.
    '''
    
    dict_extraido = {}
    
    row = df_vt[lista]
    
    dict_extraido[key] = row
     
    return dict_extraido


def get_variants_nummerierungs_dict(df_vt):
    ''' 
        Taking the df of a VT as a parameter,
        returns a dictionary with Variant as a key and
        the corresponding nummerierungs as value.
    '''

    df_nummerierungs = df_vt['Nummerierung'].to_frame().reset_index(drop=True)
    df_variants = df_vt.iloc[0:, 15:].reset_index(drop=True)
    
    # df_nummerierungs + df_variants DataFrame
    df_num_var = pd.concat([df_nummerierungs, df_variants], axis=1, ignore_index=False)
    
    # Variant list
    variants = list(df_variants.columns)

    dict_variant = {}                                                           # Key: variant, Value: Nummerierungs

    for variant in variants:                                                    # Filter for X nummerierung
        df_variant = df_num_var[df_num_var[variant]=='X']                       # in every variant
        list_nummerierungs = df_variant['Nummerierung'].unique().tolist()
        dict_variant[variant] = list_nummerierungs


    return dict_variant



# ########## Para testear
# for vt in dict_fk:
#     dv = get_variants_nummerierungs_dict(dict_fk[vt]) 
#     dict_ca = extract_ca(vt, dict_fk[vt], ['VT','MV intern','Nummerierung'])
    