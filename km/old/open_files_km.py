#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Oct 21 12:54:41 2020
"""

import os


RESOURCES_PATH = os.getcwd() + '/km/resources_km'

files = os.listdir(RESOURCES_PATH)


def get_gui_filenames():
    
    '''
        Search the directory resources for LOL and LOR
        files.
        
        Returns a list with [lol_files, lor_files]
        
    '''
    
    lol_files = []
    lor_files = []
    
    OPTIONS_LL = ['LL','L0L','LOL']
    OPTIONS_RL = ['RL','L0R','LOR','LR','ROL']
    
    for file_name in files:
        if any(x in file_name for x in OPTIONS_LL):
            lol_files.append(RESOURCES_PATH + '/' + file_name)

        elif any(x in file_name for x in OPTIONS_RL):
            lor_files.append(RESOURCES_PATH + '/' + file_name)
            
    return lol_files, lor_files
