# -*- coding: utf-8 -*-
"""
Created on Wed Sep 22 13:41:05 2021

@author: alejandro.gutierrez
"""


# -*- coding: utf-8 -*-
"""
Created on Thu Sep  2 22:22:25 2021

@author: alejandro.gutierrez
"""

#Importar paqueterias
import os
os.chdir('C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\Documents\\ALEJANDRO RAMOS GTZ\\GIT\\RITE INSIGHT INVENTORY') # relative path: scripts dir is under Lab

import unittest
import time 
import datetime
import pandas as pd
from datetime import date, timedelta, datetime

from Submain_Inventory import process_inventory


class Inventory_RITE_INSIGHT_DATA(unittest.TestCase):
    
    def setUp(self):
        
        self.PageProcess = process_inventory()
        
        #Si se corre desde otra pc distinta a la de Alejandro Mover las rutas tanto import como export pero la de export cambiar completamente a donde Arturo tenga los queries con las direcciones en su pc
        #paths to import each report 
        self.import_path_inventory = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'

        #Export path CAMBIAR PARA HACER PRUEBAS
        self.export_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\TEST ACCESS\\RITE AID'

        #Week number 
        t1 = datetime.now()
        self.week_number = str(t1.strftime("%U"))
        
        
    #@unittest.skip('Not need now')
    def test_VIRMAX(self):
        
        import_name = 'VIRMAX RITE INSIGHT week %s.xlsx'%self.week_number
        export_name = 'VIRMAX LLC.xlsx'
        
        full_import_path = self.import_path_inventory + '\\' + import_name
        full_export_path = self.export_path + '\\' + export_name
        
        self.PageProcess.inventory(full_import_path,full_export_path)
        
        print("%s is READY!!"%export_name) 
    
    
    #@unittest.skip('Not need now')
    def test_STERNO(self):
        
        import_name = 'STERNO RITE INSIGHT week %s.xlsx'%self.week_number
        export_name = 'STERNO_HOME.xlsx'
        
        full_import_path = self.import_path_inventory + '\\' + import_name
        full_export_path = self.export_path + '\\' + export_name
        
        self.PageProcess.inventory(full_import_path,full_export_path)
        
        print("%s is READY!!"%export_name) 
        

    #@unittest.skip('Not need now')
    def test_SOYLENT(self):
        
        import_name = 'SOYLENT RITE INSIGHT week %s.xlsx'%self.week_number
        export_name = 'SOYLENT NUTRITION INC.xlsx'
        
        full_import_path = self.import_path_inventory + '\\' + import_name
        full_export_path = self.export_path + '\\' + export_name
        
        self.PageProcess.inventory(full_import_path,full_export_path)
        
        print("%s is READY!!"%export_name) 
        
        
    #@unittest.skip('Not need now')
    def test_KRAVE(self):
        
        import_name = 'KRAVE RITE INSIGHT week %s.xlsx'%self.week_number
        export_name = 'KRAVE PURE FOODS INC.xlsx'
        
        full_import_path = self.import_path_inventory + '\\' + import_name
        full_export_path = self.export_path + '\\' + export_name
        
        self.PageProcess.inventory(full_import_path,full_export_path)
        
        print("%s is READY!!"%export_name) 
        
               
    #@unittest.skip('Not need now')
    def test_GOLDEN_EYE(self):
        
        import_name = 'GOLDEN_EYE RITE INSIGHT week %s.xlsx'%self.week_number
        export_name = 'GOLDEN EYE MEDIA USA INC.xlsx'
        
        full_import_path = self.import_path_inventory + '\\' + import_name
        full_export_path = self.export_path + '\\' + export_name
        
        self.PageProcess.inventory(full_import_path,full_export_path)
        
        print("%s is READY!!"%export_name) 
        

    #@unittest.skip('Not need now')
    def test_EVOLVE_BRANDS(self):
        
        import_name = 'EVOLVE RITE INSIGHT week %s.xlsx'%self.week_number
        export_name = 'EVOLVE BRANDS LLC.xlsx'
        
        full_import_path = self.import_path_inventory + '\\' + import_name
        full_export_path = self.export_path + '\\' + export_name
        
        self.PageProcess.inventory(full_import_path,full_export_path)
        
        print("%s is READY!!"%export_name) 
        
        
    #@unittest.skip('Not need now')
    def test_EOS(self):
        
        import_name = 'EOS RITE INSIGHT week %s.xlsx'%self.week_number
        export_name = 'EOS.xlsx'
        
        full_import_path = self.import_path_inventory + '\\' + import_name
        full_export_path = self.export_path + '\\' + export_name
        
        self.PageProcess.inventory(full_import_path,full_export_path)
        
        print("%s is READY!!"%export_name) 
     
        
    @unittest.skip('Not need now')
    def test_HARIBO_RITE(self):
        
        import_name = 'Haribo RITE AID Dashboard.xlsx'
        
        full_export_path = self.export_path + '\\' + import_name
        
        self.PageProcess.only_update(full_export_path)
        
        print("%s is READY!!"%import_name) 
        
        
    @unittest.skip('Not need now')
    def test_HARIBO_WEEKLY(self):
        
        import_name = 'Haribo Weekly Trends RITE AID.xlsm'
        
        full_export_path = self.export_path + '\\' + import_name
        
        self.PageProcess.only_update(full_export_path)
        
        print("%s is READY!!"%import_name) 
        

    @unittest.skip('Not need now')
    def test_CANDY(self):
                        
        # Finding the next saturday to rename the file
        #date_time_str = '20/09/21' #Estas 3 lineas es para hacerlo manual (meter la fecha tu)
        #date_time_obj = datetime.strptime(date_time_str, '%d/%m/%y')
        #today = date_time_obj
        
        today = date.today()

        idx = (today.weekday() + 1) % 7
        sat = today + timedelta(7+idx-7)  
        sat = sat.strftime("%d/%m/%Y")
        satf = sat[3:5] + '.' + sat[0:2] + '.' + sat[8:10]
        
        export_name = 'TOTAL BASIC CANDY - RITE AID WE %s.xlsx'%satf #9.18.21
        export_path_candy = 'F:\\Walgreens\\RITE AID\\Analytics\\Basic Candy Weekly Reports'
        #full_export_path = self.export_path + '\\' + import_name
        
        self.PageProcess.only_update_candy(export_path_candy, export_name)
        
        
        print("%s is READY!!"%export_name) 
        
        
        
        
        
if __name__ == '__main__':
    unittest.main()