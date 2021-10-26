# -*- coding: utf-8 -*-
"""
Created on Wed Sep 22 13:41:08 2021

@author: alejandro.gutierrez
"""

import win32com.client as win32
import glob
import os.path


class process_inventory():
    
    def __init__(self):
        self.POS_REPORT_ID = '//*[@id="ext-comp-1006__ext-comp-1002"]'
        
            
    def inventory(self,full_import_path,full_export_path):
        try:
            
            #Set app to work with Excel
            xlapp = win32.Dispatch('Excel.Application')
            xlapp.DisplayAlerts = False
            xlapp.Visible = False
            
            #Set paths
            path1 = full_import_path
            path2 = full_export_path
            
            #Open workbooks
            xlbook1 = xlapp.Workbooks.Open(path1)
            xlbook2 = xlapp.Workbooks.Open(path2)
            
            #Other process copying sheets
            sheet=xlbook1.Worksheets(1)
            #sheet.Move(Before=xlbook2.Worksheets("INVENTORY")) #This one is good too but is better use Copy instead Move 
            sheet.Copy(Before=xlbook2.Worksheets(1))
            
            #Borrar sheet llamada Inventory en el xlbook2 y cambiar el nombre a la sheet RateOfSale por Inventory
            xlbook2.Worksheets(2).Delete()
            xlbook2.Worksheets(1).Name = "INVENTORY"
            
            #Refresh connections and data
           # for conn in xlbook2.connections:
           #     conn.Refresh()
                
           # xlbook2.RefreshAll()
           # xlapp.CalculateUntilAsyncQueriesDone()# this will actually wait for the excel workbook to finish updating
            
            #Save and close
            xlbook1.Save()
            xlbook1.Close()
            
            xlbook2.Save()
            xlbook2.Close()
            
            xlapp.Quit()
            
            #Delete app and xlbooks
            del xlbook1
            del xlbook2            
            del xlapp 

                        
        except:
            print ("Something wrong please check what is happening with Submain Inventory process")
            

    def only_update(self, path):
        try:
            xlapp = win32.Dispatch('Excel.Application')
            xlapp.DisplayAlerts = False
            xlapp.Visible = True #False
            
            #Set paths
            path1 = path
            
            #Open workbooks
            xlbook1 = xlapp.Workbooks.Open(path1)
        
            #Update all
            for conn in xlbook1.connections:
                conn.Refresh()
                
            xlbook1.RefreshAll()
            xlapp.CalculateUntilAsyncQueriesDone()# this will actually wait for the excel workbook to finish updating
            
            #Save and close
            xlbook1.Save()
            xlbook1.Close()
            xlapp.Quit()
        
            del xlbook1
            del xlapp 
            
        except:
            print('Something wrong please check what is happening with Submain Update Only process')
            
            
            
    def only_update_candy(self, path, name):
        try:
            xlapp = win32.Dispatch('Excel.Application')
            xlapp.DisplayAlerts = False
            xlapp.Visible = False
                        
            #Encontrar el archivo mas reciente en el path y abrirlo
            folder_path = path
            file_type = '\*xlsx'
            files = glob.glob(folder_path + file_type)
            max_file = max(files, key=os.path.getctime)
            
            #Open workbooks
            xlbook1 = xlapp.Workbooks.Open(max_file)
        
            #Update all
            for conn in xlbook1.connections:
                conn.Refresh()
                
            xlbook1.RefreshAll()
            xlapp.CalculateUntilAsyncQueriesDone()# this will actually wait for the excel workbook to finish updating
            
            #Save and close
            #xlbook1.save(path + '\\' + name)
            xlbook1.SaveAs(Filename=path + '\\' + name)
            #xlbook1.Save
            xlbook1.Close()
            xlapp.Quit()
        
            del xlbook1
            del xlapp 
            
        except:
            print('Something wrong please check what is happening with Submain Update Only CANDY process')