import os
import subprocess
import sys
import time
from ast import Str
from asyncio.windows_events import NULL
from curses import echo
from datetime import datetime, timedelta
from os import remove
from xmlrpc.client import DateTime

import pandas as pd
import pyodbc
import win32com.client
from dotenv import load_dotenv
import numpy as np
from numpy import int64
from pyodbc import Connection
from sqlalchemy import create_engine

# Variables de entorno
config = load_dotenv(".env") 

date = datetime.today().strftime('%m/%d/%Y')

date_report = datetime.strptime(date,"%m/%d/%Y")
date_start = (date_report - timedelta(days=7)).strftime('%d.%m.%Y')
date_report = date_report.strftime('%d.%m.%Y')

# Datos ingreso SAP LOGON
user = os.getenv('USER')
passw = os.getenv('PASS')
file_zros="ReporteZROS_"+date+".XLSX"
file_delivery="ReporteDeli_"+date+".XLSX"
route = os.getenv('ROUTE')

df = pd.DataFrame(pd.read_excel(route))
df = df.fillna(value=NULL) # IMPORTANTE, reemplaza campos vacíos por NULL
df['Material'] = df['Material'].astype(int64)
df = df['Material']

df.replace(to_replace =0,value = ' ',inplace = True) 
df.to_clipboard(index=False,header=False)

try:  
    # Abre SAP
    path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    subprocess.Popen(path)
    time.sleep(2)
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine

    # Conexión que abriremos
    connection = application.OpenConnection("ECC - Production", True)
    session = connection.Children(0)  

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = passw
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 14
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZNEWSKU"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(3)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").showContextMenu
    session.findById("wnd[0]/usr").selectContextMenuItem("%011")
    time.sleep(3)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 0
    session.findById("wnd[1]").sendVKey(24)
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 0
    session.findById("wnd[1]").sendVKey(8)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/ctxtS_DATE-LOW").caretPosition = 10
    session.findById("wnd[0]/usr/ctxtS_DATE-LOW").showContextMenu
    session.findById("wnd[0]/usr").selectContextMenuItem("&012")
    session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellRow = 4
    session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "4"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]").sendVKey(0)
    time.sleep(3)
    session.findById("wnd[0]").sendVKey(8)
    time.sleep(3)
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select() #Seguimiento al teclado
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "D:\COROJA00\Archivos"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "export5.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
    session.findById("wnd[1]/tbar[0]/btn[0]").press()







    #-----------Cierra SAP Y Excel--------------------
    #os.system("TASKKILL /F /IM saplogon.exe")
    #os.system("TASKKILL /F /IM EXCEL.exe")<

except Exception as e:
    print('Hoy no fue ', e)
