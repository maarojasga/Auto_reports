import os
import pyodbc
import numpy as np
import pandas as pd
from os import remove
from datetime import datetime
from dotenv import load_dotenv
from asyncio.windows_events import NULL
from email.message import EmailMessage
import smtplib

config = load_dotenv(".env") # Usar .env

# Función para extraer los archivos de las carpetas
def extract(folder):
    dir = os.getenv('DIR_ENT')+folder # Directorio
    folders_excel = []
    folders = os.listdir(dir)
    for folder in folders:
        if folder[-5:] == '.xlsx':
            folders_excel.append(folder)
    return folders_excel

# Credenciales correo
user = os.getenv('LOGIN_USER')
password = os.getenv('LOGIN_PASS')
remitente = user
destinatario = os.getenv('LOGIN_USER')

# Función para enviar correos en caso de error
def sendEmailError(report):
     mensaje = "Algo salió mal, no se ejecutó "+report+". El error es: {}".format(e)
     email = EmailMessage()
     email["Subject"] = "ERROR "+report
     email.set_content(mensaje)
     smtp = smtplib.SMTP("smtp-mail.outlook.com", port=587)
     smtp.starttls()
     smtp.login(remitente, password)
     smtp.sendmail(remitente, destinatario, email.as_string())
     smtp.quit()

# Función para enviar correos en caso de éxito
def sendEmailOK(report):
     mensaje = "Todo salió de maravilla con: "+report
     email = EmailMessage()
     email["Subject"] = "EJECUTADO "+report
     email.set_content(mensaje)
     smtp = smtplib.SMTP("smtp-mail.outlook.com", port=587)
     smtp.starttls()
     smtp.login(remitente, password)
     smtp.sendmail(remitente, destinatario, email.as_string())
     smtp.quit()

# Extraer los archivos de las carpetas PESOS
folders_daily_cop = extract(os.getenv('DAILY_COP')) 
folders_inventory = extract(os.getenv('INVENTORY'))
folders_vendor = extract(os.getenv('VENDOR'))
print(folders_daily_cop)
# Extraer los archivos de las carpetas USD
folders_daily_usd = extract(os.getenv('DAILY_USD')) 

# Conectar con la base de datos

#cnxn = pyodbc.connect(key)
cnxn = pyodbc.connect(os.getenv('CONN'))
crsr = cnxn.cursor()
date = datetime.now()
date_reporte = datetime.today().strftime('%Y-%m-%d')

print('Aquí vamos')
# Insertar data Daily Sales en la base de datos
try:
     for file in folders_daily_cop:
     
          df = pd.DataFrame(pd.read_excel(os.getenv('DIR_ENT')+os.getenv('DAILY_COP')+file)) # Insertar los datos en un dataframe
          df = df.fillna(value=NULL) # IMPORTANTE, reemplaza campos vacíos por NULL
          df = df[:-1]
          df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], format='%d.%m.%Y')
          df['L01 Sold-to party Tax Code 1 (Key)'] = df['L01 Sold-to party Tax Code 1 (Key)'].replace({'-':''}, regex=True)
          df['L01 Sold-to party Tax Code 1 (Key)'] = df['L01 Sold-to party Tax Code 1 (Key)'].replace({'#':''}, regex=True)
          df['Vendor Number'] = df['Vendor Number'].replace({'#':'0'}, regex=True)
          df['Vendor Subrange Number'] = df['Vendor Subrange Number'].replace({'#':'0'}, regex=True)
          df['Order Date'] = df['Order Date'].replace('.','/') # Sólo aplica para Daily - convertir a fecha
          df['Order Date'] = pd.to_datetime(df['Order Date'], format='%d.%m.%Y')
          df2 = df # Copia del dataframe original para histórico
          df2["Date"] = datetime.now()
          df2["Type"] = "Daily_COP_SEMANAL"
          crsr.execute("DELETE FROM [dbo].[Daily_Sales_COP]") # Eliminar data de la tabla
          crsr.execute("DELETE FROM [dbo].[Daily_Sales_His_COP]")

          # Insertar Dataframe en SQL Server:
          for index,row in df.iterrows(): # Iterar df para insertar filas en la tabla

               crsr.execute("INSERT INTO [dbo].[Daily_Sales_COP] ([Sales_Organization_Number] ,[Sales_Organization_Name] ,[Distribution_Channel_Number] ,[Distribution_Channel_Name] ,[Division_Number] ,[Division_Name] ,[Customer_Number] ,[Customer_Name] ,[Customer_Segment] ,[Ship_To_Party_Key] ,[Ship_To_Party] ,[Shipping_Conditions] ,[Payment_Terms] ,[Web_Sales] ,[Vendor_Number] ,[Vendor_Name] ,[Vendor_Subrange_Number] ,[Vendor_Subrange_Name] ,[Sales_Manager_Number] ,[Sales_Manager_Name] ,[IS_Rep_Number] ,[IS_Rep_Name] ,[OS_Rep_Number] ,[OS_Rep_Name] ,[Invoice_Date] ,[Invoice_Number] ,[Invoice_Line_Item] ,[Order_Date] ,[Order_Type] ,[Order_Type_Description] ,[Customer_PO_Number] ,[Order_Number] ,[Line_Number] ,[Material] ,[Material_Description] ,[Line_Type] ,[Manufacturer_Part_Number] ,[Country_Code] ,[Country_Name] ,[Quantity] ,[Unit] ,[Sales_Revenue] ,[Sales_Cost] ,[Sales_Margin] ,[Calculated_Price] ,[Price_Override_Flag] ,[Currency] ,[L01_Sold_to_party_Tax_Code_1_Key] ,[L01_Valuation_type_Key] ,[L01_Material_Material_type_Key] ,[Freight_Recovery] ,[ACOPs] ,[L01_Vendor_BID_ID] ,[L01_End_User] ,[L01_IM_End_User] ,[Cost] ,[Negotiated_Price] ,[ZMRK] ,[ZVFM] ,[Internal_Price_VPRS] ,[Sales_Material_ABC]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row['Sales Organization Number'] ,row['Sales Organization Name'] ,row['Distribution Channel Number'] ,row['Distribution Channel Name'] ,row['Division Number'] ,row['Division Name'] ,row['Customer Number'] ,row['Customer Name'] ,row['Customer Segment'] ,row['Ship-To Party Key'] ,row['Ship-To Party'] ,row['Shipping Conditions'] ,row['Payment Terms'] ,row['Web Sales'] ,row['Vendor Number'] ,row['Vendor Name'] ,row['Vendor Subrange Number'] ,row['Vendor Subrange Name'] ,row['Sales Manager Number'] ,row['Sales Manager Name'] ,row['IS Rep Number'] ,row['IS Rep Name'] ,row['OS Rep Number'] ,row['OS Rep Name'] ,row['Invoice Date'] ,row['Invoice Number'] ,row['Invoice Line Item'] ,row['Order Date'] ,row['Order Type'] ,row['Order Type Description'] ,row['Customer PO Number'] ,row['Order Number'] ,row['Line Number'] ,row['Material'] ,row['Material Description'] ,row['Line Type'] ,row['Manufacturer Part Number'] ,row['Country Code'] ,row['Country Name'] ,row['Quantity'] ,row['Unit'] ,row['Sales Revenue'] ,row['Sales Cost'] ,row['Sales Margin'] ,row['Calculated Price'] ,row['Price Override Flag'] ,row['Currency'] ,row['L01 Sold-to party Tax Code 1 (Key)'] ,row['L01 Valuation type Key'] ,row['L01 Material Material type (Key)'] ,row['Freight_Recovery'] ,row['ACOPs'] ,row['L01 Vendor BID ID'] ,row['L01 End User'] ,row['L01 IM End User'] ,row['Cost'] ,row['Negotiated Price'] ,row['ZMRK'] ,row['ZVFM'] ,row['Internal Price VPRS'] ,row['Sales_Material_ABC'])
               print(index,' Ahí vamos0')
               
          for index,row in df2.iterrows(): # Iterar df2 para insertar filas en la tabla
               crsr.execute("INSERT INTO [dbo].[Daily_Sales_His_COP] ([Date], [Sales_Organization_Number] ,[Sales_Organization_Name] ,[Distribution_Channel_Number] ,[Distribution_Channel_Name] ,[Division_Number] ,[Division_Name] ,[Customer_Number] ,[Customer_Name] ,[Customer_Segment] ,[Ship_To_Party_Key] ,[Ship_To_Party] ,[Shipping_Conditions] ,[Payment_Terms] ,[Web_Sales] ,[Vendor_Number] ,[Vendor_Name] ,[Vendor_Subrange_Number] ,[Vendor_Subrange_Name] ,[Sales_Manager_Number] ,[Sales_Manager_Name] ,[IS_Rep_Number] ,[IS_Rep_Name] ,[OS_Rep_Number] ,[OS_Rep_Name] ,[Invoice_Date] ,[Invoice_Number] ,[Invoice_Line_Item] ,[Order_Date] ,[Order_Type] ,[Order_Type_Description] ,[Customer_PO_Number] ,[Order_Number] ,[Line_Number] ,[Material] ,[Material_Description] ,[Line_Type] ,[Manufacturer_Part_Number] ,[Country_Code] ,[Country_Name] ,[Quantity] ,[Unit] ,[Sales_Revenue] ,[Sales_Cost] ,[Sales_Margin] ,[Calculated_Price] ,[Price_Override_Flag] ,[Currency] ,[L01_Sold_to_party_Tax_Code_1_Key] ,[L01_Valuation_type_Key] ,[L01_Material_Material_type_Key] ,[Freight_Recovery] ,[ACOPs] ,[L01_Vendor_BID_ID] ,[L01_End_User] ,[L01_IM_End_User] ,[Cost] ,[Negotiated_Price] ,[ZMRK] ,[ZVFM] ,[Internal_Price_VPRS] ,[Sales_Material_ABC],[Type]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row['Date'], row['Sales Organization Number'], row['Sales Organization Name'] ,row['Distribution Channel Number'] ,row['Distribution Channel Name'] ,row['Division Number'] ,row['Division Name'] ,row['Customer Number'] ,row['Customer Name'] ,row['Customer Segment'] ,row['Ship-To Party Key'] ,row['Ship-To Party'] ,row['Shipping Conditions'] ,row['Payment Terms'] ,row['Web Sales'] ,row['Vendor Number'] ,row['Vendor Name'] ,row['Vendor Subrange Number'] ,row['Vendor Subrange Name'] ,row['Sales Manager Number'] ,row['Sales Manager Name'] ,row['IS Rep Number'] ,row['IS Rep Name'] ,row['OS Rep Number'] ,row['OS Rep Name'] ,row['Invoice Date'] ,row['Invoice Number'] ,row['Invoice Line Item'] ,row['Order Date'] ,row['Order Type'] ,row['Order Type Description'] ,row['Customer PO Number'] ,row['Order Number'] ,row['Line Number'] ,row['Material'] ,row['Material Description'] ,row['Line Type'] ,row['Manufacturer Part Number'] ,row['Country Code'] ,row['Country Name'] ,row['Quantity'] ,row['Unit'] ,row['Sales Revenue'] ,row['Sales Cost'] ,row['Sales Margin'] ,row['Calculated Price'] ,row['Price Override Flag'] ,row['Currency'] ,row['L01 Sold-to party Tax Code 1 (Key)'] ,row['L01 Valuation type Key'] ,row['L01 Material Material type (Key)'] ,row['Freight_Recovery'] ,row['ACOPs'] ,row['L01 Vendor BID ID'] ,row['L01 End User'] ,row['L01 IM End User'] ,row['Cost'] ,row['Negotiated Price'] ,row['ZMRK'] ,row['ZVFM'] ,row['Internal Price VPRS'] ,row['Sales_Material_ABC'], row['Type'])
               print(index,' Ahí vamos1')
          print("Daily Sales ejecutado") # Mensaje de control
          cnxn.commit()
          sendEmailOK("DAILY SALES COP")
          #remove(os.getenv('DIR_ENT')+os.getenv('DAILY_COP')+file)
     
except Exception as e:
     sendEmailError("DAILY SALES COP")
     print("Algo salió mal, no se ejecutó . El error es: {}".format(e))

try:
     for file in folders_daily_usd:
          df = pd.DataFrame(pd.read_excel(os.getenv('DIR_ENT')+os.getenv('DAILY_USD')+file)) # Insertar los datos en un dataframe
          df = df.fillna(value=NULL) # IMPORTANTE, reemplaza campos vacíos por NULL
          df = df[:-1]
          df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], format='%d.%m.%Y')
          df['L01 Sold-to party Tax Code 1 (Key)'] = df['L01 Sold-to party Tax Code 1 (Key)'].replace({'-':''}, regex=True)
          df['L01 Sold-to party Tax Code 1 (Key)'] = df['L01 Sold-to party Tax Code 1 (Key)'].replace({'#':''}, regex=True)
          df['Vendor Number'] = df['Vendor Number'].replace({'#':'0'}, regex=True)
          df['Vendor Subrange Number'] = df['Vendor Subrange Number'].replace({'#':'0'}, regex=True)
          df['Order Date'] = pd.to_datetime(df['Order Date'].replace('.','/')) # Sólo aplica para Daily - convertir a fecha
          df['Order Date'] = pd.to_datetime(df['Order Date'], format='%d.%m.%Y')
          df2 = df # Copia del dataframe original para histórico
          df2["Date"] = datetime.now()
          df2["Type"] = "Daily_USD_SEMANAL"
          crsr.execute("DELETE FROM [dbo].[Daily_Sales_USD]") # Eliminar data de la tabla
          crsr.execute("DELETE FROM [dbo].[Daily_Sales_His_USD]")
     
          # Insertar Dataframe en SQL Server:
          for index,row in df.iterrows(): # Iterar df para insertar filas en la tabla
               crsr.execute("INSERT INTO [dbo].[Daily_Sales_USD] ([Sales_Organization_Number] ,[Sales_Organization_Name] ,[Distribution_Channel_Number] ,[Distribution_Channel_Name] ,[Division_Number] ,[Division_Name] ,[Customer_Number] ,[Customer_Name] ,[Customer_Segment] ,[Ship_To_Party_Key] ,[Ship_To_Party] ,[Shipping_Conditions] ,[Payment_Terms] ,[Web_Sales] ,[Vendor_Number] ,[Vendor_Name] ,[Vendor_Subrange_Number] ,[Vendor_Subrange_Name] ,[Sales_Manager_Number] ,[Sales_Manager_Name] ,[IS_Rep_Number] ,[IS_Rep_Name] ,[OS_Rep_Number] ,[OS_Rep_Name] ,[Invoice_Date] ,[Invoice_Number] ,[Invoice_Line_Item] ,[Order_Date] ,[Order_Type] ,[Order_Type_Description] ,[Customer_PO_Number] ,[Order_Number] ,[Line_Number] ,[Material] ,[Material_Description] ,[Line_Type] ,[Manufacturer_Part_Number] ,[Country_Code] ,[Country_Name] ,[Quantity] ,[Unit] ,[Sales_Revenue] ,[Sales_Cost] ,[Sales_Margin] ,[Calculated_Price] ,[Price_Override_Flag] ,[Currency] ,[L01_Sold_to_party_Tax_Code_1_Key] ,[L01_Valuation_type_Key] ,[L01_Material_Material_type_Key] ,[Freight_Recovery] ,[ACOPs] ,[L01_Vendor_BID_ID] ,[L01_End_User] ,[L01_IM_End_User] ,[Cost] ,[Negotiated_Price] ,[ZMRK] ,[ZVFM] ,[Internal_Price_VPRS] ,[Sales_Material_ABC]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row['Sales Organization Number'] ,row['Sales Organization Name'] ,row['Distribution Channel Number'] ,row['Distribution Channel Name'] ,row['Division Number'] ,row['Division Name'] ,row['Customer Number'] ,row['Customer Name'] ,row['Customer Segment'] ,row['Ship-To Party Key'] ,row['Ship-To Party'] ,row['Shipping Conditions'] ,row['Payment Terms'] ,row['Web Sales'] ,row['Vendor Number'] ,row['Vendor Name'] ,row['Vendor Subrange Number'] ,row['Vendor Subrange Name'] ,row['Sales Manager Number'] ,row['Sales Manager Name'] ,row['IS Rep Number'] ,row['IS Rep Name'] ,row['OS Rep Number'] ,row['OS Rep Name'] ,row['Invoice Date'] ,row['Invoice Number'] ,row['Invoice Line Item'] ,row['Order Date'] ,row['Order Type'] ,row['Order Type Description'] ,row['Customer PO Number'] ,row['Order Number'] ,row['Line Number'] ,row['Material'] ,row['Material Description'] ,row['Line Type'] ,row['Manufacturer Part Number'] ,row['Country Code'] ,row['Country Name'] ,row['Quantity'] ,row['Unit'] ,row['Sales Revenue'] ,row['Sales Cost'] ,row['Sales Margin'] ,row['Calculated Price'] ,row['Price Override Flag'] ,row['Currency'] ,row['L01 Sold-to party Tax Code 1 (Key)'] ,row['L01 Valuation type Key'] ,row['L01 Material Material type (Key)'] ,row['Freight_Recovery'] ,row['ACOPs'] ,row['L01 Vendor BID ID'] ,row['L01 End User'] ,row['L01 IM End User'] ,row['Cost'] ,row['Negotiated Price'] ,row['ZMRK'] ,row['ZVFM'] ,row['Internal Price VPRS'] ,row['Sales_Material_ABC'])
               print(index,' Ahí vamos0')
          for index,row in df2.iterrows(): # Iterar df2 para insertar filas en la tabla
               crsr.execute("INSERT INTO [dbo].[Daily_Sales_His_USD] ([Date], [Sales_Organization_Number] ,[Sales_Organization_Name] ,[Distribution_Channel_Number] ,[Distribution_Channel_Name] ,[Division_Number] ,[Division_Name] ,[Customer_Number] ,[Customer_Name] ,[Customer_Segment] ,[Ship_To_Party_Key] ,[Ship_To_Party] ,[Shipping_Conditions] ,[Payment_Terms] ,[Web_Sales] ,[Vendor_Number] ,[Vendor_Name] ,[Vendor_Subrange_Number] ,[Vendor_Subrange_Name] ,[Sales_Manager_Number] ,[Sales_Manager_Name] ,[IS_Rep_Number] ,[IS_Rep_Name] ,[OS_Rep_Number] ,[OS_Rep_Name] ,[Invoice_Date] ,[Invoice_Number] ,[Invoice_Line_Item] ,[Order_Date] ,[Order_Type] ,[Order_Type_Description] ,[Customer_PO_Number] ,[Order_Number] ,[Line_Number] ,[Material] ,[Material_Description] ,[Line_Type] ,[Manufacturer_Part_Number] ,[Country_Code] ,[Country_Name] ,[Quantity] ,[Unit] ,[Sales_Revenue] ,[Sales_Cost] ,[Sales_Margin] ,[Calculated_Price] ,[Price_Override_Flag] ,[Currency] ,[L01_Sold_to_party_Tax_Code_1_Key] ,[L01_Valuation_type_Key] ,[L01_Material_Material_type_Key] ,[Freight_Recovery] ,[ACOPs] ,[L01_Vendor_BID_ID] ,[L01_End_User] ,[L01_IM_End_User] ,[Cost] ,[Negotiated_Price] ,[ZMRK] ,[ZVFM] ,[Internal_Price_VPRS] ,[Sales_Material_ABC],[Type]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row['Date'], row['Sales Organization Number'], row['Sales Organization Name'] ,row['Distribution Channel Number'] ,row['Distribution Channel Name'] ,row['Division Number'] ,row['Division Name'] ,row['Customer Number'] ,row['Customer Name'] ,row['Customer Segment'] ,row['Ship-To Party Key'] ,row['Ship-To Party'] ,row['Shipping Conditions'] ,row['Payment Terms'] ,row['Web Sales'] ,row['Vendor Number'] ,row['Vendor Name'] ,row['Vendor Subrange Number'] ,row['Vendor Subrange Name'] ,row['Sales Manager Number'] ,row['Sales Manager Name'] ,row['IS Rep Number'] ,row['IS Rep Name'] ,row['OS Rep Number'] ,row['OS Rep Name'] ,row['Invoice Date'] ,row['Invoice Number'] ,row['Invoice Line Item'] ,row['Order Date'] ,row['Order Type'] ,row['Order Type Description'] ,row['Customer PO Number'] ,row['Order Number'] ,row['Line Number'] ,row['Material'] ,row['Material Description'] ,row['Line Type'] ,row['Manufacturer Part Number'] ,row['Country Code'] ,row['Country Name'] ,row['Quantity'] ,row['Unit'] ,row['Sales Revenue'] ,row['Sales Cost'] ,row['Sales Margin'] ,row['Calculated Price'] ,row['Price Override Flag'] ,row['Currency'] ,row['L01 Sold-to party Tax Code 1 (Key)'] ,row['L01 Valuation type Key'] ,row['L01 Material Material type (Key)'] ,row['Freight_Recovery'] ,row['ACOPs'] ,row['L01 Vendor BID ID'] ,row['L01 End User'] ,row['L01 IM End User'] ,row['Cost'] ,row['Negotiated Price'] ,row['ZMRK'] ,row['ZVFM'] ,row['Internal Price VPRS'] ,row['Sales_Material_ABC'],row['Type'])
               print(index,' Ahí vamos1')
          print("Daily Sales ejecutado") # Mensaje de control
          cnxn.commit()
          sendEmailOK("DAILY SALES USD")
          #remove(os.getenv('DIR_ENT')+os.getenv('DAILY_USD')+file)
except Exception as e:
     sendEmailError("DAILY SALES USD")
     print("Algo salió mal, no se ejecutó . El error es: {}".format(e))

# Insertar data Inventory en la base de datos
try:
     for file in folders_inventory:
          df = pd.DataFrame(pd.read_excel(os.getenv('DIR_ENT')+os.getenv('INVENTORY')+file)) # Insertar los datos en un dataframe
          df = df.fillna(value=NULL)
          df = df.rename(columns=df.iloc[0]).drop(df.index[0]) # Sólo aplica para INVENTORY
          df = df[:-1]
          df2 = df # Copia del dataframe original para histórico
          df2["Date"] = datetime.now()
          df2["Type"] = "Inventory_SEMANAL"
          crsr.execute("DELETE FROM [dbo].[Inventory]") # Eliminar data de la tabla
          crsr.execute("DELETE FROM [dbo].[Inventory_HIS]")
     
          # Insertar Dataframe el SQL Server:
          for index,row in df.iterrows(): # Iterar df para insertar filas en la tabla
               crsr.execute("INSERT INTO [dbo].[Inventory]([Key_date],[Company_code],[DIV_Name],[DEPT_Name],[Vendor],[Vendor_Description],[VSR],[VSR_Name],[Material],[Storage_Location_Key],[Material_Description],[VPN],[Material_CRT],[Material_CRV],[Valuation],[Plant],[Plant_Description],[Total_BOH],[Total_BOH_Value_$],[Cost_per_Unit_(MAP)]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row['Key date'],row['Company code'],row['DIV Name'],row['DEPT Name'],row['Vendor'],row['Vendor Description'],row['VSR'],row['VSR Name'],row['Material'],row['Storage Location Key'],row['Material Description'],row['VPN'],row['Material CRT'],row['Material CRV'],row['Valuation'],row['Plant'],row['Plant Description'],row['Total BOH'],row['Total BOH Value $'],row['Cost per Unit (MAP)'])
     
               print(index,' Ahí vamos2')
          for index,row in df2.iterrows(): # Iterar df2 para insertar filas en la tabla
               crsr.execute("INSERT INTO [dbo].[Inventory_His]([Date],[Key_date],[Company_code],[DIV_Name],[DEPT_Name],[Vendor],[Vendor_Description],[VSR],[VSR_Name],[Material],[Storage_Location_Key],[Material_Description],[VPN],[Material_CRT],[Material_CRV],[Valuation],[Plant],[Plant_Description],[Total_BOH],[Total_BOH_Value_$],[Cost_per_Unit_(MAP)],[Type]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row['Date'],row['Key date'],row['Company code'],row['DIV Name'],row['DEPT Name'],row['Vendor'],row['Vendor Description'],row['VSR'],row['VSR Name'],row['Material'],row['Storage Location Key'],row['Material Description'],row['VPN'],row['Material CRT'],row['Material CRV'],row['Valuation'],row['Plant'],row['Plant Description'],row['Total BOH'],row['Total BOH Value $'],row['Cost per Unit (MAP)'],row['Type'])
               print(index,' Ahí vamos2')
          print("Inventory ejecutado") # Mensaje de control
          cnxn .commit()
          sendEmailOK("INVENTORY")

          #remove(os.getenv('DIR_INV')+file)
except Exception as e:
     sendEmailError("INVENTORY")
     print("Algo salió mal, no se ejecutó . El error es: {}".format(e))

# Insertar data Vendor en la base de datos
try:
     for file in folders_vendor:
          df = pd.DataFrame(pd.read_excel(os.getenv('DIR_ENT')+os.getenv('VENDOR')+file)) # Insertar los datos en un dataframe
          df['Invoice Date\n'] = pd.to_datetime(df['Invoice Date\n'], format="%d.%m.%Y") # Sólo aplica para Vendor
          df = df.drop(df.index[[0]]) # Sólo aplica para Vendors
          df = df.fillna(value=NULL) # IMPORTANTE, reemplaza campos vacíos por NULL
          df2 = df # Copia del dataframe original para histórico
          df2["Date"] = datetime.now()
          df2["Type"] = "Vendor_SEMANAL"
          crsr.execute("DELETE FROM [dbo].[Vendor]") # Eliminar data de la tabla
          crsr.execute("DELETE FROM [dbo].[Vendor_His]")
          # Insertar Dataframe el SQL Server:
          for index,row in df.iterrows(): # Iterar df para insertar filas en la tabla
               crsr.execute("INSERT INTO [dbo].[Vendor] ([Distributor_Name],[Vendor_No],[Vendor_Name],[L01_Material_Plant_View_Vendor_Subrange_(Name)],[L01_Material_Plant_View_Vendor_Subrange_(Key)],[Distributor_Nbr],[Valuation_type],[Customer_Name],[Customer_Nbr],[Customer_Address_1],[Customer_Address_2],[Customer_PostalCode_ZIP],[Customer_City],[Customer_state/province/region],[Customer_Country],[Customer_Telephone_Nbr],[Customer_Contact_name],[Customer_VAT_Nbr],[End_Customer_Name],[End_Customer_Address_1],[End_Customer_Address_2],[End_Customer_City],[End_Customer_State/Provinve/Region],[End_Customer_Zip/Postal_Code],[End_Customer_Country],[IM_SKU],[mfr_part_nbr],[Product_descr1],[Product_Serial_Number],[Invoice_Date],[Invoice_nbr],[Line_nbr],[Billing_Quantity],[Net_Price],[Sales_Amount]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row[' Distributor Name'],row['Vendor No'],row['Vendor Name '],row['L01 Material Plant View Vendor Subrange (Name)'],row['L01 Material Plant View Vendor Subrange (Key)'],row['Distributor Nbr'],row['Valuation type'],row['Customer Name'],row['Customer Nbr'],row['Customer Address 1'],row['Customer Address 2'],row['Customer PostalCode ZIP'],row['Customer City'],row['Customer state / province / region'],row['Customer Country'],row['Customer Telephone Nbr'],row['Customer Contact name'],row['Customer VAT Nbr'],row['End Customer Name'],row['End Customer Address 1'],row['End Customer Address 2'],row['End Customer City'],row['End Customer State / Provinve / Region\n'],row['End Customer Zip/Postal Code\n'],row['End Customer Country\n'],row['IM_SKU\n'],row['mfr_part_nbr\n'],row['product descr1\n'],row['Product Serial Number'],row['Invoice Date\n'],row['Invoice nbr\n'],row['Line nbr'],row['Billing Quantity'],row['Net Price'],row['Sales Amount'])
               print(index,' Ahí vamos3')
          for index,row in df2.iterrows(): # Iterar df2 para insertar filas en la tabla
               crsr.execute("INSERT INTO [dbo].[Vendor_His] ([Date],[Distributor_Name],[Vendor_No],[Vendor_Name],[L01_Material_Plant_View_Vendor_Subrange_(Name)],[L01_Material_Plant_View_Vendor_Subrange_(Key)],[Distributor_Nbr],[Valuation_type],[Customer_Name],[Customer_Nbr],[Customer_Address_1],[Customer_Address_2],[Customer_PostalCode_ZIP],[Customer_City],[Customer_state/province/region],[Customer_Country],[Customer_Telephone_Nbr],[Customer_Contact_name],[Customer_VAT_Nbr],[End_Customer_Name],[End_Customer_Address_1],[End_Customer_Address_2],[End_Customer_City],[End_Customer_State/Provinve/Region],[End_Customer_Zip/Postal_Code],[End_Customer_Country],[IM_SKU],[mfr_part_nbr],[Product_descr1],[Product_Serial_Number],[Invoice_Date],[Invoice_nbr],[Line_nbr],[Billing_Quantity],[Net_Price],[Sales_Amount],[Type]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row['Date'],row[' Distributor Name'],row['Vendor No'],row['Vendor Name '],row['L01 Material Plant View Vendor Subrange (Name)'],row['L01 Material Plant View Vendor Subrange (Key)'],row['Distributor Nbr'],row['Valuation type'],row['Customer Name'],row['Customer Nbr'],row['Customer Address 1'],row['Customer Address 2'],row['Customer PostalCode ZIP'],row['Customer City'],row['Customer state / province / region'],row['Customer Country'],row['Customer Telephone Nbr'],row['Customer Contact name'],row['Customer VAT Nbr'],row['End Customer Name'],row['End Customer Address 1'],row['End Customer Address 2'],row['End Customer City'],row['End Customer State / Provinve / Region\n'],row['End Customer Zip/Postal Code\n'],row['End Customer Country\n'],row['IM_SKU\n'],row['mfr_part_nbr\n'],row['product descr1\n'],row['Product Serial Number'],row['Invoice Date\n'],row['Invoice nbr\n'],row['Line nbr'],row['Billing Quantity'],row['Net Price'],row['Sales Amount'], row["Type"])
               print(index,' Ahí vamos4')
          print("Vendor ejecutado") # Mensaje de control
          cnxn.commit() 
          
          sendEmailOK("VENDOR")
          #remove(os.getenv('DIR_ENT')+os.getenv('VENDOR')+file)
except Exception as e:
     sendEmailError("VENDOR")
     print("Algo salió mal, no se ejecutó . El error es: {}".format(e))
crsr.close() # Cerrar conexión



