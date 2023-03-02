import os
from os import mkdir
import pyodbc
import smtplib
import calendar
import pandas as pd
from smtplib import SMTP
from email import encoders
from datetime import datetime
from dotenv import load_dotenv
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from asyncio.windows_events import NULL
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.message import EmailMessage
import smtplib
from datetime import datetime, timedelta
config = load_dotenv(".env") # Usar .env

# Credenciales correo
user = os.getenv('LOGIN_USER')
passw = os.getenv('LOGIN_PASS')
remitente = user


# Función para enviar correos en caso de error
def sendEmailError(report):
     mensaje = "Algo salió mal, no se envió el correo para "+report+". El error es: {}".format(e)
     email = EmailMessage()
     email["Subject"] = "ERROR DESCARGA "+report
     email.set_content(mensaje)
     smtp = smtplib.SMTP("smtp-mail.outlook.com", port=587)
     smtp.starttls()
     destinatario = os.getenv('LOGIN_USER')
     smtp.login(remitente, passw)
     smtp.sendmail(remitente, destinatario, email.as_string())
     smtp.quit()

# Función para enviar correos en caso de éxito
def sendEmailOK(report):
     mensaje = "Todo salió de maravilla con el correo de: "+report
     email = EmailMessage()
     email["Subject"] = "DESCARGADO "+report
     email.set_content(mensaje)
     smtp = smtplib.SMTP("smtp-mail.outlook.com", port=587)
     smtp.starttls()
     destinatario = os.getenv('LOGIN_USER')
     smtp.login(remitente, password)
     smtp.sendmail(remitente, destinatario, email.as_string())
     smtp.quit()


# Conectar con la base de datos
cnxn = pyodbc.connect(os.getenv('CONN'))
crsr = cnxn.cursor()

date = datetime.now()
date_report = datetime.today().strftime('%Y-%m-%d')
dir_sal = os.getenv('DIR_SAL')
mkdir(dir_sal+'/'+date_report)
dir_sal = dir_sal+'/'+date_report+'/'
# Consulta a la tabla que contiene los nombres de los vendors relacionados con sus respectivos correos
sql = "SELECT * FROM Email_Vendor"
crsr.execute(sql)
data = crsr.fetchall()
now = datetime.today().strftime('%m/%d/%Y')
date = datetime.strptime(now,"%m/%d/%Y")
last_date = (date - timedelta(days=1)).strftime('%m/%d/%Y')
route = os.getenv('ROUTE')

# Generar un dataframe con la data de la consulta anterior
df = pd.read_sql_query(sql, cnxn)
import re
def extraer_numero(celda):
    m = re.findall('\d+', celda)
    if m:
        print(m)
        return m[0]
    else:
        return celda

df_axis = pd.DataFrame(pd.read_excel(route)) 
df_axis = df_axis.rename({'Product Number US': 'Axis product part no'}, axis=1) 

# Iterar sobre las filas del dataframe
for index,row in df.iterrows(): 
    vendor = row['Vendor'] # Aquí están registrados los vendors
    email = row['Email'] # Aquí están registrados los correos, separados por una coma sin espacios

    # Consultar la vista según el vendor
    sql = "SELECT * FROM consultas_pos_"+vendor 
    crsr.execute(sql)
    data = crsr.fetchall()

    # Convertir consulta en un dataframe
    df = pd.read_sql_query(sql, cnxn) 
    df = df.set_index(df.columns[0])
    report = 'report_'+vendor+'_'+str(date_report)+'.xlsx'

    if vendor == 'LENOVO':
        if df['Serial Number'].isnull().any():
            df['QTY'] = df['QTY']
        else:
            df.loc[(df['Serial Number'] != '#') & (df['QTY'].astype(int)>0), 'QTY'] = 1
            df.loc[(df['Serial Number'] != '#') & (df['QTY'].astype(int)<0), 'QTY'] = -1
    
    if vendor == 'JUNIPER':
        df['Resale Invoice Date'] = pd.to_datetime(df['Resale Invoice Date'], errors='coerce').dt.strftime('%d-%b-%Y')
        if df['Product Serial Number'].isnull().any():
            df['Product Quantity'] = df['Product Quantity']
        else:
            df.loc[(df['Product Serial Number'] != '#') & (df['Product Quantity'].astype(int)>0), 'Product Quantity'] = 1
            df.loc[(df['Product Serial Number'] != '#') & (df['Product Quantity'].astype(int)<0), 'Product Quantity'] = -1

    if vendor == 'LENOVO_ISG':
        df['Año'] = pd.to_datetime(df['Invoice Date'], errors='coerce').dt.strftime('%Y')
        df['Mes'] = pd.to_datetime(df['Invoice Date'], errors='coerce').dt.strftime('%B')

        df.loc[(df['Mes'] == 'January') | (df['Mes'] == 'February') | (df['Mes'] == 'March'), 'Quarter'] = 'Q1' 
        df.loc[(df['Mes'] == 'April') | (df['Mes'] == 'May') | (df['Mes'] == 'June'), 'Quarter'] = 'Q2' 
        df.loc[(df['Mes'] == 'July') | (df['Mes'] == 'August') | (df['Mes'] == 'September'), 'Quarter'] = 'Q3' 
        df.loc[(df['Mes'] == 'December') | (df['Mes'] == 'November') | (df['Mes'] == 'October'), 'Quarter'] = 'Q4'  


        df['Week'] = pd.to_datetime(df['Invoice Date'], errors='coerce').dt.isocalendar().week
        df.loc[df['Mes'] == 'January', 'Mes'] = "Enero"
        df.loc[df['Mes'] == 'February', 'Mes'] = "Febrero"
        df.loc[df['Mes'] == 'March', 'Mes'] = "Marzo"
        df.loc[df['Mes'] == 'April', 'Mes'] = "Abril"
        df.loc[df['Mes'] == 'May', 'Mes'] = "Mayo"
        df.loc[df['Mes'] == 'June', 'Mes'] = "Junio"
        df.loc[df['Mes'] == 'July', 'Mes'] = "Julio"
        df.loc[df['Mes'] == 'August', 'Mes'] = "Agosto"
        df.loc[df['Mes'] == 'September', 'Mes'] = "Septiembre"
        df.loc[df['Mes'] == 'October', 'Mes'] = "Octubre"
        df.loc[df['Mes'] == 'November', 'Mes'] = "Noviembre"
        df.loc[df['Mes'] == 'December', 'Mes'] = "Diciembre"

    if vendor == 'AMD':
        df['Year'] = pd.to_datetime(df['Date (MM/DD/AA)'], errors='coerce').dt.strftime('%Y')
        df['Month'] = pd.to_datetime(df['Date (MM/DD/AA)'], errors='coerce').dt.strftime('%B')
        #pd.to_datetime(df['Date (MM/DD/AA)'], errors='coerce').dt.strftime('%M')
        df.loc[(df['Month'] == 'January') | (df['Month'] == 'February') | (df['Month'] == 'March'), 'Quarter'] = 'Q1' 
        df.loc[(df['Month'] == 'April') | (df['Month'] == 'May') | (df['Month'] == 'June'), 'Quarter'] = 'Q2' 
        df.loc[(df['Month'] == 'July') | (df['Month'] == 'August') | (df['Month'] == 'September'), 'Quarter'] = 'Q3' 
        df.loc[(df['Month'] == 'December') | (df['Month'] == 'November') | (df['Month'] == 'October'), 'Quarter'] = 'Q4'  

        df.loc[df['Month'] == 'January', 'Month'] = "Enero"
        df.loc[df['Month'] == 'February', 'Month'] = "Febrero"
        df.loc[df['Month'] == 'March', 'Month'] = "Marzo"
        df.loc[df['Month'] == 'April', 'Month'] = "Abril"
        df.loc[df['Month'] == 'May', 'Month'] = "Mayo"
        df.loc[df['Month'] == 'June', 'Month'] = "Junio"
        df.loc[df['Month'] == 'July', 'Month'] = "Julio"
        df.loc[df['Month'] == 'August', 'Month'] = "Agosto"
        df.loc[df['Month'] == 'September', 'Month'] = "Septiembre"
        df.loc[df['Month'] == 'October', 'Month'] = "Octubre"
        df.loc[df['Month'] == 'November', 'Month'] = "Noviembre"
        df.loc[df['Month'] == 'December', 'Month'] = "Diciembre"

        df = df.set_index(df.columns[0])
               
    if vendor == 'ZEBRA':
        if df['Serial Number'].isnull().any():
            df['Quantity'] = df['Quantity']
        else:
            df.loc[(df['Serial Number'] != '#') & (df['Quantity'].astype(int)>0), 'Quantity'] = 1
            df.loc[(df['Serial Number'] != '#') & (df['Quantity'].astype(int)<0), 'Quantity'] = -1

    if vendor == 'DELL':
        if df['Service Tag #'].isnull().any():
            df['Quantity Sold/ Sales Returns'] = df['Quantity Sold/ Sales Returns']
        else:
            df.loc[(df['Service Tag #'] != '#') & (df['Quantity Sold/ Sales Returns'].astype(int)>0), 'Quantity Sold/ Sales Returns'] = 1
            df.loc[(df['Service Tag #'] != '#') & (df['Quantity Sold/ Sales Returns'].astype(int)<0), 'Quantity Sold/ Sales Returns'] = -1

    if vendor == 'PLANTRONICS':
        if df['Serial Number'].isnull().any():
            df['Quantity Sold'] = df['Quantity Sold']
            df['Total Claim Amount'] = df['Total Claim Amount'] 
        else:
            df['Total Claim Amount'] = df['Total Claim Amount']/df['Quantity Sold']
            df.loc[(df['Serial Number'] != '#') & (df['Quantity Sold'].astype(int)>0), 'Quantity Sold'] = 1
            df.loc[(df['Serial Number'] != '#') & (df['Quantity Sold'].astype(int)<0), 'Quantity Sold'] = -1
        
        df.loc[df['Offer Code/NST 1 - 5 (can submit in 1 field)'] == 'FONDOS PLANTRONICS','Total Claim Amount'] = '' 
        df.loc[df['Offer Code/NST 1 - 5 (can submit in 1 field)'] == 'FONDOS PLANTRONICS', 'Offer Code/NST 1 - 5 (can submit in 1 field)'] = ''          
    
    if vendor == '3NSTAR':
        df = df.iloc[:, :-1]

    if vendor == 'AXIS':
        #print(df)
        df.loc[(df['Axis Project ref no'].str.contains("PROJ", case=False)), 'Partner level'] = 'PROJECT/'
        df['Axis Project ref no'] = df['Axis Project ref no'].map(extraer_numero)

        df.loc[(df['Axis Project ref no'] == 'AXIS_ PARTNERREBATE') | (df['Axis Project ref no'] == 'AXIS_PARTNERREBATE') | (df['Axis Project ref no'] == 'AXIS PRICE LIST') | (df['Axis Project ref no'] == '0'), 'Axis Project ref no'] = ''
        df_temp = pd.merge(df, df_axis, on='Axis product part no')
        df_temp['Purchase price per product'] = df_temp['Distributor Price']
        df_temp['Price per product after rebate'] = df_temp['Purchase price per product'] - df_temp['Rebate per product']
        
        #print(df_temp['Rebate per product'])
        #print(df_temp['Solution partner gold Rebate'])
        df_temp.loc[(df_temp['Rebate per product'] == round(df_temp['Multi-Regional partner Rebate'],2)), 'Partner level'] = 'MP/'
        df_temp.loc[(df_temp['Rebate per product'] == round(df_temp['Solution partner gold Rebate'],2)), 'Partner level'] = 'SG/'
        df_temp.loc[(df_temp['Rebate per product'] == round(df_temp['Solution partner silver Rebate'],2)), 'Partner level'] = 'SS/'
        df_temp.loc[(df_temp['Rebate per product'] == round(df_temp['Authorized partner Rebate'],2)), 'Partner level'] = 'AP/'
        
        df_temp = df_temp.drop(df_temp.columns[[16,17,18,19,20,21,22,23,25,25,26,27,28,29,30,31,32,33,34,35]], axis='columns')
        df = df_temp
        df. insert(loc = 0 , column = 'Distributor ID', value = 'INGRAM MICRO COLOMBIA')
        df = df.iloc[: , :-1]
        
        df = df.set_index(df.columns[0])
           
    if vendor == 'BELKIN':
        df['SALES QTY'] = df['SALES QTY'].replace("-","")

    if vendor == 'TARGUS_INVENTARIO':
        df = df.groupby(by=["Distributor Country", "Inventory Date", "Part Code", "Product Description"]).sum()

    if vendor == '3NSTAR_INVENTARIO':
        df = df.groupby(by=["FECHA", "PRODUCTO", "DESCRIPCION"]).sum()

    if vendor == 'WESTERN':
        fila_unica = {'WESTERN_DIGITAL_DISTRIBUTOR_ID':'9000003232','DISTRIBUTOR_NAME':'Ingram Micro SAS Colombia','DISTRIBUTOR_REGIONAL_OFFICE':'Colombia','DISTRIBUTOR_PART_NUMBER':'','WESTERN_DIGITAL_PART_NUMBER':'','PRODUCT_DESCRIPTION':'','QUANTITY_SOLD':'','SHIP_TO_RESELLER_ID':'','SHIP_TO_RESELLER_NAME':'','SHIP_TO_RESELLER_ADDRESS':'','SHIP_TO_RESELLER_CITY':'','SHIP_TO_RESELLER_STATE_OR_TERRITORY_OR_PROVINCE_OR_COUNTY':'','SHIP_TO_ZIP_CODE_OR_POSTAL_CODE':'','SHIP_TO_COUNTRY':'','INVOICE_NUMBER':'','INVOICE_DATE':last_date,'SELL_PRICE_US_CURRENCY':'','SELL_PRICE_DOMESTIC_CURRENCY':'','BILL_TO_RESELLER_ID':'','BILL_TO_RESELLER_NAME':'','BILL_TO_RESELLER_ADDRESS':'','BILL_TO_RESELLER_CITY':'','BILL_TO_RESELLER_STATE_OR_TERRITORY_OR_PROVINCE_OR_COUNTY':'','BILL_TO_ZIP_CODE_OR_POSTAL_CODE':'','BILL_TO_COUNTRY':'','COMMENTS':'','SPA_NUMBER':''}
        df = df.append(fila_unica, ignore_index=True)
        df = df.reindex(columns=['WESTERN_DIGITAL_DISTRIBUTOR_ID','DISTRIBUTOR_NAME','DISTRIBUTOR_REGIONAL_OFFICE','DISTRIBUTOR_PART_NUMBER','WESTERN_DIGITAL_PART_NUMBER','PRODUCT_DESCRIPTION','QUANTITY_SOLD','SHIP_TO_RESELLER_ID','SHIP_TO_RESELLER_NAME','SHIP_TO_RESELLER_ADDRESS','SHIP_TO_RESELLER_CITY','SHIP_TO_RESELLER_STATE_OR_TERRITORY_OR_PROVINCE_OR_COUNTY','SHIP_TO_ZIP_CODE_OR_POSTAL_CODE','SHIP_TO_COUNTRY','INVOICE_NUMBER','INVOICE_DATE','SELL_PRICE_US_CURRENCY','SELL_PRICE_DOMESTIC_CURRENCY','BILL_TO_RESELLER_ID','BILL_TO_RESELLER_NAME','BILL_TO_RESELLER_ADDRESS','BILL_TO_RESELLER_CITY','BILL_TO_RESELLER_STATE_OR_TERRITORY_OR_PROVINCE_OR_COUNTY','BILL_TO_ZIP_CODE_OR_POSTAL_CODE','BILL_TO_COUNTRY','COMMENTS','SPA_NUMBER'])
        df = df.set_index(df.columns[0])
        datew = (date - timedelta(days=1)).strftime('%Y%m%d')
        report = 'POS_WESTERN_9000003232_'+datew+'.xlsx'

    if vendor == 'WESTERN_INVENTARIO':
        fila_unica = {'WESTERN_DIGITAL_DISTRIBUTOR_ID': '9000003232','DISTRIBUTOR_NAME': 'Ingram Micro SAS Colombia','DISTRIBUTOR_REGIONAL_WHS': 'Colombia','DISTRIBUTOR_PART_NUMBER': '','WESTERN_DIGITAL_PART_NUMBER': '','PRODUCT_DESCRIPTION': '','CURRENT_INVENTORY': '','PRIME_INVENTORY': '','DEFECTIVE_INVENTORY': '','UNITS_RECEIVED_FROM_WD': '','PRIME_RETURN_UNITS_RECEIVED': '','UNITS_RETURNED_TO_WD': '','DISTRIBUTOR_IN_TRANSIT_INTERNAL_WHS': '','REPORT_DATE': last_date}
        df = df.append(fila_unica, ignore_index=True)
        df = df.reindex(columns=['WESTERN_DIGITAL_DISTRIBUTOR_ID','DISTRIBUTOR_NAME','DISTRIBUTOR_REGIONAL_WHS','DISTRIBUTOR_PART_NUMBER','WESTERN_DIGITAL_PART_NUMBER','PRODUCT_DESCRIPTION','CURRENT_INVENTORY','PRIME_INVENTORY','DEFECTIVE_INVENTORY','UNITS_RECEIVED_FROM_WD','PRIME_RETURN_UNITS_RECEIVED','UNITS_RETURNED_TO_WD','DISTRIBUTOR_IN_TRANSIT_INTERNAL_WHS','REPORT_DATE'])
        df = df.set_index(df.columns[0])
        datewi = (date - timedelta(days=1)).strftime('%Y%m%d')
        report = 'INV_WESTERN_9000003232_'+datewi+'.xlsx'
  
    if vendor == 'SANDISK':
        fila_unica = {'Partner Code':'2005316','Partner name':'Ingrammicro Colombia SAS','Transaction Date':last_date,'SanDisk SKU':'','Sales Quantity':'','Invoice Nbr':'','Reseller Name':'','Reseller Number':'','Reseller Address':'','City':'','ZipCode':'','Reseller Group':'','Reseller Country Code':'','Sales Channel':''}
        df = df.append(fila_unica, ignore_index=True)
        df = df.reindex(columns=['Partner Code','Partner name','Transaction Date','SanDisk SKU','Sales Quantity','Invoice Nbr','Reseller Name','Reseller Number','Reseller Address','City','ZipCode','Reseller Group','Reseller Country Code','Sales Channel'])
        df = df.set_index(df.columns[0])
        dates = (date - timedelta(days=1)).strftime('%Y%m%d')
        report = 'POSSANDISK_'+dates+'_'+dates+'.xlsx'
  
    if vendor == 'SANDISK_INVENTARIO':
        fila_unica = {'Partner Code':'2005316','Partner Name':'Ingrammicro Colombia SAS','SanDisk Part Number':'','Product Description':'','Quantity on Hand':'','Quantity in Backorder':'','Quantity Returned to SanDisk':'','Quantity on Order':'','Effective Inventory Date':last_date}
        df = df.append(fila_unica, ignore_index=True)
        df = df.reindex(columns=['Partner Code','Partner Name','SanDisk Part Number','Product Description','Quantity on Hand','Quantity in Backorder','Quantity Returned to SanDisk','Quantity on Order','Effective Inventory Date'])
        df = df.set_index(df.columns[0])
        datesi = (date - timedelta(days=1)).strftime('%Y%m%d')
        report = 'INVSANDISK_'+datesi+'_'+datesi+'.xlsx'
     
    
    


    # Guardar dataframe en carpeta Salida como un archivo tipo excel
    df.to_excel(dir_sal+report)

    # Separar string de emails y convertir en lista
    email = email.split(sep=',')
    emaillist = [elem.strip().split(',') for elem in email]
    #
    ## Definir contenido del correo
    msg = MIMEMultipart()
    msg['Subject'] = row['Subject']
    msg['From'] = os.getenv('REM_EMAIL')
    body = row['Text'] # Añadir texto en el cuerpo del mensaje
    msg.attach(MIMEText(body, 'plain')) 
    #
    route_att = dir_sal+report # Ruta del archivo 
    name_att = report # Asignar nombre al archivo
    file_att = open(route_att, 'rb') # Abrir el archivo
#
    object_MIME = MIMEBase('application', 'octet-stream') # Adjunto
    object_MIME.set_payload((file_att).read()) # Cargar el archivo
    encoders.encode_base64(object_MIME) # Codificar en BASE64
    object_MIME.add_header('Content-Disposition', "attachment; filename= %s" % name_att) # Añadir cabecera al objeto de MIME
    user = os.getenv('LOGIN_USER')
    password = os.getenv('LOGIN_PASS')
    msg.attach(object_MIME) # Agregar adjunto al mensaje
#
    try:
        server = smtplib.SMTP('smtp.outlook.com', 587)  # COMPROBAR ERRORES DE CONEXIÓN
        server.starttls()
        server.login(user,password)

        server.sendmail(msg['From'], emaillist , msg.as_string())
        server.close()
        print('Correo enviado para ',vendor)
        print('Correo enviado para ',vendor)

    except Exception as e:
        sendEmailError("CORREO PARA "+vendor)
    
