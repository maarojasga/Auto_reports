import os
from logging import error
from time import sleep

from dotenv import load_dotenv
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import WebDriverWait as wait
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from email.message import EmailMessage
import smtplib

config = load_dotenv(".env") # Usar .env

from datetime import datetime, timedelta

now = datetime.today().strftime('%m/%d/%Y')
start = datetime.strptime(now,"%m/%d/%Y")
end = now + " 12:00:00 A"
#end = '12/26/2022' + " 12:00:00 A"
start = (start - timedelta(days=7)).strftime('%m/%d/%Y') + " 12:00:00 A"
#start = '12/19/2022' + " 12:00:00 A"
# Credenciales
username = os.getenv('username') 
password = os.getenv('password') 
cop = 'COP'
usd = 'USD'

# Rango de fechas
start_date = start 
end_date = end 

# Credenciales correo
user = os.getenv('LOGIN_USER')
passw = os.getenv('LOGIN_PASS')
remitente = user
destinatario = os.getenv('LOGIN_USER')

# Función para enviar correos en caso de error
def sendEmailError(report):
     mensaje = "Algo salió mal, no se descargó "+report+". El error es: {}".format(e)
     email = EmailMessage()
     email["Subject"] = "ERROR DESCARGA "+report
     email.set_content(mensaje)
     smtp = smtplib.SMTP("smtp-mail.outlook.com", port=587)
     smtp.starttls()
     smtp.login(remitente, passw)
     smtp.sendmail(remitente, destinatario, email.as_string())
     smtp.quit()

# Función para enviar correos en caso de éxito
def sendEmailOK(report):
     mensaje = "Todo salió de maravilla con: "+report
     email = EmailMessage()
     email["Subject"] = "DESCARGADO "+report
     email.set_content(mensaje)
     smtp = smtplib.SMTP("smtp-mail.outlook.com", port=587)
     smtp.starttls()
     smtp.login(remitente, passw)
     smtp.sendmail(remitente, destinatario, email.as_string())
     smtp.quit()


try:
    try:
        path = os.path.dirname(os.path.abspath(os.getenv('DIR_DAILY_COP')+'prueba.txt'))
        prefs = {"download.default_directory":path}
        options = Options()
        options.add_experimental_option("prefs", prefs)

        driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
        # Login SAP
        driver.get("http://bobj/BOE/portal/1705152206/InfoView/logon.faces")

        # Encontrar campos de username y password, enviar datos
        driver.find_element(By.ID, "_id0:logon:USERNAME").send_keys(username)
        driver.find_element(By.ID, "_id0:logon:PASSWORD").send_keys(password)

        # Clic submit
        driver.find_element(By.ID,"_id0:logon:logonButton").click()

        # Esperar conexión
        WebDriverWait(driver=driver, timeout=150).until(
            lambda x: x.execute_script("return document.readyState === 'complete'")
        )

        error_message = "Credenciales inválidas."
        # Errores
        errors = driver.find_elements(By.CLASS_NAME,"logonError")

        if any(error_message in e.text for e in errors):
            print("No ingresaste")
        else:
            print("Ingreso exitoso")

        # Cambiar de html
        wait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH,'//*[@id="iframe4577-32469481"]')))
        # Encontrar daily
        my_documents = wait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH,'//a[@id="yui-gen0-1-label"]'))).click()

        sleep(2)
        find_daily = wait(driver, 40).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ListingURE_detailView_listColumn_1_0_1"]')))
        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(find_daily).perform()

        driver.switch_to.default_content()
        wait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH,'//*[contains(@id,iframe4583-32536351) and @role="presentation" and contains(@src,"http://bobj/BOE/portal/1705152206/AnalyticalReporting/WebiView.do")]')))

        print('vamos OK')

        find = wait(driver, 40).until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, 'iframe')))
        print('Todo OK')
        design = wait(driver, 40).until(EC.presence_of_element_located((By.XPATH,'//*[@id="_dhtmlLib_351"]')))

        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(design).perform()

        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@title="Edit Data Provider"]'))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="iconMenu_icon_queryPanelDlg_runQueryIconMenu"]'))).click()
        sleep(10)
        inputElement = wait(driver, 60).until(EC.visibility_of_element_located((By.ID, "lovWidgetpromptLovZone_searchTxt")))
        inputElement.send_keys(cop)

        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "_CWpromptsPromptListTWe_0"))).click()

        peso = wait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[contains(text(),"COP")]')))

        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(peso).perform()

        #wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys(start_date)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptsPromptListTWe_4"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "_CWpromptsPromptListTWe_0"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptsPromptListTWe_4"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys(start_date)
        sleep(5)
        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[contains(@aria-label,"Mandatory End Date: no values")]'))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptsPromptListTWe_4"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys('M')
        sleep(2)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptsPromptListTWe_5"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys(end_date)
        sleep(5)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys('M')

        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="RealBtn_OK_BTN_promptsDlg"]'))).click()
        sleep(700)

        a = driver.find_element(By.XPATH,'//td[@class="toolboxContSM"]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[10]/table')
        sleep(50)
        driver.execute_script("arguments[0].click()", a) #Click en elementos que se ocultan al encontrarlos :D
        sleep(5)

        driver.find_element(By.XPATH,"//select[@name='fileTypeList']/option[text()='Excel (.xlsx)']").click()
        sleep(5)
        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="OK_BTN_idExportDlg"]'))).click()
        sleep(50)

        driver.close
        sendEmailOK("DAILY SALES COP")
    except Exception as e:
        sendEmailError("DAILY SALES COP")


    try:
        path = os.path.dirname(os.path.abspath(os.getenv('DIR_DAILY_USD')+'prueba.txt'))
        prefs = {"download.default_directory":path}
        options = Options()
        options.add_experimental_option("prefs", prefs)

        driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)

        driver.get("http://bobj/BOE/portal/1705152206/InfoView/logon.faces")

        # Encontrar campos de username y password, enviar datos
        driver.find_element(By.ID, "_id0:logon:USERNAME").send_keys(username)
        driver.find_element(By.ID, "_id0:logon:PASSWORD").send_keys(password)

        # Clic submit
        driver.find_element(By.ID,"_id0:logon:logonButton").click()

        # Esperar conexión
        WebDriverWait(driver=driver, timeout=150).until(
            lambda x: x.execute_script("return document.readyState === 'complete'")
        )

        error_message = "Credenciales inválidas."
        # Errores
        errors = driver.find_elements(By.CLASS_NAME,"logonError")

        if any(error_message in e.text for e in errors):
            print("No ingresaste")
        else:
            print("Ingreso exitoso")

        # Cambiar de html
        wait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH,'//*[@id="iframe4577-32469481"]')))
        # Encontrar daily
        my_documents = wait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH,'//a[@id="yui-gen0-1-label"]'))).click()

        sleep(2)
        find_daily = wait(driver, 40).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ListingURE_detailView_listColumn_1_0_1"]')))
        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(find_daily).perform()

        driver.switch_to.default_content()
        wait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH,'//*[contains(@id,iframe4583-32536351) and @role="presentation" and contains(@src,"http://bobj/BOE/portal/1705152206/AnalyticalReporting/WebiView.do")]')))

        print('vamos OK')

        find = wait(driver, 40).until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, 'iframe')))
        print('Todo OK')
        design = wait(driver, 40).until(EC.presence_of_element_located((By.XPATH,'//*[@id="_dhtmlLib_351"]')))

        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(design).perform()

        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@title="Edit Data Provider"]'))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="iconMenu_icon_queryPanelDlg_runQueryIconMenu"]'))).click()
        sleep(10)
        inputElement = wait(driver, 60).until(EC.visibility_of_element_located((By.ID, "lovWidgetpromptLovZone_searchTxt")))
        inputElement.send_keys(usd)

        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "_CWpromptsPromptListTWe_0"))).click()

        peso = wait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[contains(text(),"USD")]')))

        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(peso).perform()

        #wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys(start_date)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptsPromptListTWe_4"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "_CWpromptsPromptListTWe_0"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptsPromptListTWe_4"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys(start_date)
        sleep(5)
        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[contains(@aria-label,"Mandatory End Date: no values")]'))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptsPromptListTWe_4"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys('M')
        sleep(2)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptsPromptListTWe_5"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys(end_date)
        sleep(5)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys('M')

        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="RealBtn_OK_BTN_promptsDlg"]'))).click()
        sleep(700)

        a = driver.find_element(By.XPATH,'//td[@class="toolboxContSM"]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[10]/table')
        sleep(50)
        driver.execute_script("arguments[0].click()", a) #Click en elementos que se ocultan al encontrarlos :D
        sleep(5)

        driver.find_element(By.XPATH,"//select[@name='fileTypeList']/option[text()='Excel (.xlsx)']").click()
        sleep(5)
        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="OK_BTN_idExportDlg"]'))).click()
        sleep(50)
        driver.close
        sendEmailOK("DAILY SALES USD")
    except Exception as e:
        sendEmailError("DAILY SALES USD")


    try:
        path = os.path.dirname(os.path.abspath(os.getenv('DIR_VEN')+'prueba.txt'))
        prefs = {"download.default_directory":path}
        options = Options()
        options.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
        # Login SAP
        driver.get("http://bobj/BOE/portal/1705152206/InfoView/logon.faces")
        # Encontrar campos de username y password, enviar datos
        driver.find_element(By.ID, "_id0:logon:USERNAME").send_keys(username)
        driver.find_element(By.ID, "_id0:logon:PASSWORD").send_keys(password)
        # Clic submit
        driver.find_element(By.ID,"_id0:logon:logonButton").click()
        # Esperar conexión
        WebDriverWait(driver=driver, timeout=150).until(
            lambda x: x.execute_script("return document.readyState === 'complete'")
        )
        error_message = "Credenciales inválidas."
        # Errores
        errors = driver.find_elements(By.CLASS_NAME,"logonError")
        if any(error_message in e.text for e in errors):
            print("No ingresaste")
        else:
            print("Ingreso exitoso")
        # Cambiar de html
        wait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH,'//*[@id="iframe4577-32469481"]')))
        # Encontrar daily
        my_documents = wait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH,'//a[@id="yui-gen0-1-label"]'))).click()
        sleep(2)
        find_daily = wait(driver, 40).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ListingURE_detailView_listColumn_4_0_1"]')))
        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(find_daily).perform()
        driver.switch_to.default_content()
        wait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH,'//*[contains(@id,iframe4583-32536351) and @role="presentation" and contains(@src,"http://bobj/BOE/portal/1705152206/AnalyticalReporting/WebiView.do")]')))
        print('vamos OK')
        find = wait(driver, 40).until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, 'iframe')))
        print('Todo OK')
        design = wait(driver, 40).until(EC.presence_of_element_located((By.XPATH,'//*[@id="_dhtmlLib_351"]')))
        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(design).perform()
        sleep(5)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "_CWpromptstrLstElt_TWe_2"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys(start_date)
        sleep(2)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptstrLstElt_TWe_3"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys(end_date)
        sleep(2)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "_CWpromptstrLstElt_TWe_2"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys('M')
        sleep(2)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "_CWpromptstrLstElt_TWe_4"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptstrLstElt_TWe_3"))).click()
        sleep(2)
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys('M')
        sleep(2)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptstrLstElt_TWe_3"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "_CWpromptstrLstElt_TWe_4"))).click()
        refresh = wait(driver, 40).until(EC.presence_of_element_located((By.XPATH,'//*[@id="lovWidgetpromptLovZone_refresh"]')))
        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(refresh).perform()
        sleep(5)
        company = wait(driver, 40).until(EC.presence_of_element_located((By.XPATH,'//*[contains(text(),"6410")]')))
        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(company).perform()
        sleep(2)

        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="RealBtn_OK_BTN_promptsDlg"]'))).click()
        sleep(700)

        a = driver.find_element(By.XPATH,'//td[@id="expZone__dhtmlLib_292"]/table/tbody/tr/td[9]/table')
        sleep(50)
        driver.execute_script("arguments[0].click()", a) #Click en elementos que se ocultan al encontrarlos :D
        sleep(5)

        driver.find_element(By.XPATH,"//select[@name='fileTypeList']/option[text()='Excel (.xlsx)']").click()
        sleep(5)
        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="OK_BTN_idExportDlg"]'))).click()
        sleep(50)

        driver.close
        sendEmailOK("VENDOR")
    except Exception as e:
        sendEmailError("VENDOR")

    try:
        path = os.path.dirname(os.path.abspath(os.getenv('DIR_INV')+'prueba.txt'))
        prefs = {"download.default_directory":path}
        options = Options()
        options.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
        driver.get("http://bobj/BOE/portal/1705152206/InfoView/logon.faces")
        # Encontrar campos de username y password, enviar datos
        driver.find_element(By.ID, "_id0:logon:USERNAME").send_keys(username)
        driver.find_element(By.ID, "_id0:logon:PASSWORD").send_keys(password)
        # Clic submit
        driver.find_element(By.ID,"_id0:logon:logonButton").click()
        # Esperar conexión
        WebDriverWait(driver=driver, timeout=150).until(
            lambda x: x.execute_script("return document.readyState === 'complete'")
        )
        error_message = "Credenciales inválidas."
        # Errores
        errors = driver.find_elements(By.CLASS_NAME,"logonError")
        if any(error_message in e.text for e in errors):
            print("No ingresaste")
        else:
            print("Ingreso exitoso")
        # Cambiar de html
        wait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH,'//*[@id="iframe4577-32469481"]')))
        # Encontrar daily
        my_documents = wait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH,'//a[@id="yui-gen0-1-label"]'))).click()
        sleep(2)
        find_daily = wait(driver, 40).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ListingURE_detailView_listColumn_3_0_1"]')))
        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(find_daily).perform()
        driver.switch_to.default_content()
        wait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH,'//*[contains(@id,iframe4583-32536351) and @role="presentation" and contains(@src,"http://bobj/BOE/portal/1705152206/AnalyticalReporting/WebiView.do")]')))
        print('vamos OK')
        find = wait(driver, 40).until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, 'iframe')))
        print('Todo OK')
        design = wait(driver, 40).until(EC.presence_of_element_located((By.XPATH,'//*[@id="_dhtmlLib_351"]')))
        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(design).perform()
        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@title="Edit Data Provider"]'))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="iconMenu_icon_queryPanelDlg_runQueryIconMenu"]'))).click()
        sleep(100)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).clear()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys(end_date)
        sleep(5)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptsPromptListTWe_1"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptsPromptListTWe_0"))).click()
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).click()
        wait(driver, 60).until(EC.presence_of_element_located((By.ID, "text_promptLovZone_RightZone_oneTextField_date0"))).send_keys('M')
        sleep(2)
        wait(driver, 60).until(EC.element_to_be_clickable((By.ID,"_CWpromptsPromptListTWe_1"))).click()
        wait(driver, 40).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="lovWidgetpromptLovZone_refresh"]'))).click()
        sleep(5)
        # double click operation and perform

        a = driver.find_element(By.XPATH,'//td[@id="_CWpromptslovWidgetpromptLovZone_lovTreetext_TWe_6"]/span')
        sleep(5)
        action = ActionChains(driver)
        # double click operation and perform
        action.double_click(a).perform()
        # double click operation and perform
        sleep(2)
        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="RealBtn_OK_BTN_promptsDlg"]'))).click()
        sleep(700)

        a = driver.find_element(By.XPATH,'//td[@class="toolboxContSM"]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[10]/table')
        sleep(50)
        driver.execute_script("arguments[0].click()", a) #Click en elementos que se ocultan al encontrarlos :D
        sleep(20)
        driver.find_element(By.XPATH,"//select[@name='fileTypeList']/option[text()='Excel (.xlsx)']").click()
        sleep(5)
        wait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="OK_BTN_idExportDlg"]'))).click()
        sleep(50)
        driver.close
        sendEmailOK("INVENTORY")
    except Exception as e:
        sendEmailError("INVENTORY")

    driver.delete_all_cookies()

    driver.quit()
    sendEmailOK("TODO")
except Exception as e:
    sendEmailError("TODO")

    driver.quit()
