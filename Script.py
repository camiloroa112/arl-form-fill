# 1st Party Libraries
import os
import json
import time

# 3rd Party Libraries
from selenium import webdriver
import edgedriver_autoinstaller
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Loading credentials
creds = json.load(os.getcwd().replace('\\', '/') + 'assets' + 'credentials.json')

# Auto install of web browser to initialize page connection
edgedriver_autoinstaller.install()

# Edge window shows and the same is maximized
driver = webdriver.Edge()
driver.maximize_window()

# Establishing connection with Alissta
driver.get("https://alissta.gov.co/AutoEvaluacionCOVID/COVID19")

#Selects Cédula and ID
doc_user = WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.ID, 'DocumentoUsuario')))
Select_Opt = Select(doc_user)

# Selecting Cédula from the dropdown
Select_Opt.select_by_visible_text("CÉDULA DE CIUDADANÍA")

# Entering Cedula ID number
search = driver.find_element_by_name("numDocumentoUsuario")
search.send_keys(creds["id"])

# Clicking on form button
driver.find_element_by_id("BuscarUsuario").click()

# Selection of company name
element = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.ID, "Empresa")))
select_elem2 = Select(element)
select_elem2.select_by_visible_text(creds["company_name"])

# Entering password and pressing on submit
password = driver.find_element_by_name("ClaveAcceso")
password.send_keys(creds["password"])
driver.find_element_by_xpath("//*[@id='datosIngreso']/div/div[2]/div/div[2]/button").click()

# Clicking on submit
element2 = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[3]/div[1]/div/div/div[5]/div/div[2]/button")))
element2.click()

# Filling form options
q1 = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'pregunta_1_2')))
driver.execute_script("arguments[0].click()", q1)
q2 = driver.find_element_by_id("pregunta_2_2")
driver.execute_script("arguments[0].click()", q2)
q4 = driver.find_element_by_id("pregunta_4_1")
driver.execute_script("arguments[0].click()", q4)
q5 = driver.find_element_by_id("pregunta_5_2")
driver.execute_script("arguments[0].click()", q5)
q6 = driver.find_element_by_id("pregunta_6_2")
driver.execute_script("arguments[0].click()", q6)
q7 = driver.find_element_by_id("pregunta_7_1")
driver.execute_script("arguments[0].click()", q7)
q8 = driver.find_element_by_id("pregunta_8_1")
driver.execute_script("arguments[0].click()", q8)
q9 = driver.find_element_by_id("pregunta_9_1")
driver.execute_script("arguments[0].click()", q9)
q10 = driver.find_element_by_id("pregunta_10_1")
driver.execute_script("arguments[0].click()", q10)
q11 = driver.find_element_by_id("pregunta_11_2")
driver.execute_script("arguments[0].click()", q11)
q12 = driver.find_element_by_id("pregunta_12_1")
driver.execute_script("arguments[0].click()", q12)
q13 = driver.find_element_by_id("pregunta_13_2")
driver.execute_script("arguments[0].click()", q13)
q14 = driver.find_element_by_id("pregunta_14_2")
driver.execute_script("arguments[0].click()", q14)
q15 = driver.find_element_by_id("pregunta_15_2")
driver.execute_script("arguments[0].click()", q15)

# Saving changes
driver.find_element_by_id("btnGuardar").click()

# Setting time for the user to visualize the form being complete
time.sleep(4)

# Submitting Alissta form
element3 = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[7]/div/button")))
element4 = element3.click()

# A message displays for the user thanking the completion of filling the form
time.sleep(8)

# Exiting page
driver.quit()

# Company Completion SharePoint ARL Form
driver2 = webdriver.Edge()

# Maximizing Browser window
driver2.maximize_window()

# Entering Sharepoint ARL Form URL
driver2.get(creds["form_url"])

# Clicking in Sharepoint ARL Form
element5 = WebDriverWait(driver2,8).until(EC.visibility_of_element_located((By.XPATH, "/html/body/form/div[12]/div[1]/div[1]/div[2]/div[2]/div[3]/div[1]/div/div/div/table/tbody/tr/td/table[1]/tbody/tr/td/a[1]/span[2]")))
element5.click()

# Entering name in Sharepoint ARL Form
enter_name = WebDriverWait(driver2,3).until(EC.presence_of_element_located((By.ID, "Title_fa564e0f-0c70-4ab9-b863-0177e6ddd247_$TextField")))
enter_name.send_keys(creds["full_name"])

# Pressing on save
driver2.find_element_by_id("ctl00_ctl33_g_debf31da_4637_47f7_906a_a45d577bc163_ctl00_toolBarTbl_RightRptControls_ctl00_ctl00_diidIOSaveItem").click()

# Time to allow user to visualize form completion
time.sleep(5)

# Exiting Sharepoint ARL Form
driver2.quit()