import os
import glob
import time
import shutil
import ntpath
from pathlib import PureWindowsPath, PurePosixPath
import base64
import json
from dotenv import load_dotenv
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from PyPDF2 import PdfFileMerger

import streamlit as st
from pathlib import Path


# getting credentials from environment variables
load_dotenv()
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD")
DOWNLOAD_PATH = r'{}'.format(os.getenv("DOWNLOAD_PATH"))
CHROME_PATH = r'{}'.format(os.getenv("CHROME_PATH"))
BANG_DIEM_PATH = r'{}'.format(os.getenv("BANG_DIEM_PATH"))

@st.cache(allow_output_mutation=True)
def load_options():
    # initialize the Chrome driver
    options = Options()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument('--headless')
    driver = webdriver.Chrome(executable_path=CHROME_PATH, chrome_options=options)

    # login page
    driver.get("https://trivietedu.ileader.vn/login.aspx")
    # find username/email field and send the username itself to the input field
    driver.find_element("id","user").send_keys(ADMIN_USERNAME)
    # find password input field and insert password as well
    driver.find_element("id","pass").send_keys(ADMIN_PASSWORD)
    # click login button
    driver.find_element(By.XPATH,'//*[@id="login"]/button').click()

    # click lop hoc
    lop_hoc_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="content"]/section/section/section/section/div/div[1]/div[1]/div/div[4]/a')))
    lop_hoc_button.click()

    # click nhap diem
    nhap_diem_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="menutop_nhapdiem"]/a/span[2]/span')))
    nhap_diem_button.click() 
    class_select = Select(WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="cp_lophoc"]'))))

    return driver, class_select

driver, class_select = load_options()


class_option = st.selectbox(
    'Class',
    tuple([class_name.text for class_name in class_select.options]))

class_select.select_by_visible_text(class_option)
# give some time for the webdriver to refresh the site after class selection
time.sleep(.1)
test_select = Select(WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="maudiem"]'))))

test_option = st.selectbox(
    'Test',
    tuple([test.text for test in test_select.options]))


PDFbyte = bytes('', 'utf-8')
placeholder = st.empty()
printing = placeholder.button('Confirm and Print',disabled=False, key='1')
if printing:
    placeholder.button('Confirm and Print', disabled=True, key='2')
    test_select.select_by_visible_text(test_option)
    time.sleep(1)
    rows = driver.find_elements(By.XPATH,"//table/tbody/tr")
    st.write("Combining", len(rows)-1)
    files = []
    for i in range(1, len(rows)):
        name = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH,f'//*[@id="dyntable"]/tbody/tr[{i}]/td[2]/div'))).text
        files.append(f'{name}.pdf')
        print_url = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH,f'//*[@id="dyntable"]/tbody/tr[{i}]/td[2]/a'))).get_attribute('href')
        chrome_options = webdriver.ChromeOptions()
        settings = {
            "recentDestinations": [{
                    "id": "Save as PDF",
                    "origin": "local",
                    "account": "",
                }],
                "selectedDestinationId": "Save as PDF",
                "version": 2
            }
        prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings)}
        chrome_options.add_experimental_option('prefs', prefs)
        chrome_options.add_argument('--kiosk-printing --headless')
        temp_driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=CHROME_PATH)
        temp_driver.get(print_url)
        pdf = temp_driver.execute_cdp_cmd("Page.printToPDF", {
        "printBackground": True
        }) 
        with open(f'{name}.pdf','wb') as f:
            f.write(base64.b64decode(pdf['data']))
        st.success(f'{name}', icon="???")     
        temp_driver.quit()

    ''' Merges all the pdf files in current directory '''
    merger = PdfFileMerger()
    st.write(files)
    #Iterate over the list of the file paths
    for pdf_file in files:
        #Append PDF files
        merger.append(pdf_file)
    merger.write(f"{class_option}-{test_option}.pdf")
    merger.close()
    files.append(f"{class_option}-{test_option}.pdf")
    time.sleep(1)
    with open(f"{class_option}-{test_option}.pdf", "rb") as pdf_file:
        PDFbyte = pdf_file.read()
    st.download_button(label="Download_PDF",
                    data=PDFbyte,
                    file_name=f"{class_option}-{test_option}.pdf",
                    mime='application/pdf')
    for f in files:
        os.remove(f)
    placeholder.button('Confirm and Print', disabled=False, key='3')
    placeholder.empty()

