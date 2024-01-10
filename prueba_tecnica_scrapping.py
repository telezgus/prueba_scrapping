

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import pandas as pd
import zipfile


options = webdriver.EdgeOptions()
options.add_argument('--start-maximized')

driver_path = 'C:\\Users\\gusti\\Downloads\\edgedriver_win64\\msedgedriver.exe'

driver = webdriver.Edge(options=options)


driver.get('https://cde.ucr.cjis.gov/LATEST/webapp/#/pages/home')

WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable(By.LINK_TEXT, "Documents & Downloads")).click()

WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.ID,"dwnnibrs-download-select")))\
    .click()

WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.ID,"nb-option-71")))\
    .click()

WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.ID,"dwnnibrsloc-select")))\
    .click()

WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.XPATH,"//nb-option[contains(text(),'Florida')]")))\
    .click()

WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.ID,"nibrs-download-button")))\
    .click()


route_zip = 'C:\\Users\\gusti\\Downloads\\victims.zip'


excel_file = 'Victims_Age_by_Offense_Category_2022.xlsx'


with zipfile.ZipFile(route_zip, 'r') as zip_ref:
    with zip_ref.open(excel_file) as file:
        df = pd.read_excel(file,header=None)



df.replace('\n', ' ', regex=True, inplace=True)
df.replace('−', ' to ', regex=True, inplace=True)



df_index_title = df.iloc[3:4, 0].reset_index(drop=True)
df_index = df.iloc[13:25, 0].reset_index(drop=True)
column_to_list = df_index.tolist()
category_column = df_index_title[0]



df_headers = df.iloc[4:5, 2:15].reset_index(drop=True)
df_registers = df.iloc[13:25, 2:15].reset_index(drop=True)
df_col2 = pd.concat([df_headers, df_registers])
df_col2.columns = df_col2.iloc[0]  # Usa la fila 5 como nombres de columnas
df_col2 = df_col2.iloc[1:]  # Elimina la fila 5 del DataFrame


df_col2.insert(0,category_column, column_to_list)


df_col2.to_csv('crimes_against_property.csv', index=False)