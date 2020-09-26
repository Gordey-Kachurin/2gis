from selenium import webdriver
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import cell
from openpyxl.styles import Alignment #, named_styles
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from selenium.webdriver.common.by import By

browser = webdriver.Firefox()

# TODO - MAKE TWO DIFFERENT FUNCTIONS
# if gis_link_to_search.startswith('https://go.2gis'):
#     browser.get(gis_link_to_search)
#     browser.find_elements(By.CLASS_NAME, "_18zamfw")[1].click()

# Provide branch url 
# KDL OLYMP BRANCH url in Almaty example: 'https://2gis.kz/almaty/branches/9429948590733484'
browser.get('https://2gis.kz/nur_sultan/branches/70000001018099222')
branch_count = browser.find_elements(By.CLASS_NAME, "_1p8iqzw")[1].text

for item in range(1,10):
    content_blocks = browser.find_elements_by_class_name('_vhuumw') # _oqoid - adress
    for block in content_blocks:
        browser.execute_script("arguments[0].scrollIntoView();"  , block )
  
raw_hrefs = []
clean_hrefs = []
 
for block in content_blocks:
    href = block.get_attribute('href')
    raw_hrefs.append(href)

for href in raw_hrefs:
    if 'firm' in str(href):
        clean_hrefs.append(href)

# for c in clean_hrefs:
#     print(c)
print(len(clean_hrefs), branch_count)

wb_urls = Workbook()
ws_urls = wb_urls.active

row = 1
def write_url(data): 
    ws_urls.cell(row = row, column = 1, value=f'{data}')  
        
for i in clean_hrefs:
    write_url(i)
    row += 1

# SAVE FILE
cwd = os.getcwd()
splitedcwd = os.path.split(cwd)
start = splitedcwd[0]
end = splitedcwd[1]
date = datetime.today().strftime('%Y-%m-%d %H-%M-%S')
filepath = os.path.join(start + '/' + end + '/' + date +  ' 2gis_urls.xlsx')
wb_urls.save(filepath)
wb_urls.close()
browser.quit()  