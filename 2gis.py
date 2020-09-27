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
import sys 
 
# PROVIDE ".xlsx" FILE WITH 2GIS LINKS IN FIRST COLUMN
# EXAMPLE OF LINK: 'https://2gis.kz/nur_sultan/geo/9570784863367204'
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
gis_links_file = askopenfilename(filetypes=[("Excel files", ".xlsx")]) # show an "Open" dialog box and return the path to the selected file
gis_links_file = load_workbook(filename = gis_links_file)
gis_links_sheet = gis_links_file.active
row_count = gis_links_sheet.max_row
wbTarget = Workbook()
wsTarget = wbTarget.active

gis_links_to_search = []
# COLLECTING LINKS FROM SOURCE FILE ".xlsx"
print('Собираю ссылки 2ГИС с первого столбца файла ".xlsx"')
for i in range(1, row_count + 1): # '+1' to write the last row
    cell_value = gis_links_sheet.cell(row = i, column = 1).value
    gis_links_to_search.append(cell_value)

print('Начал работу по сбору данных\n\n')
browser = webdriver.Firefox()
street = ''
region = ''
phone = ''
timetable_data = []
headers = ["Регион",'Улица', 'Пон', "Вт", "Ср", "Чт", "Пт", "Сб", "Вс", "Все дни", "Тел", 'Ссылка']
 

# SAVE FILE
def save_exit():
    cwd = os.getcwd()
    splitedcwd = os.path.split(cwd)
    start = splitedcwd[0]
    end = splitedcwd[1]
    date = datetime.today().strftime('%Y-%m-%d %H-%M-%S')
     
    city_name = gis_links_to_search[0].split('/')
    filename =  f'{city_name[3]} 2gis_timetable'
    filetype = '.xlsx'
    filepath = os.path.join(start + '/' + end + '/' + filename + ' ' + date + filetype)
    wbTarget.save(filepath)
    wbTarget.close()
    gis_links_file.close()
    browser.quit()

    print()
    print(f'Закончил работу.\nДанные сохранены в папку "{start}\\{end}\\", файл "{filename} {date}{filetype}"', end='\n\n')
    sys.exit() # comment this line if you want program to go on


def get_region_street_phone():
    global street
    global region
    global phone
    street = browser.find_element_by_class_name('_er2xx9') 
    street = street.text
    region = browser.find_element_by_class_name('_1p8iqzw') 
    region = region.text
    phone = browser.find_element_by_css_selector("div._b0ke8 > a").get_attribute('href')


def make_clean_list(data_list):
    
    del data_list[:2]
    if 'время работы' in data_list:
        data_list.remove('время работы')  
        return data_list

    if 'обед' in data_list:      
        data_list.remove('обед')    
        return data_list
     
    return data_list   
     


def prepare_data_for_excel(data_list):
    '''
    Assigns company working hours to variables. Prepares data for writing to excel spreadsheet.
    
    After [.find_element_by_class_name('_18zamfw')] call not all data is catched. Some working days may be missing.
    That created different cases to work with. 
     
    Function may cause [TypeError: 'NoneType' object is not iterable] in [for loop] if none of conditions apply.

    Order of [if] statements matter. More specific [if] statements should come before less specific.
    '''

    global street
    global region
    global phone


# 15 ['Пн', '08:00–17:0012:00–13:00', 'Вт', '08:00–17:0012:00–13:00', 'Ср', '08:00–17:0012:00–13:00',
#  'Чт', '08:00–17:0012:00–13:00', 'Пт', '08:00–17:0012:00–13:00', 'Сб', '——', 'Вс', '08:00–14:00—', 
# 'прием анализов: пн-пт 8:00-17:00; вс 8:00-14:00']
    if len(data_list) == 15 and  ((len(data_list[1]) > 12)  and  ('Чт'==  data_list[6] and 'Пт' ==  data_list[8] and 'Вс' ==  data_list[-3])):
        lunch =  data_list[1][11:]
        mon = data_list[1][:11] 
        tue = data_list[3][:11] 
        wed = data_list[5][:11] 
        thu = data_list[7][:11]
        fri = data_list[9][:11]         
        sat = data_list[-4]       
        sun = data_list[-2]
        all_time = data_list[-1]
        data_list = [region, street, mon, tue, wed, thu, fri, sat, sun, all_time + '. Обед будни:' + lunch, phone]
        return data_list
     
    # working with ['Пн', '07:00–17:00', 'Вт', '07:00–17:00', 'Ср', '07:00–17:00', #
    # 'Чт', '07:00–17:00', 'Пт', '07:00–17:00', 'Сб', '08:00–13:00', 'Вс', '08:00–13:00', 
    #'забор крови: пн-пт 7:00-12:00; сб-вс 8:00-11:00, выдача результатов: пн-пт14:00-17:00; сб-вс 11:00-13:00'] 
    if len(data_list) == 15:
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5] 
        thu = data_list[7] 
        fri = data_list[9]
        sat = data_list[-4]  
        sun = data_list[-2]
        all_time = data_list[-1]
        data_list = [region, street, mon, tue, wed, thu, fri, sat, sun, all_time, phone] 
        return data_list

# 13 ['Пн', '08:00–17:0012:00–13:00', 'Вт', '08:00–17:0012:00–13:00', 'Ср', '08:00–17:0012:00–13:00',
#  'Чт', '08:00–17:0012:00–13:00', 'Пт', '08:00–17:0012:00–13:00', 'Вс', '08:00–14:00—', 
# 'прием анализов: пн-пт 8:00-17:00, сб 8:00-14:00; выдача анализов: пн-пт 8:00-17:00, сб 8:00-14:00']
    if len(data_list) == 13 and  ((len(data_list[1]) > 12 ) and ( 'Чт'==  data_list[6] and 'Пт'==  data_list[8] and 'Вс'==  data_list[-3])):
        lunch =  data_list[1][11:]
        mon = data_list[1][:11] 
        tue = data_list[3][:11] 
        wed = data_list[5][:11] 
        thu = data_list[7][:11]
        fri = data_list[9][:11]             
        sun = data_list[-2]
        all_time = data_list[-1]
        data_list = [region, street, mon, tue, wed, thu, fri, '', sun, all_time + '. Обед будни:' + lunch, phone]
        return data_list

# 13 ['Пн', '07:00–16:0012:00–13:00', 'Вт', '07:00–16:0012:00–13:00', 'Ср', '07:00–16:0012:00–13:00', 
# 'Чт', '07:00–16:0012:00–13:00', 'Сб', '08:00–14:00—', 'Вс', '——', 'прием анализов: пн-пт 7:00-16:00; сб 8:00-12:00']
    if len(data_list) == 13 and  (len(data_list[1]) > 12  and  'Чт'==  data_list[-7] and 'Сб' ==  data_list[-5] and 'Вс' ==  data_list[-3]):
        lunch =  data_list[1][11:]
        mon = data_list[1][:11] 
        tue = data_list[3][:11] 
        wed = data_list[5][:11] 
        thu = data_list[7][:11]        
        sat = data_list[-4]       
        sun = data_list[-2]
        all_time = data_list[-1]
        data_list = [region, street, mon, tue, wed, thu, '', sat, sun, all_time + '. Обед будни:' + lunch, phone]
        return data_list

#  13 ['Пн', '08:00–16:00', 'Вт', '08:00–16:00', 'Ср', '08:00–16:00',
#  'Чт', '08:00–16:00', 'Пт', 'Сб', '08:00–12:00', 'Вс', '—']
    if len(data_list) == 13 and  ( 'Пт' ==  data_list[-5] and 'Вс' ==  data_list[-2] ):
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5] 
        thu = data_list[7] 
        sat = data_list[-3]         
        sun = data_list[-1]
        data_list = [region, street, mon, tue, wed, thu, '', sat, sun, '', phone]
        return data_list

#  13 ['Пн', '08:00–16:00', 'Вт', '08:00–16:00', 'Ср', '08:00–16:00',
#  'Чт', '08:00–16:00', 'Пт', '08:00–16:00', 'Вс', '—', 'прием анализов: пн-пт 8:00-12:00; сб 9:00-11:00']
    if len(data_list) == 13 and  ('Пт'  == data_list[-5]):
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5] 
        thu = data_list[7] 
        fri = data_list[-4]         
        sun = data_list[-2]
        all_time =   data_list[-1]  
        data_list = [region, street, mon, tue, wed, thu, fri, '', sun, all_time, phone]
        return data_list



 # 13 ['Пн', '08:00–17:00', 'Вт', '08:00–17:00', 'Ср', '08:00–17:00', 'Чт', 'Пт', '08:00–17:00', 'Сб', '08:00–12:00', 'Вс', '—']
    if len(data_list) == 13 and  ('выдача' not in  data_list  and  'Пт' ==  data_list[7]):
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5] 
        fri = data_list[8] 
        sat = data_list[-3]  
        sun = data_list[-1]
            
        data_list = [region, street, mon, tue, wed, '', fri,  sat, sun, '', phone]
        return data_list


# 13 [ 'Пн', '08:00–12:00', 'Вт', '08:00–12:00', 'Ср', '08:00–12:00', 'Чт', '08:00–12:00', 'Пт', 'Сб', '—', 'Вс', '—']
    if len(data_list) == 13 and  'Сб' == data_list[9]:
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5] 
        thu = data_list[7] 
        sat = data_list[-3]  
        sun = data_list[-1]
         
        data_list = [region, street, mon, tue, wed, thu, '',  sat, sun, '', phone]
        return data_list


    if len(data_list) == 13 and ('Ср' in data_list):
    
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5] 

        sat = data_list[-4]  
        sun = data_list[-2]
        all_time = data_list[-1]

        if 'Чт' not in data_list:
            fri = data_list[-6]
            data_list = [region, street, mon, tue, wed, '', fri,  sat, sun, all_time, phone]
            return data_list

        if 'Пт' not in data_list:   
            thu = data_list[-6] 
            data_list = [region, street, mon, tue, wed, thu, '',  sat, sun, all_time, phone]
            return data_list

    # working with ['Пн', '07:00–17:00', 'Вт', '07:00–17:00', 'Чт', '07:00–17:00', 'Пт', '07:00–17:00', 'Сб', '08:00–13:00', 'Вс', '—', 'забор крови: пн-пт 
    # 7:00-12:00; сб 8:00-11:00; выдача результатов: пн-пт 12:00-17:00; сб 11:00-13:00']
    if len(data_list) == 13 and ('Ср' not in data_list):
        mon = data_list[1] 
        tue = data_list[3] 
        thu = data_list[5] 
        fri = data_list[7]
        sat = data_list[9]  
        sun = data_list[-2]
        all_time = data_list[-1]
        data_list = [region, street, mon, tue, '', thu, fri,  sat, sun, all_time, phone]
        return data_list

    if len(data_list) == 13 and ('Пт'in data_list and 'Чт' in data_list):
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5] 
        fri = data_list[-5]
        sat = data_list[-3]  
        sun = data_list[-1]
        data_list = [region, street, mon, tue, wed,  '' ,fri,  sat, sun, '', phone]
        return data_list

# work with 14 ['Пн', '07:00–15:00', 'Вт', '07:00–15:00', 'Ср', '07:00–15:00', 'Чт', 'Пт', '07:00–15:00', 'Сб', '08:00–13:00', 'Вс', '—', 'забор крови: 
# пн-пт 7:00-12:00; сб 8:00-11:00; выдача результатов: пн-пт 10:00-15:00; сб 10:00-13:00']
    if len(data_list) == 14 and data_list[7] == 'Пт':
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5] 
        fri = data_list[8]
        sat = data_list[10]
        sun = data_list[12]
        all_time = data_list[-1]
        data_list = [region, street, mon, tue, wed, '', fri,  sat, sun, all_time, phone]
        return data_list

#14 ['Пн', '07:00–20:00', 'Вт', '07:00–20:00', 'Ср', '07:00–20:00', 'Чт', '07:00–20:00',
#  'Пт', 'Сб', '08:00–20:00', 'Вс', '08:00–20:00', 
# 'забор крови: пн-пт 7:00-12:00; сб-вс 8:00-11:00; выдача результатов: пн-пт 11:00-20:00; сб-вс 10:00-20:00']
    if len(data_list) == 14 and data_list[9] == 'Сб':
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5] 
        thu = data_list[7]
        sat = data_list[10]
        sun = data_list[12]
        all_time = data_list[-1]
        data_list = [region, street, mon, tue, wed, thu, '',  sat, sun, all_time, phone]
        return data_list    

    if len(data_list) == 14:
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5] 
        thu = data_list[7] 
        fri = data_list[9]
        sat = data_list[11]  
        sun = data_list[13]
        #all_time = data_list[-1]
 
        data_list = [region, street, mon, tue, wed, thu, fri,  sat, sun, '', phone]
        return data_list
 
# 12 ['Пн', '08:00–16:00', 'Вт', '08:00–16:00', 'Ср', '08:00–16:00', 
# 'Чт', '08:00–16:00', 'Сб', '08:00–13:00', 'Вс', '—']
    if len(data_list) == 12 and ('Чт' == data_list[-6] and 'Сб' == data_list[-4]):
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5] 
        thu = data_list[7]
        sat = data_list[-3]  
        sun = data_list[-1]

        data_list = [region, street, mon, tue, wed, thu, '',  sat, sun, '', phone]
        return data_list

   # 12 ['Пн', '08:00–12:00', 'Вт', '08:00–12:00', 'Чт', '08:00–12:00', 
   # 'Пт', '08:00–12:00', 'Сб', '09:00–12:00', 'Вс', '—']     
    if len(data_list) == 12 and ('Ср' not in data_list ):
        mon = data_list[1] 
        tue = data_list[3] 
        thu = data_list[5] 
        fri = data_list[7]
        sat = data_list[9]  
        sun = data_list[11]

        data_list = [region, street, mon, tue, '', thu, fri,  sat, sun, '', phone]
        return data_list

# 12 ['Пн', '08:00–17:00', 'Вт', '08:00–17:00', 'Ср', '08:00–17:00', 'Пт', '08:00–17:00', 'Сб', '08:00–12:00', 'Вс', '—']
    if len(data_list) == 12 and ('Чт' not in data_list ):
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5] 
        fri = data_list[7]
        sat = data_list[9]  
        sun = data_list[11]

        data_list = [region, street, mon, tue, wed, '', fri,  sat, sun, '', phone]
        return data_list   

# 12 ['Пн', '08:00–18:00', 'Вт', '08:00–18:00', 'Ср', '08:00–18:00', 
# 'Чт', '08:00–18:00', 'Пт', '08:00–18:00', 'Сб', '08:00–18:00']
    if len(data_list) == 12 and ('Пт' ==  data_list[-4] and 'Сб' == data_list[-2]):
            mon = data_list[1] 
            tue = data_list[3] 
            wed = data_list[5]
            thu = data_list[7] 
            fri = data_list[-3]
            sat = data_list[-1]

            data_list = [region, street, mon, tue, wed, thu, fri,  sat, '', '', phone]
            return data_list 

# 12 ['Пн', '07:00–18:00', 'Вт', '07:00–18:00', 'Ср', '07:00–18:00', 'Чт', '07:00–18:00', 'Пт', '07:00–18:00', 'Вс', '08:00–13:00']         
    if len(data_list) == 12 and ('Пт' ==  data_list[-4] and 'Вс' == data_list[-2]):
        mon = data_list[1] 
        tue = data_list[3] 
        wed = data_list[5]
        thu = data_list[7] 
        fri = data_list[-3]
        sun = data_list[-1]

        data_list = [region, street, mon, tue, wed, thu, fri,  '', sun, '', phone]
        return data_list 
        

def get_gis_data(gis_link_to_search):
    # if gis_link_to_search.startswith('https://go.2gis'):
    #     browser.get(gis_link_to_search)
    #     browser.find_elements(By.CLASS_NAME, "_18zamfw")[1].click()

    try:

        global street
        global region
        global phone
        global timetable_data

        browser.get(gis_link_to_search)
        browser.find_elements_by_class_name('_z3fqkm')[1].click() # _z3fqkm - arrow; _18zamfw - block
        get_region_street_phone()

        # NOT ALL DATA IS CATCHED. SOME DAYS MAY BE MISSING
        timetable_data = browser.find_element_by_class_name('_18zamfw') 
        timetable_data = timetable_data.text
        timetable_data = timetable_data.split('\n')
        

        if len(timetable_data) != 1:        
            # No word 'фото' in first div._18zamfw
            timetable_data = make_clean_list(timetable_data)
            print(street, '\n', len(timetable_data),timetable_data, end='\n' )             
            timetable_data = prepare_data_for_excel(timetable_data)
        else:
            # If company has photo the array will only contain text 'фото'             
            # Has word 'фото' in first div._18zamfw
            timetable_data = browser.find_elements_by_class_name('_18zamfw')[1]
            timetable_data = timetable_data.text
            timetable_data = timetable_data.split('\n')
            
            timetable_data = make_clean_list(timetable_data)
            print(street, '\n', len(timetable_data), timetable_data, end='\n' )
            timetable_data = prepare_data_for_excel(timetable_data)       
          
    except:
        try:
            # Branch temporarily doesn't work -  Филиал временно не работает
            browser.find_element_by_class_name('_1xm5wvm').text # to check if there will be an error
            get_region_street_phone()
            timetable_data = [region, street, 'не работает', 'не работает', 'не работает', 'не работает', 'не работает', 'не работает', 'не работает', 'не работает', phone] 
            print(street, '\n', len(timetable_data), timetable_data, end='\n' )

        except:
             
            get_region_street_phone()
            timetable_data = browser.find_elements_by_class_name('_18zamfw')[0] 
            timetable_data = timetable_data.text
            timetable_data = timetable_data.split('\n')
            print(street , '\n', timetable_data,  end='\n' )

            # Works everyday             
            if any('Ежедневно' in s for s in timetable_data) and (len(timetable_data) == 3):
                timetable_data = [region, street, 'Ежедневно', 'Ежедневно', 'Ежедневно', 'Ежедневно', 'Ежедневно', 'Ежедневно', 'Ежедневно', timetable_data[-1], phone]
            elif any('Ежедневно' in s for s in timetable_data) and (len(timetable_data) == 2):
                timetable_data = [region, street, timetable_data[0], timetable_data[0], timetable_data[0], timetable_data[0], timetable_data[0], timetable_data[0], timetable_data[0], '', phone]
            else:
            # Works monday to friday
                timetable_data = make_clean_list(timetable_data)             
                timetable_data = prepare_data_for_excel(timetable_data)
                
            
        
         
def write_headers(headers):
    col = 1  
    for header in headers:
        wsTarget.cell(row=1, column=(col), value=f'{header}')  
        col += 1  
 

row = 2
def write_row(timetable_data_list):
    col = 1  
    for data in timetable_data_list:
        wsTarget.cell(row = row, column = col, value=f'{data}')  
        col += 1  
    
# MAIN LOOP 
write_headers(headers)


for gis_link in gis_links_to_search:
    try:
        get_gis_data(gis_link)
        write_row(timetable_data)
        wsTarget.cell(row = row, column = len(headers)   , value=f'{gis_link}')
        print(f'Обработано ссылок: {row - 1} из {len(gis_links_to_search)}.\n')
    except Exception as e:
        print('Ошибка:', e)
        print(f'Обработано ссылок: {row - 2} из {len(gis_links_to_search)}.')
        save_exit()    

    row += 1
   
save_exit()


# html = browser.page_source
# soup = BeautifulSoup(html,'html.parser' )
# print(soup.prettify())
