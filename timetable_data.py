
count_1 = ['фото']
count_18 = [
 'Сегодня c 07:00 до 14:00',
 'Закрыто. Откроется завтра в 07:00', 
 'Пн', '07:00–14:00', 
 'Вт', '07:00–14:00', 
 'Ср', 'время работы', '07:00–14:00', 
 'Чт', '07:00–14:00', 
 'Пт', '07:00–14:00', 
 'Сб', '07:00–14:00', 
 'Вс', '—', 
 'прием анализов: пн-сб 7:00-12:00; выдача анализов: пн-сб 7:00-14:00']

count_16 = [
    'Сегодня c 07:00 до 16:00', 
    'Закрыто. Откроется завтра в 07:00', 
    'Пн', '07:00–16:00', 
    'Вт', '07:00–16:00', 
    'Ср', 'время работы', '07:00–16:00', 
    'Чт', '07:00–16:00', 
    'Сб', '08:00–14:00', 
    'Вс', '—', 
    'прием анализов: пн-пт 7:00-12:00; сб 8:00-12:00; выдача анализов пн-пт 7:00-16:00; сб 8:00-14:00'
    ]       

count_17 =  [
    'Сегодня с 07:00 до 16:00, обед c 12:00 до 13:00', 
    'Закрыто. Откроется завтра в 07:00', 
    'Пн', '07:00–16:0012:00–13:00', 
    'Вт', '07:00–16:0012:00–13:00', 
    'Ср', 'время работы', 'обед', '07:00–16:0012:00–13:00', 
    'Пт', '07:00–16:0012:00–13:00', 
    'Сб', '08:00–14:00—', 
    'Вс', '——', 
    'прием анализов: пн-пт 7:00-16:00; сб 8:00-12:00'
    ]

count_19 =  [
    'Сегодня с 08:00 до 17:00, обед c 12:00 до 13:00', 
    'Закрыто. Откроется завтра в 08:00', 
    'Пн', '08:00–17:0012:00–13:00', 
    'Вт', '08:00–17:0012:00–13:00', 
    'Ср', 'время работы', 'обед', 
    '08:00–17:0012:00–13:00', 
    'Чт', '08:00–17:0012:00–13:00', 
    'Пт', '08:00–17:0012:00–13:00', 
    'Сб', '——', 
    'Вс', '08:00–14:00—', 
    'прием анализов: пн-пт 8:00-17:00; вс 8:00-14:00']

# del count_19[:2]
# count_19.remove('время работы')    
# count_19.remove('обед')    
# print(count_19)
def make_clean_list(dirty_list):
    
    del dirty_list[:2]
    if 'время работы' in dirty_list:
        dirty_list.remove('время работы')  
    if 'обед' in dirty_list:      
        dirty_list.remove('обед')    
    print(len(dirty_list) )    
    print(dirty_list)
    for i in dirty_list:
        print(i)

def make_time_only_list(clean_list):
    '''
    Assigns company Working hours to variables.
    '''

    if len(clean_list) == 15:
        mon = clean_list[1] 
        tue = clean_list[3] 
        wed = clean_list[5] 
        thu = clean_list[7] 
        fri = clean_list[9]
        sat = clean_list[-4]  
        sun = clean_list[-2]
        all_time = clean_list[-1]
        clean_list = [mon, tue, wed, thu, fri, sat, sun, all_time] 
        for i in clean_list:
            print(i)

    if len(clean_list) == 13:
        mon = clean_list[1] 
        tue = clean_list[3] 
        wed = clean_list[5] 

        sat = clean_list[-4]  
        sun = clean_list[-2]
        all_time = clean_list[-1]

        # Sometimes in data there are days that missing. This is a workaround.
        if 'Чт' not in clean_list:
            fri = clean_list[-6]
            clean_list = [mon, tue, wed, fri,  sat, sun, all_time]

        if 'Пт' not in clean_list:   
            thu = clean_list[-6] 
            clean_list = [mon, tue, wed, thu,  sat, sun, all_time]
        
         
        for i in clean_list:
            print(i)

make_clean_list(count_16)
make_time_only_list(count_16)
make_clean_list(count_17)
make_time_only_list(count_17)
make_clean_list(count_18)
make_time_only_list(count_18)
make_clean_list(count_19)
make_time_only_list(count_19)

# make_clean_list(count_17)
# make_clean_list(count_18)
# make_clean_list(count_19)
 