import pyautogui as pag
from openpyxl import load_workbook
import time
import pandas as ps
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
import uuid
import win32api

wb=Workbook()


newnumber = 'порядковый номер'
def pag_prompt_new_number(newnumber):
    new_number=pag.prompt('Порядковый номер', 'Порядковый номер')
    newnumber=new_number
    return newnumber

newname = 'пример'
def pag_prompt_new_name(newname):
    new_name=pag.prompt('Введите имя нового сотрудника', 'Имя')
    newname=new_name
    return newname

newsurname = 'пример'
def pag_prompt_new_surname(newsurname):
    new_surname=pag.prompt('Введите фамилию нового сотрудника', 'Фамилия')
    newsurname=new_surname
    return newsurname

newbirthday = 'пример'
def pag_prompt_new_birthday(newbirthday):
    new_birthday=pag.prompt('Введите дату рождения нового сотрудника', 'Дата рождения')
    newbirthday=new_birthday
    return newbirthday

newpassportdata = 'пример'
def pag_prompt_new_passport_data(newpassportdata):
    new_passport_data=pag.prompt('Введите паспортные данные нового сотрудника', 'Паспортные данные')
    newpassportdata=new_passport_data
    return newpassportdata

newposition = 'пример'
def pag_prompt_new_position(newposition):
    new_position=pag.prompt('Введите должность нового сотрудника', 'Должность')
    newposition=new_position
    return newposition

newdepartment = 'пример'
def pag_prompt_new_department(newdepartment):
    new_department=pag.prompt('Введите отдел нового сотрудника', 'Отдел')
    newdepartment=new_department
    return newdepartment

#print ("The random id using uuid1() is : ",end="")
#print (uuid.uuid1())


#newid = 'пример'
#def pag_prompt_new_id(newid):
#    new_id=pag.prompt('На экране id этого сотрудника, можете записать newid', 'id')
#    newid = new_id
#    return newid



newpassword = 'пример'
def pag_prompt_new_password(newpassword):
    new_password=pag.password('На экране пароль нового сотрудника, можете записать', 'password')
    newpassword = new_password
    return newpassword

a=pag.confirm('Выберите что вы хотите сделать',
            'Управление данными персонала',
            buttons=['Новый сотрудник', 'Вывод данных старого сотрудника по ФИО', 'Удаление данных сотрудника', 'Изменение данных сотрудника'])




if a =='Новый сотрудник':


    newnumber=pag_prompt_new_number(newnumber)
    newname=pag_prompt_new_name(newname)
    newsurname=pag_prompt_new_surname(newsurname)
    newbirthday=pag_prompt_new_birthday(newbirthday)
    newpassportdata=pag_prompt_new_passport_data(newpassportdata)
    newposition=pag_prompt_new_position(newposition)
    newdepartment=pag_prompt_new_department(newdepartment)
    #newid=pag_prompt_new_id(newid)
    newid = uuid.uuid1() 
    win32api.MessageBox(0, 'На экране id этого сотрудника, можете записать', str(newid))
    newpassword=pag_prompt_new_password(newpassword)

elif a=='Вывод данных старого сотрудника по ФИО':
    print("a")
elif a=='Удаление данных сотрудника':
    print("a")
elif a=='Изменение данных сотрудника':
    print("a")
else: 
    pag.alert(text='Программа завершилась с неизвестной ошибкой, при возникновении вопросов обратитесь к создателю', title='Ошибка', button='OK')

print(newname)

wb = Workbook() # Стандартная инициализация книги Excel из openpyxl

dest_filename = 'empty_book3.xlsx' # Инициализация новой книги для сохранения (Сохраняется в директории с файлом .py)

ws = wb.active # Присваиваем переменную первому листу Excel

ws.title = "range names" # Даем имя первому листу Excel

# Начало функции присвоения листу значений

wb = load_workbook(filename='empty_book3.xlsx', read_only=False)
ws = wb['range names']

#print(ws['AA19'].value)

#ws.append(["Hello, Bitches"])



An = newnumber
Bn = newname
Cn = newsurname
Dn = newbirthday
En = newpassportdata
Fn = newposition
Gn = newdepartment
Hn = newid
In = newpassword

# Начало Конца функции присвоения листу значений

Spisok = tuple
Spisok = str(An), str(Bn), str(Cn), str(Dn), str(En), str(Fn), str(Gn), str(Hn), str(In)

print(Spisok)

ws.append(Spisok)



ws['A1'] = 'Порядковый номер сотрудника'
ws['B1'] = 'Имя'
ws['C1'] = 'Фамилия'
ws['D1'] = 'Дата рождения'
ws['E1'] = 'Паспортные данные'
ws['F1'] = 'Должность'
ws['G1'] = 'Отдел'
ws['H1'] = 'id'
ws['I1'] = 'password'
#ws['Y1'] = 
wb.save(filename = dest_filename) #Сохраняем данные в нашу книгу Excel
wb.close()
#C:\Users\Khodu\Desktop\Datalist.xlsx
exit()

  
