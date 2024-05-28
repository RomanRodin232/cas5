from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import datetime

element_size = (25, 1)

layout = [
    [sg.Text('Название товара', size=element_size), sg.Input(key='nazvanie', size=element_size)],
    [sg.Text('Количество товара', size=element_size), sg.Input(key='kolichestvo', size=element_size)],
    [sg.Text('Фамилия', size=element_size), sg.Input(key='familia', size=element_size)],
    [sg.Text('Имя', size=element_size), sg.Input(key='name', size=element_size)],
    [sg.Text('Отчество', size=element_size), sg.Input(key='otchestvo', size=element_size)],
    [sg.Text('e-mail', size=element_size), sg.Input(key='mail', size=element_size)],
    [sg.Text('Номер телефона', size=element_size), sg.Input(key='phone_number', size=element_size)],
    [sg.Text('Тип доставки', size=element_size), sg.Input(key='tip', size=element_size)],
    [sg.Text('Адрес доставки (заполнять если тип доставки "курьерская доставка")', size=(25, 3)), sg.Input(key='adres', size=(25, 3))],
    [sg.Button('Добавить'), sg.Button('Закрыть')]
]

window = sg.Window('Учет движения материалов на складе', layout, element_justification='center')

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Закрыть':
        break
    if event == 'Добавить':
        try:
            wb = load_workbook('sklad.xlsx')
            sheet = wb['Лист1']
            ID = len(sheet['ID'])
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            data = [
                ID,
                values['nazvanie'],
                values['kolichestvo'],
                values['familia'],
                values['name'],
                values['otchestvo'],
                values['mail'],
                values['phone_number'],
                values['tip'],
                values['adres'],
                time_stamp
            ]
            sheet.append(data)
            wb.save('sklad.xlsx')

            # Очистка полей ввода
            for key in values:
                window[key].update(value='')
            window['name'].set_focus()
            sg.popup('Данные сохранены')
        except PermissionError:
            sg.popup('Ошибка доступа', 'Файл используется другим пользователем.\nПопробуйте позже.')


window.close()