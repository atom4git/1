#!/usr/bin/env python3
import xlrd, xlwt, datetime, re
from xlutils.copy import copy
import json


colors_json = open('colors.json')
colors_dict = json.load(colors_json)
# variables:
brand = "Glossi"
t = datetime.datetime.today().strftime("%d%m%Y")
list_date = brand + t

n = 0
k = 0
list = []
out_list = []

def find_color(name):
    for k, v in colors_dict.items():
        for i in v:
            if "светло-" in name.lower() and i in name.lower():
                return "светло-" + k
            if "темно-" in name.lower() and i in name.lower():
                return "темно-" + k
            if i in name.lower():
                return k


def open_exel_file(file, sheet=0):
    """
    function take data fron exs file fnd return list
    :param file:
    :param sheet:
    :return: list
    """
    xls_book = xlrd.open_workbook(file)
    xls_sheet = xls_book.sheet_by_index(sheet)
    for row_num in range(xls_sheet.nrows):
        list.append(xls_sheet.row_values(row_num))
    return list


def make_data_name(data, n, k):
    """
    This module take list with data and convert it
    :param data:
    :param n:
    :return: out_list
    """

    for i in data:
        # генерируем первую строку(head)
        if n < 1:
            # k += 1
            # i.insert(0, " N/N")
            out_list.append(i)  # добавляем 1 строку в лист
        # остальный строки()
        else:
            size = str(i[6]).split(" ")  # получение листа размеров
            # получение списка размеров с их кол-вом dict
            count = {}  # create empty dict
            for character in size:
                if character != "" and character != " ":  # проверка наличия пуcтых значений
                    if str(character).isalpha() and str(
                            character).islower():  # проверка на нижний регистр букв в размерах типа "XL"
                        character = str(character).upper()
                    count.setdefault(character, 0)
                    count[character] += 1
            name_tmp = str(i[2]).strip().capitalize() + " " + str(i[3] + " "  + " " + str(i[5]))# + str(i[4])
            name_tmp = " ".join(name_tmp.split())  # проверка на наличие лишних пробелов в имени
            look_tmp = name_tmp.split() # получение вида товара
            look = str(look_tmp[0]).lower() # получение вида товара
            for key, value in count.items():
                k += 1
                name = name_tmp + " " + key
                art = " ".join(str(i[5]).strip().split())  # проверка на наличие лишних пробелов в артикуле
                out_list.append([k, i[1], name, i[3], i[4], str(art), (key), i[7], i[8], value, "=I2*J2", i[11], find_color(name), (key), " ", look, "",  brand])

        n += 1
    return out_list


def write_exel_file(output_file, name_of_sheet):
    """
    This module take list with data and make excel file
    :param output_file:
    :param name_of_sheet:
    :return: Excel file
    """
    x, y = -1, -1
    work_book = xlrd.open_workbook(output_file, formatting_info=True)
    new_book = copy(work_book)
    for l in out_list:
        x += 1
        y = 0
        for item in l:
            new_book.get_sheet(0).write(x, y, item)
            y += 1
    new_book.save(output_file)


open_exel_file("11.xls")
out_list = (make_data_name(list, n, k))
try:
    write_exel_file("12.xls", brand)
    print("Sucesful!")
except PermissionError:
    print("Закройте ваш файл EXCEL!")


