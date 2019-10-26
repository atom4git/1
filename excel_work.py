import xlrd, xlwt, datetime
from xlutils.copy import copy

# variables:
brand = "Bogner"
t = datetime.datetime.today().strftime("%d%m%Y")
list_date = brand + t

n = 0
list = []
out_list = []


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


def make_data_name(data, n):
    """
    This module take list with data and convert it
    :param data:
    :param n:
    :return: out_list
    """
    for i in data:
        # генерируем первую строку(head)
        if n < 1:
            out_list.append(i)  # добавляем 1 строку в лист

        # остальный строки()
        else:
            size = i[2].split(" ")  # получение листа размеров
            # получение списка размеров с их кол-вом dict
            count = {}  # create empty dict
            for character in size:
                if character != "" and character != " ":  # проверка наличия пуcтых значений
                    count.setdefault(character, 0)
                    count[character] += 1
            name_tmp = str(i[0]).strip() + " " + i[1]
            for key, value in count.items():
                name = name_tmp + " " + key

                out_list.append([name, str(i[1]).strip(), key, value, i[4], i[5]])

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
out_list = (make_data_name(list, n))
try:
    write_exel_file("12.xls", brand)
except PermissionError:
    print("Закройте ваш файл EXCEL!")
