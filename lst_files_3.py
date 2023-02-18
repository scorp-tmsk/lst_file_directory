import os
import time
import pandas as pd
import openpyxl
from docxtpl import DocxTemplate
from docx import Document


def get_list_files(*args) -> list:
    """возращает список файлов в директории с нужным расширением. по умолчанию
    ('docx', 'doc', 'xls', 'xlsx', 'pptx', 'ppt', 'txt',)"""

    if not args:
        lst_type = ('docx', 'doc', 'xls', 'xlsx', 'pptx', 'ppt', 'txt',)
    else:
        lst_type = args

    lst_paths = []
    folder = os.getcwd()  # определяем корневой каталог
    for root, dirs, files in os.walk(folder):
        for file in files:
            dict_atr_file = {}
            if file.endswith(lst_type) and not file.startswith('~'):
                dict_atr_file.setdefault('Имя файла', os.path.join(root, file))
                dict_atr_file.setdefault('Формат файла', (os.path.splitext(file)[1]))
                dict_atr_file.setdefault('Дата создания', get_create_data_file(os.path.join(root, file)))
                dict_atr_file.setdefault('Пометка конфиденциальности', get_mark_file(file))
                lst_paths.append(dict_atr_file)
    return lst_paths


def get_mark_file(file_name: str) -> str:
    """Определяет категорию файла"""
    if file_name.startswith('0'):
        return 'Гриф конфидециальности 1'
    elif file_name.startswith('1'):
        return 'Гриф конфидециальности 2'
    else:
        return 'Общедоступно'


def get_create_data_file(file_path: str) -> str:
    """возращает дату  и время создания файла"""
    time_create = os.path.getctime(file_path)
    m_ti = time.ctime(time_create)
    t_obj = time.strptime(m_ti)
    t_stamp = time.strftime("%d-%m-%Y %H:%M:%S", t_obj)
    return t_stamp


def create_table_file(lst_files: list):
    """Создаем opis.docx таблицу по шаблону Normal"""
    doc = Document("Normal.docx")
    table = doc.tables[0]

    for ind, atr_file in enumerate(lst_files, start=1):
        table = table.table.add_row()
        row = table.table.rows[ind + 2]
        # запись данных в ячейки
        row.cells[0].text = str(ind)
        row.cells[1].text = atr_file['Имя файла']
        row.cells[2].text = atr_file['Формат файла']
        row.cells[3].text = atr_file['Дата создания']
        row.cells[4].text = atr_file['Пометка конфиденциальности']

    doc.save("Opis.docx")


def create_table_excel(lst_files: list):
    # строим opis.elsx таблицу из списка файлов сортируем по дате
    tb_lst_file = pd.DataFrame(lst_files)

    sor_tb_lst_file = tb_lst_file.sort_values(by='Дата создания').reset_index(drop=True)
    sor_tb_lst_file.index.rename('№ пп', inplace=True)
    sor_tb_lst_file.index += 1

    sor_tb_lst_file.to_excel('Opis.xlsx')


if __name__ == '__main__':
    lst_files = get_list_files()
    create_table_file(sorted(lst_files, key=lambda x: x['Дата создания']))
    create_table_excel(lst_files)
