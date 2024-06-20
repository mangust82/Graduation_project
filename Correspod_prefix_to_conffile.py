"""Модуль выполняет сопоставление префиксам (котрые хранятся в файле настроек) файлам конфигурации 
    с указанием их полного пути.
"""
from pathlib import Path
import os
import csv
from datetime import datetime
import logging


settings_file = 'List_G2X_for_DVBC.xls'
dir = 'Alias'
path_conf_folder = 'C:\work\MVP\Scripts\dowload_conf\Chan_plan_140620241643'

logging.basicConfig(level=logging.INFO, filename="ETL.log", filemode="a",
                    format="%(asctime)s %(levelname)s %(message)s")

def check_time_update(file_old, file_new, path_folder):
    """searches for a more recent configuration file and returns it"""
    try:
        # Получение времени последнего изменения файла в секундах
        file_path_old = Path() / path_folder / file_old
        print(file_path_old)
        file_path_new = Path() / path_folder / file_new
        last_modified_old = os.path.getmtime(file_path_old)
        last_modified_new = os.path.getmtime(file_path_new)
        if last_modified_old >= last_modified_new:
            return file_old
        else:
            return file_new
    except Exception as e:
        logging.error({type(e).__name__} - {str(e)}, exc_info=True)

def dict_conf_file(filename:str, path_conf_folder:str) -> dict:
    """the name of the analyzer is matched to the configuration file"""
    try:
        # Получаем множество префиксов(маски) по котрым будем искать файлы конфигурации
        with open(filename, 'r', encoding='windows-1251', errors='ignore', newline='') as f:
            probe_f = csv.reader(f, dialect='excel-tab')
            next(probe_f, None)
            set_prefix = set()
            for line in probe_f:
                if line[4] != '':
                    set_prefix.add(line[4])
        # Получаем список файлов конфигурации            
        conf_list = os.listdir(path_conf_folder)
        # Сопоставляем маску файлам конфигурации и получаем словарь соответствия на выходе
        dict_file_name = {'IQAlias_G2X-MSK-Test2-B2':'C:\work\MVP\Scripts\dowload_conf\IQAlias_G2X-MSK-Test2-B2.xls', \
                        'IQAlias_IQ Cricket-QAM Plus':"C:\work\MVP\Scripts\dowload_conf\IQAlias_IQ Cricket-QAM Plus.xls"}
        for conf_name in conf_list:
            for prefix in set_prefix:
                # Проверяем есть ли уже ключ и если нет записываем найденное соответсвие
                if prefix == conf_name[:-16] and dict_file_name.get(prefix) == None:
                    dict_file_name[prefix] = str(Path() / path_conf_folder / conf_name)
                # Проверяем есть ли уже ключ и если есть выбираем какой файл записать в ключ
                elif prefix == conf_name[:-16] and dict_file_name.get(prefix) != None:
                    dict_file_name[prefix] =  str(Path() / path_conf_folder / check_time_update(dict_file_name.get(prefix), conf_name, path_conf_folder))
        logging.info(f'Перечень соответсвия устройсва файлу конфигурации {dict_file_name}')
        return dict_file_name
    except Exception as e:
         logging.error({type(e).__name__} - {str(e)}, exc_info=True)

if __name__ == '__main__':
    dict_file_name = dict_conf_file(settings_file, path_conf_folder)
    for item in dict_file_name.items():
        print(item)
