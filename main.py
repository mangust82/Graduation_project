""" основной модуль в котором вызываются методы из других программ
load_chan_plan модуль загружает файл канального плана с confluence в целевую папку
move_conf_file модуль перемещает файл канального плана в рабочую папку
start_macros модуль запускает файл Excel и вызывает макрос vb в нем на исполнение 
start_conf_downl модуль вызывает методы которые считывают файл настроек и запускают загрузку конфигурационных 
файлов на исполнение
"""
from load_channel_plan import load_chan_plan, move_conf_file
from start_macros import start_macros
from Correspod_prefix_to_conffile import dict_conf_file
from multithread_load import start_conf_downl
from pathlib import Path


SETTIGS_FILE = 'List_G2X_for_DVBC.xls'
URL_CONFLUENCE = 'https://confluence.mts.ru/XXXXXXX/XXXXXXXXXXXXXXXXXXX'
MASK = 'ЦТВ_канальный_и_пакетный_план_*.xlsm'
DEFAULT_LOAD_FOLDER = Path(r'C:\Users\username\Downloads')

if __name__ == '__main__':
    # load_chan_plan(URL_CONFLUENCE)
    path_chan_file = move_conf_file(MASK, DEFAULT_LOAD_FOLDER)
    start_macros(path_file=path_chan_file)
    dir_chan_file = Path(path_chan_file).parent
    start_conf_downl(SETTIGS_FILE, dir_chan_file)

