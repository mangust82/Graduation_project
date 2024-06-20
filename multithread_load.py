"""Этот модуль предназначен для загрузки конфигураций на устройства Geminus G2-X и IQ Cricket-QAM Plus,
используя информацию из CSV-файла. Он работает в многопоточном режиме для ускорения процесса.
"""
import csv
import threading
from pathlib import Path
from G2x_download_v3 import load_conf_G2X
from QAM_download_v1 import load_conf_QAM
from Correspod_prefix_to_conffile import dict_conf_file
import logging


filename = 'List_G2X_for_DVBC.xls'
dir_chan_file = r'C:\work\MVP\Scripts\dowload_conf\Chan_plan_140620241911'

logging.basicConfig(level=logging.INFO, filename="ETL.log", filemode="a",
                    format="%(asctime)s %(levelname)s %(message)s")

def process_G2X(probe_username, probe_password, IP_adr, conf_name):
    load_conf_G2X(probe_username, probe_password, IP_adr, conf_name)

def process_QAM(probe_username, probe_password, IP_adr, conf_name):
    load_conf_QAM(probe_username, probe_password, IP_adr, conf_name)

def start_conf_downl(file_settings:str, dir_chan_file:str) -> None:
    """The function reads the CSV file, parses the data, and starts threads to load configurations. 
    It uses `dict_conf_file` to get a mapping between the device name from the CSV file and the name 
    of the configuration file."""
    dict_file_name = dict_conf_file(file_settings, dir_chan_file)
    try:
        with open(filename, 'r', encoding='windows-1251', errors='ignore') as f:
            probe_f = csv.reader(f, dialect='excel-tab')
            threads = []
            for probe in probe_f:
                if probe[3] == '1':
                    IP_adr = probe[1]
                    conf_name = dict_file_name.get(probe[4])   #str(Path.cwd()) + '\\' + probe[4] + '.xls'
                    print(conf_name)
                    probe_username = probe[5]
                    probe_password = probe[6]
                    if probe[2] == 'Geminus G2-X':
                        thread = threading.Thread(target=process_G2X, args=(probe_username, probe_password, IP_adr, conf_name))
                        thread.start()
                        threads.append(thread)
                    elif probe[2] == 'IQ Cricket-QAM Plus':
                        thread = threading.Thread(target=process_QAM, args=(probe_username, probe_password, IP_adr, conf_name))
                        thread.start()
                        threads.append(thread)
        # Дождитесь завершения всех потоков
        for thread in threads:
            thread.join()
        logging.info(f"Все threads выполнены")

    except Exception as e:
        logging.error({type(e).__name__} - {str(e)}, exc_info=True)

if __name__ == '__main__':
    start_conf_downl(filename, dir_chan_file)