"""открывает файл канального плана и запускает на исполнение три макроса VBA которые создают 
файлы конфигурации для анализаторов
"""
import xlwings as xw
from pathlib import Path
import time
import logging


logging.basicConfig(level=logging.INFO, filename="ETL.log", filemode="a",
                    format="%(asctime)s %(levelname)s %(message)s")

path_chan_file = Path(r'C:\Users\userneme\Downloads\ЦТВ_канальный_и_пакетный_план_2024-06-10v2.xlsm')
def start_macros(path_file):
    """Open *.xlsm file and run three macroses config alias files. 
    The result of which is configuration files for analyzers"""
    try:
        vba_book = xw.Book(path_file)

        vba_macro = vba_book.macro("Module1.DVBC_CS_MPTS")
        vba_macro()
        logging.info(f"Макрос DVBC_CS_MPTS выполнен")

        vba_macro = vba_book.macro("Module2.DVBC_SPTS")
        vba_macro()
        logging.info(f"Макрос DVBC_SPTS выполнен")

        vba_macro = vba_book.macro("Module3.PSISI")
        vba_macro()
        logging.info(f"Макрос PSISI выполнен")

    except Exception as e:
        logging.error({type(e).__name__} - {str(e)}, exc_info=True)

if __name__ == '__main__':
    start_macros(path_chan_file)


# import win32com.client

# # Запускаем приложение Excel
# excel = win32com.client.Dispatch("Excel.Application")
# excel.Visible = True

# # Открываем книгу Excel
# workbook = excel.Workbooks.Open(r'C:\Users\yygolovy\Downloads\ЦТВ_канальный_и_пакетный_план_2024-06-10v2')

# # Получаем доступ к модулю VBA и запускаем макрос
# macro_name = "Module1.DVBC_CS_MPTS"  # Замените на имя вашего макроса
# excel.Run(macro_name)

# # Сохраняем и закрываем книгу
# workbook.Save()
# excel.Quit()
