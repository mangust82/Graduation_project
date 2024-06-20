"""load_chan_plan модуль загружает файл канального плана с confluence в целевую папку
move_conf_file модуль перемещает файл канального плана в рабочую папку
"""
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options 
import time
from pathlib import Path
from datetime import datetime
import logging


logging.basicConfig(level=logging.INFO, filename="ETL.log", filemode="a",
                    format="%(asctime)s %(levelname)s %(message)s")
# asctime = datetime.datetime.fromtimestamp(record.created).strftime('%Y-%m-%d %H:%M:%S')

url = 'https://confluence.mts.ru/XXXXX/XXXXXXXXXXXXXXXXXXXXXXXXX'


def load_chan_plan(url:str) -> None:
    """Function loads channel plan *.xlsx file from confluence via Selenium"""
    print(url)
    options = Options() 
    # options.add_argument('--ignore-certificate-errors-spki-list') 
    # options.add_argument("--headless")  # Запуск браузера в фоновом режиме без интерфейса
    # включаем удаленную отладку на указанном порту 
    options.add_argument("--remote-debugging-port=60513") 

    # Создание экземпляра драйвера Chrome с указанными опциями
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    try:
        # Открытие указанного URL в браузере
        driver.get(url)
        time.sleep(10)
   
        next_button = driver.find_element(By.XPATH, '/html/body/div/div[2]/div/div/form[1]/div[2]/button') 
        next_button.click() 
        time.sleep(10)

        driver.get('https://confluence.mts.ru/XXXXX/XXXXXXXXXXXXXXXXXXXXXXXXX')
        time.sleep(2)

        next_button = driver.find_element(By.XPATH, '//*[@id="main-content"]/div[1]/table/tbody/tr[3]/td[5]/div/p/a') 
        next_button.click() 
        time.sleep(2)

        next_button = driver.find_element(By.XPATH, '/html/body/div[14]/div[2]/div[1]/div[1]/div/div/a') 
        time.sleep(2)
        next_button.click() 
        time.sleep(20)
        # Закрытие браузера и завершение работы Selenium WebDriver
        driver.quit()
        # return True
    except Exception as e:
        print(f"Произошла ошибка: {type(e).__name__} - {str(e)}")
        logging.error({type(e).__name__} - {str(e)}, exc_info=True)
        # return False
    
mask = 'ЦТВ_канальный_и_пакетный_план_*.xlsm'
default_load_folder = Path(r'C:\Users\yygolovy\Downloads')

def move_conf_file(mask:str, default_load_folder:Path) -> None:
    """ Create work folder and move there channel plan *.xlsx file"""
    # Получаем текущую дату и время
    current_datetime = datetime.now()
    # Форматируем текущую дату и время в нужный формат
    folder_name = f"Chan_plan_{current_datetime.strftime('%d%m%Y%H%M')}"
    try:
        # Создаем папку
        new_folder = Path.cwd() / folder_name
        new_folder.mkdir()

        files = list(default_load_folder.glob(mask))

        if files:
            # Находим последний файл по дате изменения
            latest_file = max(files, key=lambda f: f.stat().st_mtime)
            # Перемещаем найденный файл в целевую папку
            new_location = new_folder / latest_file.name
            latest_file.rename(new_location)
            print(f"Файл {latest_file.name} перемещен в {folder_name}")
            logging.info(f"Файл {latest_file.name} перемещен в {folder_name}")
            return new_location
        else:
            print("Файлов по заданной маске не найдено")
            logging.warning(f"Файлов по заданной маске {mask} не найдено")
    except Exception as e:
        print(f"Произошла ошибка: {type(e).__name__} - {str(e)}")
        logging.error({type(e).__name__} - {str(e)}, exc_info=True)

if __name__ == '__main__':
    print('проверка')   
    load_chan_plan(url)
    move_conf_file(mask, default_load_folder)