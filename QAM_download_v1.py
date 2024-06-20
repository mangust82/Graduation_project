"""the module downloads the configuration file to the QAM Cricket+ analyzer using the 
selenium module which simulates user action"""
from selenium import webdriver 
from selenium.webdriver.common.by import By 
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.chrome.options import Options 
import time
from pathlib import Path
import logging


g2x_username = '******'
g2x_password = '*****'
IP_address = '10.XX.XX.XX'
conf_name = str(Path("C:\work\MVP\Scripts\dowload_conf\IQAlias_IQ Cricket-QAM Plus.xls"))

logging.basicConfig(level=logging.INFO, filename="ETL.log", filemode="a",
                    format="%(asctime)s %(levelname)s %(message)s")

def load_conf_QAM(g2x_username:str, g2x_password:str, IP_address:str, conf_name:str) -> None:
    """Function loads configuration *.csv file on the DVBC analuser QAM Cricket+ via Selenium"""
    # Создание объекта опций драйвера Chrome
    prefix = 'http://'
    url = prefix + IP_address
    print(url)
    options = Options() 
    # options.add_argument('--ignore-certificate-errors-spki-list') 
    options.add_argument("--headless")  # Запуск браузера в фоновом режиме без интерфейса
    # включаем удаленную отладку на указанном порту 
    # options.add_argument("--remote-debugging-port=60513") 

    # Создание экземпляра драйвера Chrome с указанными опциями
    driver = webdriver.Chrome(options=options)
    try:
        # Открытие указанного URL в браузере
        driver.get(url) 
        
        # Нахождение поля для ввода имени пользователя
        username = driver.find_element(By.XPATH, '//*[@id="mytab"]/div[2]/form[1]/table/tbody/tr[1]/td[3]/input') 
        username.send_keys(g2x_username) 
        
        # Нахождение поля для ввода пароля
        password = driver.find_element(By.XPATH, '//*[@id="mytab"]/div[2]/form[1]/table/tbody/tr[2]/td[3]/input') 
        password.send_keys(g2x_password)
        
        # Нахождение кнопки для логина и нажатие на нее
        login = driver.find_element(By.XPATH, '//*[@id="login"]/td[2]/button') 
        login.click() 
        time.sleep(2) 
        
        # Нахождение и клик на первом фрейме чтобы исполнился скрипт
        test = driver.find_element(By.XPATH, '/html/frameset/frame[1]') 
        test.click() 
        
        # Переключение на фрейм с идентификатором "contents"
        driver.switch_to.frame("contents") 
        time.sleep(1) 
        
        # Нахождение элемента для загрузки предварительной конфигурации и клик на него
        dowload_pre_conf = driver.find_element(By.XPATH, '/html/body/div[1]/div/a[9]') 
        dowload_pre_conf.click() 
        time.sleep(1) 
        
        # Нахождение элемента для загрузки конфигурации и клик на него
        dowload_conf = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/a[2]') 
        dowload_conf.click() 
        time.sleep(1) 
        
        # Переключение на фрейм с именем main
        driver.switch_to.default_content() 
        driver.switch_to.frame("main") 
        
        # Нахождение поля для выбора файла, ввод пути к файлу и отправка формы
        choose_file = driver.find_element(By.XPATH, '//*[@id="tbl"]/tbody/tr[2]/td/form/input') 
        choose_file.send_keys(conf_name) 
        choose_file.submit() 
        time.sleep(2)  # Пауза на 2 секунды
        
        # Нахождение заголовка h2 и вывод его текста
        log = driver.find_element(By.XPATH, '/html/body/h2') 
        print(log.text)
        logging.info(f'{IP_address} {log.text}')
        # Закрытие браузера и завершение работы Selenium WebDriver
        driver.quit()
    except Exception as e:
        print(f"Произошла ошибка: {type(e).__name__} - {str(e)}")
        logging.error({IP_address} - {type(e).__name__} - {str(e)}, exc_info=True)

if __name__ == '__main__':    
    load_conf_QAM(g2x_username, g2x_password, IP_address, conf_name)