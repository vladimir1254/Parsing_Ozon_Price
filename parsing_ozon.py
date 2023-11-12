from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
import json
import pandas as pd
import csv
import telebot
from datetime import datetime
import openpyxl

bot_token = # Your bot token
bot = telebot.TeleBot(bot_token)


csv_file_path = r'ozon.xlsx'

last_cena  = {}
def get_product_page_html_with_selenium(product_codes,message):
    
    chrome_options = Options()
    #chrome_options.add_argument("--headless")

    
    chrome_options.add_argument("--start-maximized")
    #chrome_options.add_argument("--disable-infobars")
    #chrome_options.add_argument('--disable-web-security')
    # Укажите абсолютный путь к chromedriver.exe на вашем компьютере
    
    chrome_driver_path = "E:\\projects\\BUS_BUS_BUS\\chromeDriver\\chromedriver.exe"

    # Инициализируем объект Service
    service = Service(executable_path=chrome_driver_path)
    data = []
    for product_code in product_codes:
        row = []
        driver = webdriver.Chrome(service=service, options=chrome_options)
      #  driver.minimize_window()

        
        wait = WebDriverWait(driver, 20)
        wait2 = WebDriverWait(driver, 5)
        print(product_code)
        url = f"https://www.ozon.ru/product/{product_code}/"
        try:
            driver.get(url)
         #   elem = WebDriverWait(driver, 20).until(
         #       EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="checkbox"]'))
         #   )
         #   print(elem)
         #   time.sleep(10000)
            name = ''
            try:
                name = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//h1")))            
            except:
                print('Ошибка при получении названия товара')
            for i in name:
                print(i.text)
           # print(name[0].text)
            cena = ''
            try:
                cena = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//span")))
            except Exception as ex:
                print('Ошибка при парсинге цены товара')
            while True:
                cena = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//span")))
                try:
                    for num,i in enumerate(cena):
                        if len(i.text)>0 and i.text[-1]=='₽':
                            cena = int(i.text[:-1].replace(' ',''))
                            break                    
                    break
                except Exception as ex:
                    print(ex)
            color_name = ''
            brand_name = ''
            try:
                color_spans = wait2.until(EC.presence_of_all_elements_located((By.XPATH, f"//span[contains(text(), 'Цвет')]")))
                for i in color_spans:
                    if i.text == 'Цвет':
                        # Найдем предка элемента на два уровня вверх
                        grandparent_element = i.find_element(By.XPATH, '../..')
                        # Вывести текст предка элемента
                        color_name = grandparent_element.text.split('\n')[1]
            except Exception as ex:
                print('not color')
            try:
                color_spans = wait2.until(EC.presence_of_all_elements_located((By.XPATH, f"//span[contains(text(), 'Бренд')]")))
                for i in color_spans:
                    if i.text == 'Бренд':
                        # Найдем предка элемента на два уровня вверх
                        grandparent_element = i.find_element(By.XPATH, '../..')
                        # Вывести текст предка элемента
                        #print('grandpa=',grandparent_element.text)
                        brand_name = grandparent_element.text.split('\n')[1]
            except Exception as ex:
                print('not brand')
            seller = ''
            try:
                element1 = wait2.until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-widget="webCurrentSeller"]')))
                seller = element1.text.split('\n')[1]              
            except Exception as ex:
                    print('not seller')
            # Вывести текст найденного элемента
            name = name[0].text.replace('`','')
            cena_start = ''
            cena_owner = ''
            cena_market = ''
            try:
                color_spans = wait2.until(EC.presence_of_all_elements_located((By.XPATH, f"//span[contains(text(), 'Добавить в корзину')]")))
                for i in color_spans:
                    if i.text == 'Добавить в корзину':
                        # Найдем предка элемента на два уровня вверх
                        grandparent_element = i.find_element(By.XPATH, '../../../../../../../../../../..')
                        # Вывести текст предка элемента
                        lst = grandparent_element.text.replace('\u2009', '').replace('₽','').split('\n')
                        if 'c Ozon Картой' in lst:
                            cena_market = lst[0]
                            cena_owner = lst[2]
                            cena_start = lst[3]
                        else:
                            cena_owner = lst[0]
                            cena_start = lst[1] 
                        print('grandpa=',grandparent_element.text.replace('\u2009', '').replace('₽','').split('\n'))
                        #['747\u2009₽', '1\u2009192\u2009₽', '146\u2009₽', '× 6 месяцев в Ozon Рассрочку', 'Добавить в корзину']
#['14\u2009921\u2009₽', 'c Ozon Картой', '15\u2009543\u2009₽', 'без Ozon Карты', '2\u2009908\u2009₽', '× 6 месяцев в Ozon Рассрочку', 'Добавить в корзину']
                        #cena_name = grandparent_element.text.split('\n')[1]
            except Exception as ex:
                print('error cena',ex)
            # Получаем текущую дату и время
            current_datetime = datetime.now()
            # Форматируем дату и время в заданный формат
            formatted_date = current_datetime.strftime('%Y-%m-%d %H:%M:%S')
            row.append(formatted_date)
            row.append('OZON')
            row.append('-')
            row.append(product_code)
            row.append(name)
            row.append(color_name)
            row.append(brand_name)
            row.append(seller)
            row.append(cena_start)
            row.append(str((int(cena_start)-int(cena_owner))/int(cena_start)*100))
            row.append(cena_owner)
            if cena_market!='':
              row.append(str((int(cena_owner)-int(cena_market))/int(cena_owner)*100))
              row.append(cena_market)
            else:
                row.append('0')
                row.append(cena_owner)
            data.append(row)
            if not product_code in last_cena:
            	last_cena[product_code] = (cena,True,0)
            else:
            	if last_cena[product_code][0]!=cena:
            		last_cena[product_code] = (cena,True,cena-last_cena[product_code][0])
            	else:
            		last_cena[product_code] = (cena,False,0)
        except Exception as e:
            print("Произошла ошибка при выполнении запроса:", e)
       # time.sleep(2000)
        driver.quit()
    print(data)
    workbook = openpyxl.load_workbook(csv_file_path)
    sheet = workbook.active
    for row in data:
	     if last_cena[row[3]][0]:
	      sheet.append(row)
	     #bot.send_message(message.chat.id,"Товар:"+str(row[0])+'\nНовая цена: '+str(row[2])+'\nИзменилась на (новая-старая): '+str(last_cena[row[1]][2])+'\nВ процентах:'+str(round(last_cena[row[1]][2]/(row[2]-last_cena[row[1]][2]),4)*100))
	     else:
	      print('Cena NE IZM propusk'+str(formatted_date))
	# Сохраняем изменения в файле .xlsx
    workbook.save(csv_file_path)
    '''
    with open(csv_file_path, mode='a', newline='', encoding='utf-8') as csv_file:
	    csv_writer = csv.writer(csv_file, delimiter='`')
	    
	    # Записываем новые данные
	    for row in data:
	     if last_cena[row[0]][1]:   
	      csv_writer.writerow(row)
	      bot.send_message(message.chat.id,"Товар:"+str(row[0])+'\nНовая цена: '+str(row[2])+'\nИзменилась на (новая-старая): '+str(last_cena[row[0]][2])+'\nВ процентах:'+str(round(last_cena[row[0]][2]/(row[2]-last_cena[row[0]][2]),4)*100))
	     else:
	      print('Cena NE IZM propusk'+str(formatted_date))

   '''

flag = False

@bot.message_handler(commands=['s'])
def start(message):
    global flag
    flag = not flag
    while flag:
        with open('id_list.txt', 'r') as file:
                # Читаем содержимое файла и сохраняем его в переменную content
            content = file.read()
        products = eval(content)
        print(content)
        get_product_page_html_with_selenium(products,message)
        time.sleep(600) # Каждые 10 минут проверка цены
        print('ok')
@bot.message_handler(commands=['add'])
def add(message):
    if len(message.text.split('/add '))>0:
         st = message.text.split('/add ')[1].replace('  ',' ').replace('\'','').replace('"','')
         my_list = st.split(',')
         with open('1.txt', 'w') as file:
   #     # Записываем список в файл
             file.write(str(my_list))    

bot.polling()
