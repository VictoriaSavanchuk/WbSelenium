import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
import telebot
import schedule

bot = telebot.TeleBot('your token')

excel_file_path = 'sku.xlsx'

driver = webdriver.Edge(executable_path='msedgedriver.exe')
# Список для хранения отправленных отзывов
sent_reviews = []

def load_sent_reviews():
    # Загрузка отправленных отзывов из файла или базы данных
    try:
        with open('sent_reviews.txt', 'r',  encoding='utf-8') as file:
            for line in file:
                sent_reviews.append(line.strip())
    except FileNotFoundError:
        pass

def save_sent_reviews():
    # Сохранение отправленных отзывов в файл или базу данных
    with open('sent_reviews.txt', 'w', encoding='utf-8') as file:
        for review in sent_reviews:
            file.write(review + '\n')

def process_sku():
    workbook = load_workbook(excel_file_path)
    sheet = workbook.active
    sku_column = sheet['A']
    sku_list = [sku.value for sku in sku_column if sku.value is not None]

    for sku in sku_list:
        url = f'https://www.wildberries.ru/catalog/{sku}/detail.aspx'
        driver.get(url)
        driver.maximize_window()   
        # Ожидание загрузки страницы и элементов с информацией о негативном отзыве
        wait = WebDriverWait(driver, 20)
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "product-page__title")))

        product_name = driver.find_element_by_class_name("product-page__title")
        product_name = product_name.text
    
        driver.find_element_by_css_selector('#comments_reviews_link > span.product-review__rating.address-rate-mini.address-rate-mini--sm').click()  
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'rating-product__numb')))
        rating = driver.find_element_by_class_name('rating-product__numb').text   
    
        reviews = driver.find_elements(By.CLASS_NAME,'comments__item')
        
        negative_reviews = []
        for review in reviews:
            stars = review.find_element(By.CLASS_NAME, 'feedback__rating')
            stars = stars.get_attribute('class').split(' ')[2][4]
            if int(stars) < 4:            
                review_text = review.find_element(By.CLASS_NAME,'j-feedback__text').text
                negative_reviews.append({'sku': sku, 'product_name': product_name, 'rating': rating, 'review_text': review_text})  
                #print(negative_reviews)       
                review_id = f'{sku}_{stars}_{review_text}'  # Генерация уникального идентификатора отзыва
                if review_id not in sent_reviews:  # Проверка, отправлен ли отзыв ранее
                    sent_reviews.append(review_id)  # Добавление отзыва в список отправленных
                    message = f"Негативный отзыв\nНазвание товара: {product_name}\nSKU товара: {sku}\nКоличество звезд: {stars}\nТекст отзыва: {review_text}\nТекущий рейтинг товара: {rating}"
                    bot.send_message('CHAT_ID', message)
    
        time.sleep(2)
    driver.quit()
    save_sent_reviews()

# Расписание выполнения задачи
schedule.every().day.at("10:27").do(process_sku)  

load_sent_reviews()  # Загрузка отправленных отзывов перед запуском программы

# Запуск цикла обработки расписания
while True:
    schedule.run_pending()
    time.sleep(1)



# Метод запуска можно реализовать с помощью планировщика задач, например, используя библиотеку schedule.
# Необходимо задать расписание выполнения скрипта (например, каждый день в определенное время)
# и вызвать функцию, которая будет запускать парсинг и отправку уведомлений