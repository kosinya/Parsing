from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import csv

from parsing_funcs import fill_excel_template, click_show_more_button, get_new_hotels_count, extract_hotel_data

base_dir = os.path.dirname(os.path.abspath(__file__))
CHROME_DRIVER_PATH = os.path.join(base_dir, 'driver', 'chrome', 'chromedriver.exe')


service = Service(CHROME_DRIVER_PATH)
driver = webdriver.Chrome(service=service)
actions = ActionChains(driver)

# ССЫЛКА НА СТРАНИЦУ С ДАННЫМИ С УЖЕ ПРИМЕНЕННЫМИ ФИЛЬТРАМИ
url = "https://tourism.fsa.gov.ru/ru/resorts/showcase/hotels?regionIdList=2&categoryIdList=1&statusIdList=6"


if __name__ == '__main__':
    try:
        driver.get(url)
        time.sleep(10)

        page_number = 1
        hotel_number = 0
        csv_data = []

        while True:
            all_hotels = driver.find_elements(By.TAG_NAME, "hotels-resort-card")
            current_count = len(all_hotels)

            while hotel_number in range(len(all_hotels)):
                print(f"\nОтель {hotel_number + 1}/{len(all_hotels)}")

                for i in range(page_number):
                    click_show_more_button(driver)
                    time.sleep(1)

                hotels = driver.find_elements(By.TAG_NAME, "hotels-resort-card")
                hotel = hotels[hotel_number]

                driver.execute_script("arguments[0].click();", hotel)
                print(f"Открыли: {driver.current_url}")
                time.sleep(5)

                hotel_data = extract_hotel_data(driver)
                hotel_data['number'] = hotel_number + 1

                hotel_csv_data = {
                    'Название гостиницы': hotel_data.get('name', ''),
                    'Адрес отеля': hotel_data.get('address', ''),
                    'Тип средства размещения': hotel_data.get('type', ''),
                    'Звёздность': f"{hotel_data.get('stars', 0)} звезд" if hotel_data.get('stars',
                                                                                          0) > 0 else "нет категории",
                    'Статус': 'действует',
                    'Email': hotel_data.get('hotel_email', ''),
                    'Телефон отеля': hotel_data.get('hotel_phone', ''),
                    'Телефон владельца': hotel_data.get('owner_phone', ''),
                    'ИНН': hotel_data.get('inn', ''),
                    'Адрес владельца': hotel_data.get('owner_address', ''),
                    'Номерной фонд': hotel_data.get('num', '0'),
                    'ФИО руководителя': hotel_data.get('owner', '')
                }
                print(hotel_csv_data)
                csv_data.append(hotel_csv_data)

                hotel_number += 1
                driver.back()
                time.sleep(5)

            if click_show_more_button(driver):
                page_number += 1
            else:
                break

            time.sleep(5)

            if not get_new_hotels_count(driver, current_count):
                break

        print("\nСбор окончен.")

        if csv_data:
            csv_filename = 'отели_данные.csv'
            fieldnames = [
                'Название гостиницы',
                'Адрес отеля',
                'Тип средства размещения',
                'Звёздность',
                'Статус',
                'Email',
                'Телефон отеля',
                'Телефон владельца',
                'ИНН',
                'Адрес владельца',
                'Номерной фонд',
                'ФИО руководителя'
            ]

            with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as file:
                writer = csv.DictWriter(file, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(csv_data)

            print(f"\nДанные сохранены в файл: {csv_filename}")

            excel_template_filename = '1_1_Форма_для_выходных_данных_с_сайта_по_средствам_размещения.xlsx'
            output_filename = 'Реестр_средств_размещения_заполненный.xlsx'

            if os.path.exists(excel_template_filename):
                print(f"Найден шаблон Excel: {excel_template_filename}")
                fill_excel_template(csv_filename, excel_template_filename, output_filename)
            else:
                print(f"Шаблон Excel '{excel_template_filename}' не найден!")


    finally:
        driver.quit()