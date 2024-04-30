from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException 
import keyboard as kb

import requests as rq
from bs4 import BeautifulSoup
import openpyxl

def add_data_in_column(worksheet, data_array, column_num):
    for row_num, value in enumerate(data_array, start=1):
        worksheet.cell(row=row_num, column=column_num, value=value)
    
def string_cleaner(s):
    while("\r" in s or "\n" in s or "\t" in s):
        s = s.replace("\r", "", 1)
        s = s.replace("\n", "", 1)
        s = s.replace("\t", "", 1)
    return s
def number_of_reviews_cleaner(number_of_reviews):
    number_of_reviews = string_cleaner(number_of_reviews)
    number_of_reviews = number_of_reviews.replace("(", "", 1)
    number_of_reviews = number_of_reviews.replace(")", "", 1)
    return number_of_reviews

def check_exists_by_xpath(xpath, driver):
    try:
        driver.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True

def find_game_title_elements(driver, elements, start):
    game_title_elements = []
    current_elements = elements
    try:
        while True:
            game_title_element = driver.find_element(By.XPATH, f'//*[@id="SaleSection_13268"]/div[2]/div[2]/div[2]/div[2]/div/div[{start}]/div/div/div/div[2]/div[2]/a/div')
            game_title_elements.append(game_title_element.text)
            start+=1
    except NoSuchElementException:
        if (game_title_elements):
            for el in game_title_elements:
                current_elements.append(el)
                
            return (current_elements, start)
        return (False, start)
    
def get_game_element(driver, xpath):
    try:
        game_element = driver.find_element(By.XPATH, xpath)
        return game_element.text
    except NoSuchElementException:
        return ""

game_title_elements = ["Название"]
game_date_elements = ["Дата выхода"]
game_rating_elements = ["Отзывы"]
game_numbers_of_reviews = ["Кол-во отзывов"]
game_price_elements = ["Цена"]
game_picture_elements = ["Картинка"]
game_link_elements = ["Ссылка на игру"]
game_description_elements = ["Описание"]




def get_source_link(driver, xpath):
    try:
        game_link = driver.find_element(By.XPATH, xpath)
        return game_link.get_attribute('href')
    except NoSuchElementException:
        return ""
def get_picture_link(driver, xpath):
    try:
        picture_link = driver.find_element(By.XPATH, xpath)
        return picture_link.get_attribute('src')
    except NoSuchElementException:
        return ""
def get_game_discount(driver, xpath_discount, xpath):
    
    try:
        price_with_discount = driver.find_element(By.XPATH, xpath_discount)
        return price_with_discount.text
    except NoSuchElementException:
        try:
            price = driver.find_element(By.XPATH, xpath)
            return price.text
        except NoSuchElementException:
            return ""
def get_game_prices(url):
    try:
        page = rq.get(url)
        soup = BeautifulSoup(page.text, "lxml")
        prices = soup.find('div', class_='game_purchase_price price').text
        prices = string_cleaner(prices)
        return prices
    except Exception:
        return ""
    

def get_game_discount_prices(url):
    try:
        page = rq.get(url)
        soup = BeautifulSoup(page.text, "lxml")
        prices = soup.find('div', class_='discount_final_price').text
        prices = string_cleaner(prices)
        return prices
    except Exception:
        return ""

def get_game_reviews(url):
    try:
        page = rq.get(url)
        soup = BeautifulSoup(page.text, "lxml")
        reviews = soup.find("span", class_="game_review_summary positive").text
        return reviews
    except Exception:
        return ""
def get_game_number_of_reviews(url):
    try:
        page = rq.get(url)
        soup = BeautifulSoup(page.text, "lxml")
        number_of_reviews = soup.find("span", class_="responsive_hidden").text
        number_of_reviews = number_of_reviews_cleaner(number_of_reviews)
        return number_of_reviews
    except:
        return ""
def get_game_discription(url):
    try:
        page = rq.get(url)
        soup = BeautifulSoup(page.text, "lxml")
        discription = soup.find('div', class_='game_description_snippet').text

        discription = string_cleaner(discription)
        return discription
    except:
        return ""
def parse(driver, start):
    
    
    try:
        while True:
            game_title_element = driver.find_element(By.XPATH, f'//*[@id="SaleSection_13268"]/div[2]/div[2]/div[2]/div[2]/div/div[{start}]/div/div/div/div[2]/div[2]/a/div')
            game_title_elements.append(game_title_element.text)

            game_date_elements.append(get_game_element(driver, f'//*[@id="SaleSection_13268"]/div[2]/div[2]/div[2]/div[2]/div/div[{start}]/div/div/div/div[2]/div[3]/div[2]/div'))
            
            link = get_source_link(driver, f'//*[@id="SaleSection_13268"]/div[2]/div[2]/div[2]/div[2]/div/div[{start}]/div/div/div/div[2]/div[2]/a')
            game_link_elements.append(link)
            game_picture_elements.append(get_picture_link(driver, f'//*[@id="SaleSection_13268"]/div[2]/div[2]/div[2]/div[2]/div/div[{start}]/div/div/div/div[1]/a/div/div[2]/img'))
            game_description_elements.append(get_game_discription(link))
            price = get_game_discount(driver, f'//*[@id="SaleSection_13268"]/div[2]/div[2]/div[2]/div[2]/div/div[{start}]/div/div/div/div[2]/div[5]/div/div[2]/div[2]/div[2]', f'//*[@id="SaleSection_13268"]/div[2]/div[2]/div[2]/div[2]/div/div[{start}]/div/div/div/div[2]/div[5]/div/div[2]/div')
            game_price_elements.append(price)
            
            game_rating_elements.append(get_game_reviews(link))
            game_numbers_of_reviews.append(get_game_number_of_reviews(link))
            start+=1
    except NoSuchElementException:
        return start

def main():
    url = "https://store.steampowered.com/tags/ru/%D0%98%D0%BD%D0%B4%D0%B8/"
    o = Options()
    o.add_experimental_option("detach", True)
    driver = webdriver.Chrome(options=o)
    driver.get(url)
    
    num = 13
    count = 0
    iterator = 0
    b = True
    game_titles = []
    start = 1
    deep_search = 20
    while(b):

        for i in range(iterator, driver.execute_script("return document.body.scrollHeight") , 20):
            
            if(kb.is_pressed("esc")):
                b = False
                break
            

            start = parse(driver, start)

            driver.execute_script(f"window.scrollTo(0, {i})")
            if(check_exists_by_xpath(f'//*[@id="SaleSection_13268"]/div[2]/div[2]/div[2]/div[2]/div/div[{num}]/button', driver)):

                button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="SaleSection_13268"]/div[2]/div[2]/div[2]/div[2]/div/div[{num}]/button')))
                button.click()
                num+=12
                count+=1
            if(count == deep_search):
                b = False
                break
        iterator = i
    

    wb = openpyxl.Workbook()
    ws = wb.active
    add_data_in_column(ws, game_title_elements, 1)
    add_data_in_column(ws, game_description_elements, 2)
    add_data_in_column(ws, game_picture_elements, 3)
    add_data_in_column(ws, game_date_elements, 4)
    add_data_in_column(ws, game_rating_elements, 5)
    add_data_in_column(ws, game_numbers_of_reviews, 6)
    add_data_in_column(ws, game_link_elements, 7)
    add_data_in_column(ws, game_price_elements, 8)
    
    wb.save("parsing_result2.xlsx")
if __name__ == "__main__":
    main()
