import os
import time
import csv
import datetime
import re

import xlrd
from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys

PATH = f"{os.getcwd()}\\drivers\\chromedriver.exe"
options = webdriver.chrome.options.Options()
driver = webdriver.Chrome(executable_path=PATH)

# Список дерева заданных путей к категориям
book = xlrd.open_workbook("grocery_browse.xls")
list_of_tree = []
excel_worksheet = book.sheet_by_index(1)
for row in range(1, excel_worksheet.nrows):
    list_of_tree.append(int(excel_worksheet.cell_value(row, 0)))
print(list_of_tree)


def scrapeElementsFromUl(div_with_uls):
    '''Собираем сслыки на всех продавцов в списке'''
    list_of_links_to_shop = []
    try:
        list_of_uls = WebDriverWait(div_with_uls, 5).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, "ul"))
            )
        print("Ишем таблицы с продавцами")
    except TimeoutException as e:
        print(e)
        driver.quit()
    for ul_element in list_of_uls:
        a_elements = ul_element.find_elements_by_tag_name("a")
        for a in a_elements:
            if a.get_attribute("title") == "Amazon.com":
                continue
            list_of_links_to_shop.append(a)
    print("Добавили каждого продавца")
    return list_of_links_to_shop


def get_seller_id(seller_link):
    '''Находим id продавцы в html'''
    ActionChains(driver).key_down(Keys.CONTROL).click(seller_link).key_up(Keys.CONTROL).perform()
    driver.switch_to.window(driver.window_handles[1])
    print("Открыл продавца")
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((
                By.XPATH, "//*[@id='search']/div[1]/div[2]/div/span[3]/div[2]/div[1]/div/span/div/div/span/a/div/img"))
                ).click()
    except TimeoutException:
        print("Can't find seller`s id")
        driver.quit()
    text = driver.find_element_by_id("merchantID").get_attribute("value")
    return text


def getInfoAboutProducts():
    '''Выводим данные о всех продуктах продавца и слохраняем ее в csv файл'''

    product_text = []
    pages = int(driver.find_element_by_xpath(
        '//*[@id="search"]/div[1]/div[2]/div/span[3]/div[2]/div[17]/span/div/div/ul/li[6]'
        ).text)
    for _ in range(pages - 45):
        # time.sleep нужен чтобы успели прогрузиться все продуктов на странице
        # Так как страница динамическая, сложно отслеживать нахождение всех элементов при помощи WebDriverWait
        # 1.5 секунд достаточно для загрузки
        time.sleep(1.5)
        product_block = driver.find_elements_by_xpath("//div[@data-component-type='s-search-result']")

        for product_name in product_block:
            name = product_name.find_element_by_css_selector(".a-size-medium.a-color-base.a-text-normal")
            try:
                # Количество элементов в наборе
                num_of_elems = re.findall("\d+", name.text.split("(")[1])[0]
                print(num_of_elems)
            except:
                num_of_elems = 1
                print(num_of_elems)
            try:
                price = product_name.find_element_by_css_selector(".a-price")
                price = ".".join(price.text.split("$")[1].split("\n"))
            except NoSuchElementException:
                price = 'No price'
            count_reviews = product_name.find_element_by_css_selector(".a-size-base")

            # Делаем проверку на начало ревью с доллара, так как если нет ревью то соханяется цена за штуку товара
            if count_reviews.text.startswith("$"):
                count_reviews = "No reviews"
            else:
                count_reviews = "".join(count_reviews.text.split(','))

            try:
                rating = product_name.find_element_by_xpath(".//div[@class='a-section a-spacing-none a-spacing-top-micro']/div/span[1]")
                rating = rating.get_attribute("aria-label").split(' ')[0]
            except NoSuchElementException:
                rating = "No rating"
            # rating.split(" ")[0] нужен для того чтобы разделить строк [<rate>] out of 5 нам нужен только rate
            product_text.append([name.text, price, count_reviews, rating, num_of_elems])
            print("Добавляю текст")
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((
                    By.CLASS_NAME, 'a-last'))
                ).click()
            print("Нажал next")
        except TimeoutException:
            print("Can't find the Next button")
            driver.quit()

    return product_text


def writeZipCode():
    '''Ставим зип код Сша чтобы отображались все продукты Поставищка'''
    driver.get("https://www.amazon.com/ref=nav_logo")
    driver.find_element_by_xpath('//*[@id="glow-ingress-line2"]').click()
    try:
        zip_code = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="GLUXZipUpdateInput"]'))
            )
    except TimeoutException as e:
        print(e)
        driver.quit()
    zip_code.send_keys('85001')
    driver.find_element_by_xpath('//*[@id="GLUXZipUpdate"]/span/input').click()
    driver.find_element_by_xpath('//*[@id="a-popover-3"]/div/div[2]/span').click()


def openTree(tree_id):
    '''Открываем Топ продавцов в нужной нам категории товаров'''
    driver.get(f"https://www.amazon.com/gp/search/other/?pickerToList=enc-merchantbin&node={tree_id}")
    try:
        div_with_uls = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "refinementList"))
            )
    except TimeoutException:
        print("Не удалось найти продавца")
    return div_with_uls


def getSellerInfo(tree_id, seller_id):
    '''
    Здесь переходим к товарам продавца
    Здесь же нужно будет собрать всю инфу о продавце и занести в csv файл
    '''
    driver.get(f"https://www.amazon.com/sp?_encoding=UTF8&seller={seller_id}")
    name = driver.find_element_by_xpath("//*[@id='sellerName']").text
    merchantID = seller_id
    count_reviews = driver.find_element_by_xpath('//*[@id="feedback-summary-table"]/tbody/tr[5]/td[5]/span').text
    rating = driver.find_element_by_xpath('//*[@id="seller-feedback-summary"]/i[1]/span').get_attribute("innerHTML").split(' ')[0]
    address = ' '.join(driver.find_element_by_xpath('//*[@id="seller-profile-container"]/div[2]/div/ul/li[2]/span/ul').text.split('\n'))

    driver.find_element_by_xpath('//*[@id="products-link"]/a').click()
    seller_info = [name, merchantID, rating, count_reviews, address, 
                   tree_id, driver.current_url,
                   f"https://www.amazon.com/sp?_encoding=UTF8&seller={merchantID}"]
    with open(f"saved_csv_merchant\\{tree_id}_merchant", 'a', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(seller_info)

    print(seller_info)
    return seller_info


def run_amazon_scraper(list_of_tree):
    writeZipCode()
    for tree in list_of_tree:
        div_with_uls = openTree(int(tree))
        list_of_links = scrapeElementsFromUl(div_with_uls)
        for link in list_of_links:
            seller_id = get_seller_id(link)
            seller_info = getSellerInfo(int(tree), seller_id)
            result = getInfoAboutProducts()

            driver.close()
            driver.switch_to.window(driver.window_handles[0])

            with open(f'saved_csvs\\{tree}.csv', 'a', encoding='utf-8') as csvfile:
                csvwriter = csv.writer(csvfile)
                csvwriter.writerow(["product`s name", "price", "count_reviews", "rating"])
                for product_info in result:
                    csvwriter.writerow((datetime.date.today(), seller_info[1], seller_info[0], seller_info[-1],
                                        seller_info[-2], product_info[4], seller_info[4],
                                        product_info[0], product_info[1],
                                        product_info[2], product_info[3]))


run_amazon_scraper(list_of_tree)
