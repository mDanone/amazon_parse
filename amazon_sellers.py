import os

import xlrd
from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains

PATH = f"{os.getcwd()}\\drivers\\chromedriver.exe"
options = webdriver.chrome.options.Options()
driver = webdriver.Chrome(executable_path=PATH)

book = xlrd.open_workbook("grocery_browse.xls")
list_of_tree = []
excel_worksheet = book.sheet_by_index(1)

for row in range(1, excel_worksheet.nrows):
    list_of_tree.append(int(excel_worksheet.cell_value(row, 0)))


def scrapeElementsFromUl(div_with_uls):
    '''
    Getting all links of shops from top sellers
    '''
    list_of_links_to_shop = []
    list_of_uls = div_with_uls.find_elements_by_tag_name("ul")
    for ul_element in list_of_uls:
        a_elements = ul_element.find_elements_by_tag_name("a")
        for a in a_elements:
            if a.get_attribute("title") == "Amazon.com":
                continue
            list_of_links_to_shop.append(a)
    return list_of_links_to_shop


def get_seller_id(seller_link):
    ActionChains(driver).click(seller_link).perform()
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((
            By.XPATH, "//*[@id='search']/div[1]/div[2]/div/span[3]/div[2]/div[1]/div/span/div/div/span/a/div/img"))
            ).click()
    text = driver.find_element_by_id("merchantID").get_attribute("value")
    return text


def getInfoAboutProducts():
    pages = int(driver.find_element_by_xpath('//*[@id="search"]/div[1]/div[2]/div/span[3]/div[2]/div[18]/span/div/div/ul/li[6]').text)
    content_block = driver.find_elements(By.XPATH, "//span[contains(@class,'a-link-normal a-text-normal')]")
    return content_block


def writeZipCode():
    '''
    Need to get working Scraper when you live in other countries, nt USA
    '''
    driver.find_element_by_xpath('//*[@id="glow-ingress-line2"]').click()
    zip_code = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="GLUXZipUpdateInput"]')))
    zip_code.send_keys('85001')
    driver.find_element_by_xpath('//*[@id="GLUXZipUpdate"]/span/input').click()
    driver.find_element_by_xpath('//*[@id="a-popover-3"]/div/div[2]/span').click()


def openTree(tree_id):
    '''
    Here we are opening link with top sellers from the category id we got
    '''
    driver.get(f"https://www.amazon.com/gp/search/other/?pickerToList=enc-merchantbin&node={tree_id}")
    writeZipCode()
    div_with_uls = driver.find_element_by_id("ref_275225011")
    return div_with_uls


def getSellerInfo(seller_id):
    driver.get(f"https://www.amazon.com/sp?_encoding=UTF8&seller={seller_id}")
    driver.find_element_by_xpath('//*[@id="products-link"]/a').click()


if __name__ == "__main__":
    div_with_uls = openTree(16323201)
    list_of_links = scrapeElementsFromUl(div_with_uls)
    seller_id = get_seller_id(list_of_links[0])
    getSellerInfo(seller_id)
    print(getInfoAboutProducts())







# //*[@id="ref_275225011"]/ul[1]/li[1]/span/a ref_275225011 ref_275225011
# //*[@id="ref_275225011"]/ul[1]/li[2]/span/a
# first link
# https://www.amazon.com/gp/search/other/?pickerToList=enc-merchantbin&node={tree_id}

# second link
# https://www.amazon.com/sp?_encoding=UTF8&seller={seller_id}

# third link
