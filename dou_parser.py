import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException
import xlsxwriter
import time
import undetected_chromedriver


user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) ' \
             'Chrome/107.0.0.0 Safari/537.36'
headers = {'user-agent': f'{user_agent}'}
options = webdriver.ChromeOptions()
options.add_argument(f'user-agent={user_agent}')
options.add_argument('--disable-blink-features=AutomationControlled')
# options.add_argument(r"--user-data-dir=C:\Users\accof\AppData\Local\Google\Chrome\User Data\Default")
executable_path = r'chromedriver_win32\chromedriver.exe'


# decoder for emails
def decode(g):
    r = int(g[:2], 16)
    email = ''.join([chr(int(g[i:i + 2], 16) ^ r) for i in range(2, len(g), 2)])
    return email


# extraction data to viewable look
def data_extract(value):
    return '\n'.join([
        i for i in list(map(lambda new_value: new_value.strip(), value.findAll(text=True, recursive=False))) if i])


# run webdriver, getting a list of websites, and scroll the site when the list ends
def get_url():
    url = "https://jobs.dou.ua/companies/"
    key = 'Львів'
    driver = undetected_chromedriver.Chrome(options=options)
    driver.maximize_window()
    driver.get(url)
    time.sleep(10)
    driver.find_element(By.CSS_SELECTOR, 'input.company').send_keys(key)
    driver.find_element(By.CSS_SELECTOR, 'input.btn-search').click()
    load_more_button = driver.find_element(By.CSS_SELECTOR, 'div.more-btn a')
    j = 0
    try:
        while load_more_button.is_displayed():
            load_more_button.click()
            time.sleep(2)
            companies = driver.find_elements(By.CSS_SELECTOR, 'div.company')
            for company in companies[j:]:
                yield company.find_element(By.CLASS_NAME, 'cn-a').get_attribute('href')
            j = len(companies)
    except StaleElementReferenceException as ex:
        print(ex)
    finally:
        driver.close()
        driver.quit()


def parser():
    for url in get_url():
        offices = url + 'offices/'
        if not (response := requests.get(url=offices, headers=headers)):
            print(response.status_code)
            continue
        print(response.status_code)
        soup = BeautifulSoup(response.text, 'html.parser')
        contacts = soup.select('#lvov + div div.contacts div.contacts')
        company_name = soup.find('h1', class_='g-h2').text.strip()
        if site := soup.find('div', class_='site'):
            company_site = site.find('a').get('href')
        else:
            company_site = url
        vacancies = url + 'vacancies/'
        company_vacancies = vacancies if soup.find('ul', class_='company-nav').find_all(text='Вакансії') else ''

        for contact in contacts:
            company_address, company_address_map, company_phones, company_mail = [''] * 4
            if value := contact.find('div', class_='address'):
                company_address = data_extract(value)
                company_address_map = contact.find('a').get('href')
            if value := contact.find('div', class_='phones'):
                company_phones = data_extract(value)
            if value := contact.find('div', class_='mail'):
                company_mail = decode(value.find('span').get('data-cfemail'))
            yield company_name, company_address, company_phones, company_vacancies, company_site, company_mail, company_address_map


def sheet(parameter):
    book = xlsxwriter.Workbook('IT_companies_lviv.xlsx')
    page = book.add_worksheet('IT_companies_lviv')
    row, column = 0, 0
    page.set_column('A:A', 15)
    page.set_column('B:B', 25)
    page.set_column('C:C', 15)
    page.set_column('D:D', 20)
    for item_row in parameter():
        print(item_row)
        for col_num, item in enumerate(item_row):
            page.write(row, col_num, item)
        row += 1
    book.close()


if __name__ == "__main__":
    sheet(parser)
