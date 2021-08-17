import os
import time
from datetime import datetime
from dotenv import load_dotenv
import yagmail
import xlsxwriter
from lxml import etree
import urllib.request
from selenium import webdriver
from selenium.webdriver.support.ui import Select


def get_rate_xml_url():
    geckodriver = os.path.abspath('geckodriver')
    browser = webdriver.Firefox(executable_path=geckodriver)
    browser.maximize_window()
    browser.delete_all_cookies()
    url = 'https://www.moex.com/'
    browser.get(url)

    menu_xpath = '/html/body/div[3]/div[2]/div/div/div/div[2]\
        /nav/span[1]/a'
    browser.find_element_by_xpath(menu_xpath).click()
    time.sleep(1)

    futures_market_xpath = '/html/body/div[3]/div[2]/div/div/div\
        /div[2]/nav/span[1]/div/div/div/div[1]/div[3]/a'
    browser.find_element_by_xpath(futures_market_xpath).click()
    time.sleep(2)

    i_agree_xpath = '//*[@id="content_disclaimer"]/div/div/div\
        /div[1]/div/a[1]'
    browser.find_element_by_xpath(i_agree_xpath).click()
    time.sleep(1)

    indicative_courses_xpath = '//*[@id="ctl00_frmLeftMenuWrap"]\
        /div/div/div/div[2]/div/a[13]'
    browser.find_element_by_xpath(indicative_courses_xpath).click()
    time.sleep(1)

    start_day = browser.find_element_by_xpath('//*[@id="d1day"]')
    start_day_dd = Select(start_day)
    start_day_dd.select_by_value('1')
    
    start_month = browser.find_element_by_xpath('//*[@id="d1month"]')
    start_month_dd = Select(start_month)
    current_month = datetime.now().month
    start_month_dd.select_by_value(str(current_month))
    
    browser.find_element_by_name('bSubmit').click()
    
    get_xml_xpath_link = '/html/body/div[3]/div[3]/div/div/div[1]\
        /div[2]/div/div/div/form/div[5]/div[2]/div/a'
    xml_url = {}
    xml_url['usd'] = browser.find_element_by_xpath(get_xml_xpath_link).get_attribute('href')

    currency = browser.find_element_by_xpath('//*[@id="ctl00_PageContent_CurrencySelect"]')
    currency_dd = Select(currency)
    currency_dd.select_by_value('EUR_RUB')

    start_day = browser.find_element_by_xpath('//*[@id="d1day"]')
    start_day_dd = Select(start_day)
    start_day_dd.select_by_value('1')
    
    start_month = browser.find_element_by_xpath('//*[@id="d1month"]')
    start_month_dd = Select(start_month)
    current_month = datetime.now().month
    start_month_dd.select_by_value(str(current_month))
    
    browser.find_element_by_name('bSubmit').click()

    get_xml_xpath_link = '/html/body/div[3]/div[3]/div/div/div[1]\
        /div[2]/div/div/div/form/div[5]/div[2]/div/a'
    xml_url['eur'] = browser.find_element_by_xpath(get_xml_xpath_link).get_attribute('href')
    browser.quit()

    return xml_url


def read_xml(url):
    web_file = urllib.request.urlopen(url)
    data = web_file.read()
    return data


def parse(xml_data):
    exchange_rates = []
    currency_row = ()
    root = etree.fromstring(xml_data)
    for appt in root.getchildren():
        date = [
            elem[1].get('moment').split()[0]
            for elem in enumerate(appt.getchildren()) 
            if elem[0] % 2 == 1
        ]
        rate = [
            elem[1].get('value')
            for elem in enumerate(appt.getchildren()) 
            if elem[0] % 2 == 1
        ]
        rate_tmp = [
            elem[1].get('value')
            for elem in enumerate(appt.getchildren()) 
            if elem[0] % 2 == 0
        ]
        change = [
            str(float(r)-float(rt)) 
            for r, rt in zip(rate, rate_tmp)
        ]
    exchange_rates = [date, rate, change]
    return exchange_rates


def dividing_eur_by_usd(eur, usd):
    divide = [str(float(e)/float(u)) for e, u in zip(eur[1], usd[1])]
    exchange_rates = usd + eur
    exchange_rates.append(divide)
    return exchange_rates


def create_xlsx_report(exchange_rates_usd_eur):
    current_date = datetime.now().strftime("%m-%Y")
    workbook = xlsxwriter.Workbook(f'exchange_rates_{current_date}.xlsx')
    worksheet = workbook.add_worksheet('exchange rates')
    worksheet.write('A1', 'Дата')
    worksheet.write('B1', 'Курс')
    worksheet.write('C1', 'Изменение')
    worksheet.write('D1', 'Дата')
    worksheet.write('E1', 'Курс')
    worksheet.write('F1', 'Изменение')
    worksheet.write('G1', 'Частное eur и usd')
    head_width_column_set = (4, 4, 10, 4, 4, 10, 12)
    currency_symbol = {0:'', 1:'$', 2:'$', 3:'', 4:'€', 5:'€', 6:''}
    column_number = 0
    for column, column_width in zip(exchange_rates_usd_eur, head_width_column_set):
        row_number = 1
        for value in column:
            if len(value) > column_width:
                column_width = len(value)+1
            if column_number == 0 or column_number == 3:
                worksheet.write(row_number, column_number, value)
            else:
                num_after_point = len(value.split('.')[1])
                symbol = currency_symbol[column_number]
                finance_format = workbook.add_format({
                    'num_format': f'[${symbol}]#,{"#"*num_after_point}0.{"0"*num_after_point}', 
                    'align': 'left'
                })
                worksheet.write(row_number, column_number, float(value), finance_format)
            row_number+=1
        worksheet.set_column(column_number, column_number, column_width)
        column_number+=1
    workbook.close()


def decline(number: int):
    row_declines = ['строка', 'строки', 'строк']
    cases = [ 2, 0, 1, 1, 1, 2 ]
    if 4 < number % 100 < 20:
        idx = 2
    elif number % 10 < 5:
        idx = cases[number % 10]
    else:
        idx = cases[5]
    return row_declines[idx]


def send_email(number_rows):
    load_dotenv('.env')
    current_date = datetime.now().strftime("%m-%Y")
    decline_row = decline(number_rows)
    smtp_server = os.getenv('SMTP_SERVER')
    smtp_password = os.getenv('SMTP_PASSWORD')
    receiver = os.getenv('RECEIVER')
    body = f"В файле без учёта названия столбцов: {number_rows} {decline_row}."
    subject='Тестовое задание "ГРИНАТОМ"'
    filename = f'exchange_rates_{current_date}.xlsx'
    yag = yagmail.SMTP(smtp_server, smtp_password)
    yag.send(
        to=receiver,
        subject=subject,
        contents=body,
        attachments=filename,
    )


def main():
    xml_url = get_rate_xml_url()
    usd_data = read_xml(xml_url['usd'])
    eur_data = read_xml(xml_url['eur'])
    usd = parse(usd_data)
    eur = parse(eur_data)
    exchange_rates_usd_eur = dividing_eur_by_usd(eur, usd)
    create_xlsx_report(exchange_rates_usd_eur)
    number_rows = len(exchange_rates_usd_eur[0])
    send_email(number_rows)


if __name__=='__main__':
    main()
