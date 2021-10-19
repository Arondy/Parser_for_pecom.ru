# system libraries
import os
from time import sleep
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ex_cond

# Необходимые переменные
url = "https://pecom.ru/services-are/shipping-request/"
directory = os.path.dirname(os.path.realpath(__file__))
env_path = directory + "\chromedriver"
chromedriver_path = directory + "\chromedriver\chromedriver.exe"
mass = '100'
volume = '0.6'
rate = '100'
cargo_type = 'КАНЦЕЛЯРСКИЕ ТОВАРЫ'
read_page1 = pd.read_excel('C:\Python\Python Parsers\Parser_PEK\PEK.xlsx', usecols=[2], index_col='Аппараты ТБ')
read_page2 = pd.read_excel('C:\Python\Python Parsers\Parser_PEK\PEK.xlsx', usecols=[3], index_col='ГОСБ')
column1 = read_page1.index.tolist()
column2 = read_page2.index.tolist()

column1.pop(0)
column2.pop(0)
column1.pop(0)
column2.pop(0)
column1.pop(0)
column2.pop(0)
start = 'Россия, '
array6 = []
array7 = []
array8 = []
array9 = []
array10 = []
array11 = []
array12 = []

# Добавляет chromedriver в PATH
os.environ['PATH'] += env_path


def openChrome():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ['enable-automation'])
    driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
    actions = ActionChains(driver)
    driver.maximize_window()
    driver.get(url=url)
    sleep(10)
    try:
        WebDriverWait(driver, 10).until(ex_cond.presence_of_element_located(
            (By.XPATH, "(//input[@name='CargoTab1TotalWeight'])[1]")))

        # Ввод первых 4 значений (масса, объем, тип, расч. стоимость)
        driver.find_element_by_xpath("(//input[@name='CargoTab1TotalWeight'])[1]").send_keys(mass)
        driver.find_element_by_xpath("(//input[@name='CargoTab1TotalVolume'])[1]").send_keys(volume)
        driver.find_element_by_xpath("(//input[@type='select-one'])[1]").send_keys(cargo_type)
        driver.find_element_by_xpath("//div[@class='bid-cargo__block-chars cargo-tab1']").click()
        sleep(1)
        driver.find_element_by_xpath("(//input[@class='control'])[5]").clear()
        driver.find_element_by_xpath("(//input[@class='control'])[5]").send_keys(rate)

        # Ввод даты
        actions.move_to_element(driver.find_element_by_xpath("//i[@class='fa fa-calendar']")).perform()
        driver.find_element_by_xpath("//i[@class='fa fa-calendar']").click()
        sleep(0.5)
        driver.find_element_by_xpath("//div[@data-date='21']").click()
        sleep(0.5)

        for k in range(len(column1)):
            first_address = start + column1[k]
            second_address = start + column2[k]
            array1 = []
            array2 = []
            array3 = []
            array4 = []
            array5 = []
            print(column1[k], ' - ', column2[k], end=' ')

            # Ввод первого адреса
            actions.move_to_element(driver.find_element_by_xpath("(//img[@alt='Очистить поле'])[1]")).perform()
            driver.find_element_by_xpath("(//img[@alt='Очистить поле'])[1]").click()
            driver.find_element_by_xpath("(//textarea[@row='1'])[1]").send_keys(first_address)
            WebDriverWait(driver, 10).until(ex_cond.presence_of_element_located(
                (By.XPATH, "(//button[@class='bid-input-direction__item'])[1]")))
            sleep(0.1)
            driver.find_element_by_xpath("(//button[@class='bid-input-direction__item'])[1]").click()

            # Ввод второго адреса
            driver.find_element_by_xpath("(//img[@alt='Очистить поле'])[2]").click()
            driver.find_element_by_xpath("(//textarea[@row='1'])[2]").send_keys(second_address)
            WebDriverWait(driver, 10).until(ex_cond.presence_of_element_located(
                (By.XPATH, "(//button[@class='bid-input-direction__item'])[1]")))
            sleep(0.1)
            driver.find_element_by_xpath("(//button[@class='bid-input-direction__item'])[1]").click()
            sleep(10)

            # Получение стоимостей и сроков
            a = driver.find_element_by_xpath("(//div[@class='_right'])[1]").text
            b = driver.find_element_by_xpath("(//span[@class='_info']//span)[1]").text
            c = driver.find_element_by_xpath("(//span[@class='_info']//span)[2]").text
            d = driver.find_element_by_xpath("//div[@class='bid-check__est-days']//span[1]").text

            # Форматирование полученных данных
            for i in a:
                if i.isdecimal() or i == ",":
                    array1.append(i)
            for i in b:
                if i.isdecimal() or i == ",":
                    array2.append(i)
            for i in c:
                if i.isdecimal() or i == ",":
                    array3.append(i)
            for i in d:
                if i.isdecimal():
                    array4.append(i)

            # Добавление в массивы
            m = ''.join(str(i) for i in array1)
            n = ''.join(str(i) for i in array2)
            o = ''.join(str(i) for i in array3)
            p = ''.join(str(i) for i in array4)
            array6.append(m)
            array7.append(n)
            array8.append(o)
            array9.append(p)
            print(m, n, o, p, end=' ')
            m *= 0
            n *= 0
            o *= 0
            p *= 0

            driver.find_element_by_xpath("(//div[contains(@class,'bid-radio radio')]//i)[1]").click()
            driver.find_element_by_xpath("(//div[contains(@class,'bid-radio radio')]//i)[3]").click()
            sleep(8)

            # Получение, форматирование и добавление последнего срока
            e = driver.find_element_by_xpath("//div[@class='bid-check__est-days']//span[1]").text
            for i in e:
                if i.isdecimal():
                    array5.append(i)
            q = ''.join(str(i) for i in array5)
            array10.append(q)
            print(q)
            q *= 0
            sleep(1)
            driver.find_element_by_xpath("(//div[contains(@class,'bid-radio radio')]//i)[2]").click()
            driver.find_element_by_xpath("(//div[contains(@class,'bid-radio radio')]//i)[4]").click()
    except Exception as ex:
        print(ex)


openChrome()
write_in_file = pd.DataFrame({
    'Склад-Склад': array6,
    'От Двери': array7,
    'До Двери': array8,
    'Срок Склад': array9,
    'Срок Дверь': array10,
})
write_in_file.to_excel('C:\Python\Python Parsers\Parser_PEK\Test.xlsx', index=False)
