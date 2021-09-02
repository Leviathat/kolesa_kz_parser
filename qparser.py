# coding=utf-8
from selenium import webdriver
import xlsxwriter
import re

URL = "https://kolesa.kz/cars/body-coupe/mercedes-benz/e-klasse/almaty/?_sys-hasphoto=2&auto-custom=2&auto"\
      "-car-transm=2345&car-dwheel=3&auto-car-order=1&auto-car-volume[to]=4&year[from]=2013&year["\
      "to]=2018&price[to]=20000010"
kolesa_shows = []  # links ( kolesa.kz/a/show/<int:pk> )
cars_info = [["Год", "Цена", "Цвет", "Пробег"]]  # Car objects list


class Car:

    def __init__(self, year, amount, color, mileage):
        self.year = year
        self.amount = amount
        self.color = color
        self.mileage = mileage


def write_into_xlsx(car_info):
    workbook = xlsxwriter.Workbook('list.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column(0, 3, 25)
    worksheet.set_column(0, 0, 30)

    for row_num, data in enumerate(car_info):
        worksheet.write_row(row_num, 0, data)

    workbook.close()


class KolesaParser(object):

    def __init__(self, driver):
        self.driver = driver

    def parse(self, url):
        self.driver.get(url)
        links = self.driver.find_elements_by_css_selector('div > div.a-info-top > span.a-el-info-title > a')

        for link in links:
            kolesa_shows.append(link.get_attribute('href'))
        self.driver.quit()

    def get_car_info(self, car_link_list):

        for car_link in car_link_list:
            driver = webdriver.Chrome("chromedriver")
            driver.get(car_link)

            year = driver.find_element_by_xpath('/html/body/main/div/div/div/header/h1/span[3]')
            amount = driver.find_element_by_xpath('/html/body/main/div/div/div/section/div[1]/div[1]/div[1]/div[1]')
            color = driver.find_element_by_xpath('/html/body/main/div/div/div/section/div[1]/div[1]/div[2]/dl[7]/dd')
            mileage = driver.find_element_by_xpath('/html/body/main/div/div/div/section/div[1]/div[1]/div[2]/dl[4]/dd')
            car = [year.text, amount.text, color.text, mileage.text]
            cars_info.append(car)

            driver.quit()


def main():

    chrome_driver_binary = "chromedriver"
    driver = webdriver.Chrome(chrome_driver_binary)
    parser = KolesaParser(driver)
    parser.parse(URL)
    parser.get_car_info(kolesa_shows)
    write_into_xlsx(cars_info)


if __name__ == '__main__':
    main()
