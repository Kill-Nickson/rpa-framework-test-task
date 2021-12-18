import os
import time

from openpyxl import Workbook
import configparser
from datetime import timedelta

import selenium.common.exceptions
from RPA.Browser.Selenium import Selenium


def open_the_website(browser, url):
    browser.open_available_browser(url)


def write_list_to_xlsx_sheet(data_list, workbook, sheet_name):
    workbook.create_sheet(sheet_name)
    sheet = workbook[sheet_name]

    for row in data_list:
        sheet.append(row)


def collect_agencies_spendings_to_workbook(browser, workbook):
    dive_in_btn = browser.find_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a')
    dive_in_btn.click()

    browser.wait_until_element_is_visible('//*[@id="agency-tiles-widget"]/div/div[1]/div[1]')

    agencies_titles = [el.text for el in browser.find_elements(
        '//*[@id="agency-tiles-widget"]/div/div/div/div/div/div/div[1]/a/span[1]')]
    agencies_spendings = [el.text for el in browser.find_elements(
        '//*[@id="agency-tiles-widget"]/div/div/div/div/div/div/div[1]/a/span[2]')]
    agencies_info = list(zip(agencies_titles, agencies_spendings))
    write_list_to_xlsx_sheet(agencies_info, workbook, 'Agencies')


def open_agency_page(browser, agency):
    agencies = browser.find_elements('//*[@id="agency-tiles-widget"]/div/div/div/div/div/div/div[1]/a')
    for a in agencies:
        if agency in a.text:
            a.click()
            return True
    return False


def collect_agency_ind_investments_to_workbook(browser, workbook):
    browser.wait_until_element_is_visible('//*[@id="investments-table-object_info"]', timeout=timedelta(seconds=60))
    ind_invests_entries_div = browser.find_element('//*[@id="investments-table-object_info"]')
    max_ind_invests_entries = ind_invests_entries_div.text.split('of ')[1].split(' ')[0]

    # Show all entries
    browser.select_from_list_by_value('//*[@id="investments-table-object_length"]/label/select', '-1')
    browser.wait_until_element_contains('//*[@id="investments-table-object_info"]',
                                        f'{max_ind_invests_entries} of {max_ind_invests_entries}',
                                        timeout=timedelta(seconds=60))
    table_rows = []
    table_trs = browser.find_elements(f'//*[@id="investments-table-object"]/tbody/tr')
    for i in range(1, len(table_trs) + 1):
        tds = [td.text for td in browser.find_elements(
            f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td')]
        table_rows.append(tds)
    write_list_to_xlsx_sheet(table_rows, workbook, 'Individual Investments')


def download_pdf_of_accessible_uuis(browser, href_list):
    browser.wait_until_element_is_visible('//*[@id="investments-table-object"]/tbody/tr/td[1]/a', timeout=timedelta(60))

    for url in href_list:
        browser.go_to(url)
        browser.wait_until_page_contains('Download Business Case PDF', timeout=timedelta(seconds=60))
        browser.find_element('//*[@id="business-case-pdf"]/a').click()

        max_download_timeout = 5
        current_download_timeout = 0
        while True:
            time.sleep(0.1)
            current_download_timeout += 0.1
            if os.path.exists(os.path.join(os.getcwd(), 'uuis', url.split('/')[-1] + '.pdf')) or \
                    int(current_download_timeout) > max_download_timeout:
                break
        browser.go_back()


def main():
    browser_lib = Selenium()
    browser_lib.set_download_directory(os.path.join(os.getcwd(), 'uuis'))
    wb = Workbook()

    try:
        open_the_website(browser_lib, "http://itdashboard.gov/")
        collect_agencies_spendings_to_workbook(browser_lib, wb)

        config = configparser.ConfigParser()
        config.read("config.ini")

        agency = config["Parser"]["agency"]
        if open_agency_page(browser_lib, agency):
            collect_agency_ind_investments_to_workbook(browser_lib, wb)

            uploadable_uui_list = browser_lib.find_elements(f'//*[@id="investments-table-object"]/tbody/tr/td[1]/a')
            uploadable_uui_href_list = [uui.get_attribute('href') for uui in uploadable_uui_list]
            download_pdf_of_accessible_uuis(browser_lib, uploadable_uui_href_list)

        wb.save('data.xlsx')
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()
