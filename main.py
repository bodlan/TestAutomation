from RPA.Browser.Selenium import Selenium
from RPA.PDF import PDF
from datetime import timedelta
from collections import defaultdict
import pandas as pd
import os
import time
import re

selected_agency = "National Science Foundation"
page = "http://itdashboard.gov/"
path_to_download_directory = f"{os.getcwd()}/output"

browser_lib = Selenium()
browser_lib.set_download_directory(directory=path_to_download_directory)


def open_the_website(url):
    browser_lib.open_available_browser(url)


def go_to_agencies_page():
    browser_lib.maximize_browser_window()
    browser_lib.scroll_element_into_view("xpath://a[@href='#home-dive-in']")
    browser_lib.click_element("xpath://a[@href='#home-dive-in']")
    browser_lib.wait_until_element_is_visible("xpath://div[@*='agency-tiles-container']")


def get_agencies_amounts():
    agencies_data = defaultdict(list)
    agency_cnt = browser_lib.get_element_count("xpath://div[@*='agency-tiles-container']"
                                                 "//span[contains(concat(' ', normalize-space(@class),' '), 'h4 w200')]")
    count = 0
    xpath_for_used_agency_names = "contains(text(), '#')"
    while count < agency_cnt:
        agency_name = browser_lib.find_element(f"xpath://div[@*='agency-tiles-container']"
                                               f"//span[contains(concat(' ', normalize-space(@class),' '), 'h4 w200')"
                                               f" and not ({xpath_for_used_agency_names})]")
        agency_name_text = agency_name.text
        agency_amount = browser_lib.find_element(f"xpath://div[@*='agency-tiles-container']"
                                                 f"//span[contains(concat(' ', normalize-space(@class),' '), 'h4 w200')"
                                                 f" and not ({xpath_for_used_agency_names})]/ancestor::a/span[2]")
        xpath_for_used_agency_names += f" or contains(text(), '{agency_name_text}')"
        count += 1
        agencies_data['Agency'].append(agency_name_text)
        agencies_data['Amount'].append(agency_amount.text)
    agencies_data_frame = pd.DataFrame(agencies_data)
    agencies_data_frame.to_excel('output/Agencies.xlsx', sheet_name="agencies", index=False)


def get_agency_individual_investment():
    browser_lib.scroll_element_into_view("xpath://span[text()='{}']/ancestor::a".format(selected_agency))
    browser_lib.click_link("xpath://span[text()='{}']/ancestor::a".format(selected_agency))
    browser_lib.execute_javascript("window.scrollBy(0,1500)")
    browser_lib.wait_until_element_is_visible("xpath://select[@name]", timeout=timedelta(seconds=15))
    browser_lib.scroll_element_into_view("xpath://select[@name]")
    browser_lib.select_from_list_by_value("xpath://select[@name]", '-1')
    browser_lib.wait_until_page_does_not_contain_element(
        "xpath://div[@class='dataTables_paginate paging_full_numbers']/a[@class='paginate_button next']",
        timeout=timedelta(seconds=10))

    table = browser_lib.find_element('css:#investments-table-object')
    individual_investment = defaultdict(list)
    rows = table.find_elements_by_xpath('//tbody/tr[@role]')
    uii_link_list = []
    uii_names_list = []
    titles_list = []
    file_names = []

    def get_element_from_xpath(xpath):
        return ''.join([cell.text for cell in row.find_elements_by_xpath(xpath)])

    for row in rows:
        uii_link_raw = row.find_elements_by_xpath('td[1]/a')
        uii_link = [url.get_attribute("href") for url in uii_link_raw]
        individual_investment['UII'].append(get_element_from_xpath('td[1]'))
        individual_investment['Bureau'].append(get_element_from_xpath('td[2]'))
        individual_investment['Investment Title'].append(get_element_from_xpath('td[3]'))
        individual_investment['Total FY2021 Spending ($M)'].append(get_element_from_xpath('td[4]'))
        individual_investment['Type'].append(get_element_from_xpath('td[5]'))
        individual_investment['CIO Rating'].append(get_element_from_xpath('td[6]'))
        individual_investment['# of Projects'].append(get_element_from_xpath('td[7]'))
        if uii_link:
            uii_link_list.append(''.join(uii_link))
            uii_names_list.append(get_element_from_xpath('td[1]'))
            titles_list.append(get_element_from_xpath('td[3]'))
            file_names.append(''.join(uii_link).split(sep="/")[-1])
    individual_investment_data_frame = pd.DataFrame(individual_investment)
    individual_investment_data_frame.to_excel('./output/Individual_Investment.xlsx', sheet_name="investments",
                                              index=False)
    return zip(uii_link_list, uii_names_list, titles_list, file_names)


def get_business_case(data):
    files_count = len(os.listdir(path_to_download_directory))
    for link, uii, investment_title, file_name in data:
        print("link:", link, "ui:", uii, "title:", investment_title, "file name:", file_name)
        browser_lib.execute_javascript("window.open('{}')".format(link))
        windows = browser_lib.get_window_handles()
        browser_lib.switch_window(windows[-1])
        browser_lib.wait_until_element_is_visible("id:business-case-pdf",
                                                  timeout=timedelta(seconds=10))
        browser_lib.scroll_element_into_view("xpath://div[@id='business-case-pdf']/a")
        browser_lib.click_link("xpath://div[@id='business-case-pdf']/a")
        wait = True
        while wait:
            time.sleep(1)
            if len(os.listdir(path_to_download_directory)) > files_count:
                wait = False
                files = os.listdir(path_to_download_directory)
                for fname in files:
                    if fname.endswith('.crdownload'):
                        wait = True
        files_count += 1
        pdf = PDF()
        pdf_text_raw = pdf.get_text_from_pdf(f"./output/{file_name}.pdf", pages=[1])
        pdf_text = pdf_text_raw.get(1)
        pdf_uii = re.search('\(UII\): (.+?)Section B', pdf_text).group(1)
        pdf_investment_title = re.search('Investment: (.+?)2', pdf_text).group(1)
        if pdf_uii == uii:
            if pdf_investment_title == investment_title:
                print(f"Data matched in {os.getcwd()}/output/{file_name}.pdf")
            else:
                print(f"Data mismatched in {os.getcwd()}/output/{file_name}.pdf")
                print(f"Mismatched data in Investment title: {pdf_investment_title} and {investment_title}")
        elif pdf_investment_title == investment_title:
            print(f"Data mismatched in {os.getcwd()}/output/{file_name}.pdf")
            print(f"Mismatched data in UII: {pdf_uii} and {uii}")
        else:
            print(f"Data mismatched in {os.getcwd()}/output/{file_name}.pdf")
            print(f"Mismatched data in UII and Investment title:"
                  f"\n{pdf_uii} and {uii}"
                  f"\n{pdf_investment_title} and {investment_title}")


def main():
    try:
        open_the_website(page)
        go_to_agencies_page()
        get_agencies_amounts()
        data = get_agency_individual_investment()
        get_business_case(data)
    finally:
        browser_lib.close_all_browsers()


if __name__ == '__main__':
    main()
