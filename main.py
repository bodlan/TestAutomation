from RPA.Browser.Selenium import Selenium
from RPA.PDF import PDF
from datetime import timedelta
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
    agencies_and_value = []
    agencies = browser_lib.find_elements("xpath://div[@*='agency-tiles-container']//a/span")
    for agency in agencies:
        agencies_and_value.append(agency.text)
    agencies_data = {
        'Agency': agencies_and_value[::2],
        'Amount': agencies_and_value[1::2]
    }
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

    uii_links_raw = browser_lib.find_elements("xpath://div[@class='dataTables_scrollBody']//tbody//a")
    uii_links = [url.get_attribute("href") for url in uii_links_raw]
    investments_body_with_links_raw = browser_lib.find_elements("xpath://tbody/descendant::a[starts-with(@href,'/')]"
                                                                "/ancestor::tr/td")
    investments_body_without_links_raw = browser_lib.find_elements("xpath://div[@class='dataTables_scrollBody']"
                                                                   "//tbody/tr[not(descendant::a)]/td")
    investments_body_with_links = [cell.text for cell in investments_body_with_links_raw]
    investments_body_without_links = [cell.text for cell in investments_body_without_links_raw]
    file_names = []
    for link in uii_links:
        file_names.append(link.split(sep="/")[-1])
    titles = investments_body_with_links[2::7]
    uii_names = investments_body_with_links[::7]
    individual_investment = {
        'UII': investments_body_with_links[::7] + investments_body_without_links[::7],
        'Bureau': investments_body_with_links[1::7] + investments_body_without_links[1::7],
        'Investment Title': investments_body_with_links[2::7] + investments_body_without_links[2::7],
        'Total FY2021 Spending ($M)': investments_body_with_links[3::7] + investments_body_without_links[3::7],
        'Type': investments_body_with_links[4::7] + investments_body_without_links[4::7],
        'CIO Rating': investments_body_with_links[5::7] + investments_body_without_links[5::7],
        '# of Projects': investments_body_with_links[6::7] + investments_body_without_links[6::7]
    }
    individual_investment_data_frame = pd.DataFrame(individual_investment)
    individual_investment_data_frame.to_excel('./output/Individual_Investment.xlsx', sheet_name="investments",
                                              index=False)
    return zip(uii_links, uii_names, titles, file_names)


def get_business_case(data):
    files_count = len(os.listdir(path_to_download_directory))
    for link, uii, investment_title, file_name in data:
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
                        wait=True
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
        else:
            print(f"Data mismatched in {os.getcwd()}/output/{file_name}.pdf")


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
