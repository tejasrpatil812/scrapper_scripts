import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import xlsxwriter

ser = Service("/Users/tejaspatil/Common/Code/Python/.venv/chromedriver")
driver = webdriver.Chrome(service=ser)

URL = "https://www.levels.fyi/internships/"
SCROLL_PAUSE_TIME = 0.5
workbook = xlsxwriter.Workbook('Companies.xlsx')


def scroll_to_end():
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(SCROLL_PAUSE_TIME)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break

    last_height = new_height

def setup_excel_sheet():
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Company')
    worksheet.write('B1', 'Pay')
    worksheet.write('C1', 'Link')
    return worksheet

if __name__ == "__main__":
    print(f"Using Chrome Driver to open {URL}")
    try:
        driver.get(URL)
        time.sleep(2)
        
        print(f"Page loaded extracting data now")
        scroll_to_end()
        companies = driver.find_elements("xpath", '//tbody//tr[@data-index]')
        print(f"Extracted {len(companies)} Companies data")

        worksheet = setup_excel_sheet()
        index = 1

        for company in companies:
            name = company.find_element("xpath", ".//td[contains(@class,'name')]//h6").text
            try:
                comp = company.find_element("xpath", ".//td[contains(@class,'salary')]//h6").text
            except:
                comp = "N/A"
            link = company.find_element("xpath", ".//td[contains(@class,'apply')]//a").get_attribute("href")
            
            worksheet.write(index, 0, name)
            worksheet.write(index, 1, comp)
            worksheet.write(index, 2, link)
            index+=1
    
    except Exception as e:
        print(f"Error : {e}")

    driver.close()
    workbook.close()
