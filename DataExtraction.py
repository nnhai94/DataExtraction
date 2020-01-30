from selenium import webdriver
import chromedriver_binary
from bs4 import BeautifulSoup
import pandas as pd

def Scrape(site):

    driver = webdriver.Chrome(executable_path=r'C:\Users\Admin PC\Desktop\Project\chromedriver_win32\chromedriver.exe')
    # Get the website
    driver.get('https://www.asx.com.au/asx/share-price-research/company/CBA/details')

    soup = BeautifulSoup(driver.page_source) # driver.page_source return page source code in html format

    # Find all the tables on page
    table = soup.find_all('table')

    # Read all tables into dataframe
    df = pd.read_html(str(table)) # df is a list of dataframes

    # Convert all the dataframes into excel workbook
    with pd.ExcelWriter('Company details.xlsx') as writer:  
        df[0].to_excel(writer, sheet_name='Service_info')
        df[1].to_excel(writer, sheet_name='Board_of_directors')
        df[2].to_excel(writer, sheet_name='Secretaries')

if __name__ == '__main__':
    site = input('Input the company details page')
    Scrape(site)
