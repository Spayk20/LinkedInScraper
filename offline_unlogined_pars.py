import time
import undetected_chromedriver as uc
from undetected_chromedriver.options import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import openpyxl
import pandas as pd

options = ChromeOptions()
# options.add_argument( '--headless' )
options.add_experimental_option('prefs', {'intl.accept_languages': 'en,en_US'})
driver = uc.Chrome(options=options, version_main = 111)
_output = {}

def checking_link(link):
    driver.get(link)
    time.sleep(10)
    try:
        h2 = WebDriverWait(driver, timeout=25).until(driver.find_elements(By.TAG_NAME, "h2"))
        print(h2)
    except:
        print('nope')
    # if h2 != "This LinkedIn Page isn’t available":
    #     prs(link)
    # else:
    #     pass

def prs(link):
    website = None
    industries = None
    headquarters = None
    company_type = None
    founded = None
    specialties = None
    company_size = None

    # driver.get(link)
    time.sleep(2)

    grid = driver.find_element(By.CLASS_NAME, "core-section-container__content.break-words")
    labels = WebDriverWait(driver, timeout=5).until(lambda grid: grid.find_elements(By.TAG_NAME, "dt"))
    values = WebDriverWait(driver, timeout=5).until(lambda grid: grid.find_elements(By.TAG_NAME, "dd"))
    num_attributes = min(len(labels), len(values))

    x_off = 0
    for i in range(num_attributes):
        txt = labels[i].text.strip()
        if txt == 'Website':
            website = (values[i + x_off].text.strip())
        elif txt == 'Industries':
            industries = (values[i + x_off].text.strip())
        elif txt == 'Company size':
            company_size = (values[i + x_off].text.strip())
            if len(values) > len(labels):
                x_off = 1
        elif txt == 'Headquarters':
            headquarters = (values[i + x_off].text.strip())
        elif txt == 'Type':
            company_type = (values[i + x_off].text.strip())
        elif txt == 'Founded':
            founded = (values[i + x_off].text.strip())
        elif txt == 'Specialties':
            specialties = (values[i + x_off].text.strip())

    _output['website'] = website
    _output['company_size'] = company_size
    _output['industry'] = industries
    _output['headquarters'] = headquarters
    _output['company_type'] = company_type
    _output['founded'] = founded
    _output['specialities'] = specialties
    df_output = (pd.DataFrame.from_dict([_output], orient='columns'))
    with pd.ExcelWriter('REZ - companies_list.xlsx', engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df_output.to_excel(writer, index=False, startrow=openpyxl.load_workbook('REZ - companies_list.xlsx').worksheets[0].max_row, header=False,)

    # df_output = df_output.append(pd.DataFrame([_output]))
    # df_output = pd.DataFrame.from_dict([_output], orient='columns')
    return print(_output)


lnk_num = 1
files = ["D:\Python\LinkedInParser\isnt_available.html"]
for i in files:
    checking_link(i)
    lnk_num += 1
driver.quit()