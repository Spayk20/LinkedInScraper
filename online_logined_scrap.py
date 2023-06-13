import time
import undetected_chromedriver as uc
from undetected_chromedriver.options import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import openpyxl
import pandas as pd

def login(driver, LINK_LOGIN, LINK_PASSWORD):
    try:
        driver.get('https://www.linkedin.com/login/')
        time.sleep(5)

        login_field = driver.find_element(By.ID, 'username')
        login_field.send_keys(LINK_LOGIN)

        login_field = driver.find_element(By.ID, 'password')
        login_field.send_keys(LINK_PASSWORD)

        driver.find_element(By.CLASS_NAME, "btn__primary--large.from__button--floating").click()
    except:
        print('Login Error')

def checking_link(link):
    driver.get(link + 'about/')
    if driver.find_element(By.TAG_NAME, "h2").text != "This LinkedIn Page isn’t available":
        scrap(link)
    else:
        pass

def scrap(link):
    website = None
    industries = None
    headquarters = None
    company_type = None
    founded = None
    specialties = None
    company_size = None
    _output = {}

    # driver.get(link + 'about/')
    time.sleep(2)

    grid = driver.find_element(By.CLASS_NAME, "org-grid__content-height-enforcer")
    labels = WebDriverWait(driver, timeout=5).until(lambda grid: grid.find_elements(By.TAG_NAME, "dt"))
    values = WebDriverWait(driver, timeout=5).until(lambda grid: grid.find_elements(By.TAG_NAME, "dd"))
    num_attributes = min(len(labels), len(values))

    x_off = 0
    for i in range(num_attributes):
        txt = labels[i].text.strip()
        if txt == 'Website':
            website = (values[i + x_off].text.strip())
        elif txt == 'Industry':
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
        df_output.to_excel(writer, index=False, startrow=openpyxl.load_workbook('REZ - companies_list.xlsx').worksheets[0].max_row, header=False)
    time.sleep(2)
    return print(_output)


with open("pss.txt") as pss:
    lines = pss.readlines()
    LINK_LOGIN, LINK_PASSWORD = lines[0], lines[1]

options = ChromeOptions()
options.add_experimental_option('prefs', {'intl.accept_languages': 'en,en_US'})
driver = uc.Chrome(options=options, version_main=111)

login(driver, LINK_LOGIN, LINK_PASSWORD)

df_input = pd.read_excel('companies_list.xlsx')
num = 1
for link in df_input['Link']:
    print(f'Processing...{num} , link')
    checking_link(link)
    num += 1
driver.quit()
