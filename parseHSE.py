from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import requests

def parseSelenium(conkurses):
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)

    driver.get("https://ba.hse.ru/base2024")
    span = driver.find_elements(By.CSS_SELECTOR, "._link._link--pseudo")[1]

    driver.execute_script("arguments[0].click();", span)
    parent = ((span.find_elements(By.XPATH, "..")[0]).find_elements(By.XPATH, "..")[0]).find_elements(By.XPATH, "..")[0]
    table = parent.find_elements(By.CSS_SELECTOR, ".foldable table tbody tr")
    for row in table:
        tds = row.find_elements(By.CSS_SELECTOR, "td")
        print(tds[0].text)
        if tds[0].text in conkurses:
            url = (tds[1].find_elements(By.CSS_SELECTOR, 'a')[0].get_attribute("href"))


            response = requests.get(url)
            file_Path = "loads/"+"ВШЭ_"+tds[0].text+'.xlsx'

            if response.status_code == 200:
                with open(file_Path, 'wb') as file:
                    file.write(response.content)
                print('File downloaded successfully')
            else:
                print('Failed to download file')

