from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import pandas
from selenium.webdriver.chrome.service import Service

from time import sleep

options = Options()
read_excel_file = pandas.read_excel("newspapers.xlsx")

df = read_excel_file.loc[~pandas.isnull(read_excel_file["source III"])]
for index, row in df.iterrows():
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )
    name_newspapper = row["Newspaper"]
    url = row["source III"]
    driver.maximize_window()
    driver.get(url)
    sleep(1)
    WebDriverWait(driver, 40).until(
        EC.presence_of_element_located((By.ID, "nombre"))
    ).send_keys(name_newspapper)

    WebDriverWait(driver, 40).until(
        EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="formaReporte"]/div[3]/div/button')
        )
    ).click()
    sleep(1)
    try:
        check_visibility = WebDriverWait(driver, 40).until(
            EC.visibility_of_element_located(
                (By.XPATH, "/html/body/main/div/div[2]/div[2]")
            )
        )
        check_visibility.click()

        WebDriverWait(driver, 40).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/main/div/div[3]/div/ul/li[4]/a")
            )
        ).click()
    except Exception as e:
        driver.quit()
        continue

    alltr = driver.find_elements(By.CSS_SELECTOR, "#tablaDistribusciones > tbody>tr")
    municipalities_distributed_name = list()

    for i in range(len(alltr)):
        td_element = alltr[i].find_elements(By.TAG_NAME, "td")
        municipalities_distributed_name.append(td_element[2].text)

    municipalities_distributed_name = [i for i in municipalities_distributed_name if i]
    municipalities_distributed_name_comma = ", ".join(municipalities_distributed_name)
    read_excel_file.at[index, "Municipalities distributed"] = (
        municipalities_distributed_name_comma
    )
    driver.quit()

read_excel_file.to_excel("Newspapers_Output.xlsx", index=None)
