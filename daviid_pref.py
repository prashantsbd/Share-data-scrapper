import gspread
from google.oauth2.service_account import Credentials
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common import keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome import options as option
from time import sleep

scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)
sheet_id = "15pCof2Va_ReARy78PhGCAK51AfDeOVXfbG6gCkXpGz4"
workbook = client.open_by_key(sheet_id)
worksheet = workbook.worksheet("daily")

option = webdriver.ChromeOptions()
option.add_argument("--windows-size=1100,660")
option.add_argument("--disable-notification")
option.add_argument("--disable-blink-feature=AutomationControlled")

driver = webdriver.Chrome(options=option)
driver.get("https://nepalstock.com/company")
driver.find_element(By.XPATH, "//div[@class='table__perpage']/select[1]").click()
driver.implicitly_wait(5)
driver.find_element(By.XPATH, "//div[@class='table__perpage']/select/option[6]").click()
driver.implicitly_wait(5)
rows = []
chiz = len(driver.find_elements(By.XPATH, "//tbody/tr")) + 1

for i in range(1, chiz):
    symbol = driver.find_element(By.XPATH, f"//tr[{i}]/td[3]").text
    name = driver.find_element(By.XPATH, f"//tr[{i}]/td[2]").text
    stock = f"({symbol}) {name}"

    rows.append([stock])  # important → list inside list (for columns)

    # when batch reaches 10 rows, send request
    if len(rows) == 10:
        start_row = i + 1 - 9  # 10 rows back
        end_row = i + 1
        cell_range = f"A{start_row}:A{end_row}"
        worksheet.update(values=rows, range_name=cell_range)
        rows = []  # clear buffer

# update leftover rows (<10)
if rows:
    start_row = chiz - len(rows) + 1
    end_row = chiz
    cell_range = f"A{start_row}:A{end_row}"
    worksheet.update(values=rows, range_name=cell_range)