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

def url_changed(driver):
    return driver.current_url != initial_url

scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)
sheet_id = "15pCof2Va_ReARy78PhGCAK51AfDeOVXfbG6gCkXpGz4"
workbook = client.open_by_key(sheet_id)

option = webdriver.ChromeOptions()
option.add_argument("--windows-size=1100,660")
option.add_argument("--disable-notification")
option.add_argument("--disable-blink-feature=AutomationControlled")

driver = webdriver.Chrome(options=option)
# driver.get("https://nepalstock.com/company")
# driver.find_element(By.XPATH, "//div[@class='table__perpage']/select[1]").click()
# driver.implicitly_wait(5)
# driver.find_element(By.XPATH, "//div[@class='table__perpage']/select/option[6]").click()
# driver.implicitly_wait(5)
# rows = []
# filler = {}
# # 1. Extract sectors
# sector_options = driver.execute_script("""
#     let select = document.querySelectorAll('.box__filter--field select')[1];
#     let values = [...select.querySelectorAll('option')]
#         .map(o => o.value.trim())
#         .filter(v => v !== '');  // remove the empty 'Sector'
#     return values;
# """)

# # 2. List of existing sheet titles
# existing_sheets = [ws.title for ws in workbook.worksheets()]

# # 3. Create worksheets if not exist
# for sector in sector_options:
#     rows.append([sector])
#     filler[sector] = []
#     if sector not in existing_sheets:
#         workbook.add_worksheet(title=sector, rows=120, cols=10)
#         ws = workbook.worksheet(sector)
#         ws.update_cell(1, 1, 'Stock Symbol')
#         print(f"Created worksheet: {sector}")
#     else:
#         ws = workbook.worksheet(sector)
#         ws.clear()
#         ws.update_cell(1, 1, 'Stock Symbol')
#         print(f"Worksheet already exists: {sector}")

# worksheet = workbook.worksheet("sectorwise")
# worksheet.update(
#     values=rows,
#     range_name=f"A2:A{len(rows)+1}"
# )
# print("sectorwise sheet updated successfully.")

# table_data = driver.execute_script("""
#     let rows = [...document.querySelectorAll("table tbody tr")];
#     return rows.map(r => {
#         let tds = r.querySelectorAll("td");
#         return {
#             symbol: tds[2]?.innerText.trim(),
#             sector: tds[4]?.innerText.trim()
#         };
#     });
# """)

# for entry in table_data:
#     symbol = entry["symbol"]
#     sector = entry["sector"]

#     # Append this symbol to correct sector-buffer
#     filler[sector].append([symbol])   # must be 2D list for gspread

#     # batch size = 50
#     if len(filler[sector]) == 50:
#         ws = workbook.worksheet(sector)

#         # get existing used rows to append properly
#         cell_count = len(ws.col_values(1))  # how many rows already filled
#         start_row = cell_count + 1
#         end_row = start_row + 49

#         ws.update(
#             values=filler[sector],
#             range_name=f"A{start_row}:A{end_row}"
#         )

#         filler[sector] = []  # clear buffer

# for sector, pending_rows in filler.items():
#     if not pending_rows:
#         continue

#     ws = workbook.worksheet(sector)

#     cell_count = len(ws.col_values(1))
#     start_row = cell_count + 1
#     end_row = start_row + len(pending_rows) - 1

#     ws.update(
#         values=pending_rows,
#         range_name=f"A{start_row}:A{end_row}"
#     )

driver.get("https://nepalstock.com/")
initial_url = driver.current_url
stock = "MEN"
driver.find_element(By.XPATH, "//input").send_keys(f"({stock})")
driver.find_element(By.XPATH, "//button[@class='dropdown-item active']").click()
WebDriverWait(driver, 5).until(url_changed)
driver.set_script_timeout(45)
driver.execute_script("""
document.querySelector('#pricehistory-tab')?.click();
""")
sleep(0.1)
result = driver.execute_async_script("""
const done = arguments[arguments.length - 1];
async function wait(ms) {
    return new Promise(res => setTimeout(res, ms));
}

let latest_trade_date = '';
let latest_vol = 0;
let total = 0;
let trade_days = 0;

function parseNum(str){
    return parseInt(str.replace(/,/g,''));
}
function latestState(){
    let row = document.querySelector('div#pricehistorys table tbody tr');
    latest_vol = parseNum(row.children[6]?.innerText.trim())
    latest_trade_date = row.children[1]?.innerText.trim();
}
function sumCurrentPage(){
    let rows = document.querySelectorAll('div#pricehistorys table tbody tr');
    rows.forEach(r => {
        let val = r.children[6]?.innerText.trim();
        if(val) {
            total += parseNum(val);
            trade_days ++;
        }
    });
}

async function scrapeAllPages() {
    latestState();
    sumCurrentPage();

    let initPages = [...document.querySelectorAll('.ngx-pagination li a')];
    if(!initPages.length) return;

    let pageFolds = initPages[initPages.length - 2].querySelectorAll('span')[1].innerText - 1;

    for(let i = 0; i < pageFolds; i++){
        let nextPageNumber = i + 2;

        let pages = [...document.querySelectorAll('.ngx-pagination li a')];
        let nextPageLink = pages.find(a => a.querySelectorAll('span')[1]?.innerText == nextPageNumber);

        if(nextPageLink){
            nextPageLink.click();          
            await wait(1500);  // ⬅ wait for page to fully re-render
            sumCurrentPage();
        }
    }
    return { total, trade_days, latest_vol, latest_trade_date };
}

scrapeAllPages().then(done);
""")
print(result)