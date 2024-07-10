from selenium import webdriver
from selenium.webdriver.common.keys import Keys             # simuate keystrokes
from selenium.webdriver.common.by import By                 # import the By thing idk
from selenium.webdriver.support.ui import Select            # to select dropdown
from selenium.webdriver.support.wait import WebDriverWait   # wait for dropdown to appear
from selenium.webdriver.support import expected_conditions as EC
import time
import requests                 # handle cookies
import pandas as pd
import math


# ---PREPARE LIST OF LISTS AND FUNCTIONS FOR DATA---

inbox_list = []
headers = ["Kode", "Nomor Registrasi", "Terbit", "Nama Produk", "Merk",
           "Kemasan", "Pendaftaran", "Alamat Pendaftaran", "Dll."]
inbox_list.append(headers)

# function to open and close each search page's item
# @param index: index of item's (non-zero) index on the search page
# @return data: all the scraped data saved as a string
def iterate_inbox(index):
    
    # BUG: this method always gets hit with timeout
    driver.implicitly_wait(20)  # set timeout limit

    # click on item to pop up window: Wait is needed in case of race condition
    try:
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="inbox-list"]/div[{}]'.format(index))))
        time.sleep(5)   # extra insurance because the site is slow lmao
    finally:
        time.sleep(5)
        item = driver.find_element(By.XPATH, '//*[@id="inbox-list"]/div[{}]'.format(index))
        item.click()

    # get window info
    max_attempts = 5    # to try again if timeout occurs
    for i in range(max_attempts):
        try:
            driver.implicitly_wait(60)
            WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="detailobat"]/table/tbody')))
            time.sleep(5)
            break
        except TimeoutError:
            continue
    obat = driver.find_element(By.XPATH, '//*[@id="detailobat"]/table/tbody')
    data = obat.text

    # close window
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="exampleModal2"]/div/div/div[3]/button')))
    time.sleep(5)
    close = driver.find_element(By.XPATH, '//*[@id="exampleModal2"]/div/div/div[1]/button')
    close.click()

    return data

# function to move data into lists
# @param input: list made by splitting input at \n
def to_list(input):
    inner_list = []
    count = 0
    index = 1
    for line in input.split("\n"):
        if count == 8:      # all columns are filled
            print(index)
            print(inner_list)
            data = iterate_inbox(index)     # click on item to scrape data
            inner_list.append(data)
            inbox_list.append(inner_list)
            print(inner_list)
            print()
            inner_list = []
            count = 0
            index += 1
        elif count == 2:    # strip 'Tertib: '
            line = line.strip("Tertib: ")
        elif count == 4:    # strip 'Merk: '
            line = line.strip("Merk: ")
        elif count == 5:    # strip 'Kemasan: '
            line = line.strip("Kemasan: ")
        inner_list.append(line)
        count += 1
    # append the last item in input to inbox_list
    print(index)
    print(inner_list)
    data = iterate_inbox(index)     # click on item to scrape data
    inner_list.append(data)
    inbox_list.append(inner_list)
    print(inner_list)
    print()
    print("---END---")
    return


# ---SET UP DRIVER---

# Selenium v4.6.0 or above: no need to set driver path
driver = webdriver.Chrome()
session = requests.Session()

url = "https://cekbpom.pom.go.id/search_home_produk"
driver.get(url)

# set session for POST requests later
res = session.get(url, verify=False)
# cookie = res.headers.get("Set-Cookie")
# token = cookie[7:39]    # hardcode, regex later

# searching for search bar, then clicking it
driver.find_element(By.CLASS_NAME, "form-control").click()
time.sleep(5)

# wait for dropdown to appear, then click it + select option
WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "st_filter"))).click()
select = Select(driver.find_element(By.ID, "st_filter"))
time.sleep(5)
select.select_by_visible_text("Nama Produk")
driver.find_element(By.ID, "st_filter").click()     # click off dropdown to close it

# input text in search bar and hit enter to search
time.sleep(5)
search = driver.find_element(By.ID, "input_search")
search.send_keys("vitamin d3")
time.sleep(5)
search.send_keys(Keys.RETURN)


# --- FIRST RESULT PAGE ---

# get search results (that are requested with POST)
payload = {
    "st_filter": "2",
    "input_search": "vitamin d3",
    "from_home_flag": "Y"
}
res = session.post(url, data=payload, verify=False)

# store number of entries
WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CLASS_NAME, "kt-inbox__perpage")))
data_per_page = driver.find_element(By.CLASS_NAME, "kt-inbox__perpage").text.split(" ")
print(data_per_page)
next_prev = int(data_per_page[2])    # how many results per page
entries = int(data_per_page[4])      # total number of results
# next_prev = 10
# entries = 135     # hardcode for now, because data_per_page is inconsistent

# get inbox 'preview' search results, make them a list of lists
inbox = driver.find_element(By.ID, "inbox-list")
# print(inbox.text)
to_list(inbox.text)


# --- OTHER RESULT PAGES: iterate through each page, get search results ---

page_no = 1
for i in range(math.ceil(entries / next_prev) - 1):
    time.sleep(5)
    try:
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CLASS_NAME, "flaticon2-right-arrow")))
        time.sleep(5)
    finally:
        driver.find_element(By.CLASS_NAME, "flaticon2-right-arrow").click()
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, "flaticon2-right-arrow")))
        print()
        print("PAGE NO: {}".format(page_no + 1))
        print()
        payload = {
            "st_filter": "2",
            "input_search": "vitamin d3",
            "from_home_flag": "Y",
            "offset": str(next_prev * page_no - 9),
            "next_prev": str(next_prev * page_no),
            "count_data_all_produk": str(entries),
            "marked": "next"
        }
        res = session.post(url, data=payload, verify=False)
        driver.implicitly_wait(20)  # set timeout limit
        inbox = driver.find_element(By.ID, "inbox-list")
        # print(inbox.text)
        time.sleep(5)
        to_list(inbox.text)
        page_no += 1
        time.sleep(5)   # give time for response to load


# write to an excel sheet
df = pd.DataFrame(inbox_list)
writer = pd.ExcelWriter("produk_vitD3.xlsx", engine="xlsxwriter")
df.to_excel(writer, sheet_name="Produk Vitamin D3", index=True)
writer.close()

# # delay so program doesn't quit immediately
# time.sleep(5)

print("PROGRAM DONE")

driver.quit()