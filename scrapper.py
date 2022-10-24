import timeit

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
import os

def highlight(element, effect_time, color, border):
    driver = element._parent

    def apply_style(s):
        driver.execute_script("arguments[0].setAttribute('style', arguments[1]);",
                              element, s)

    original_style = element.get_attribute('style')
    apply_style("border: {0}px solid {1};".format(border, color))
    time.sleep(effect_time)
    apply_style(original_style)


chrome_options = Options()
chrome_options.add_argument("--headless")


# Initialization
driver = webdriver.Chrome(
    chrome_options=chrome_options,
    executable_path="resource/chromedriver")
driver.get("https://talkbot.us/#/login")
driver.implicitly_wait(10)


# Login
uname = driver.find_element(By.ID, "username")
pw = driver.find_element(By.ID, "password")
login = driver.find_element(By.XPATH, "//jhi-talkbot-login//form/div/button")

uname.clear()
uname.send_keys("botbuilder")
pw.clear()
pw.send_keys("SaltluxUSVN1234!@#$")
login.click()



# Finding Arelli
# folder = driver.find_element(By.XPATH, '//div[@class="treeList"]//label[@title="Arelli"]/../../../*[@class = "fa font-tree-item fa-folder"]')
# folder.click()
# Project = driver.find_element(By.XPATH, '//*[@class = "treeList"]//span[text()="Astro"]')
# Project.click()

# Finding NavienMate
folder = driver.find_element(By.XPATH, '//div[@class="treeList"]//label[@title="NavienMate"]')
folder.click()

driver.implicitly_wait(10)

for i in range(3):
    try:
        driver.find_element(By.XPATH, "//div[@class = 'botList']/div/ol[1]/li[4]/p[1]").click()
        break
    except Exception as e:
        print(e)

driver.find_element(By.XPATH, '//*[@id="nav-left"]/ul/li[last()]').click()


# pick a starting date and ending date


def select_date(start_or_end, tab, date):
    month_to_name = {
        "01": "Jan", "02": "Feb", "03": "Mar",
        "04": "Apr", "05": "May", "06": "Jun",
        "07": "Jul", "08": "Aug", "09": "Sep",
        "10": "Oct", "11": "Nov", "12": "Dec"
    }

    tab_to_index = {
        "Overall": 1,
        "Email": 7
    }

    # click date icon
    driver.find_elements(By.XPATH, f'//*[@class = "bi-board-tab"]/jhi-tab[position() = {tab_to_index[tab]}]//div[@class = "bi-filter-cont"]//*[@class="icon-date"]')[start_or_end].click()

    # Choose a year
    driver.find_elements(By.XPATH, '//span[@class ="owl-dt-control-button-content"]')[1].click()
    driver.find_element(By.XPATH, f'//tbody[@class = "owl-dt-calendar-body"]//span[text() = "{date[0]}"]').click()

    # Choose a month
    driver.find_element(By.XPATH, f'//owl-date-time-calendar//td/span[text()="{month_to_name[date[1]]}"]').click()

    # Choose a date
    driver.find_element(By.XPATH,
                        f'//table[@class = "owl-dt-calendar-table"]//td[@aria-label="{date[0]}.{date[1]}.{date[2]}"]').click()


# starting date
select_date(0, "Overall",["2022", "09", "01"])

# ending date
select_date(1, "Overall",["2022", "09", "30"])

driver.find_element(By.XPATH, "//div[@class= 'bi-filter-cont']//button").click()

# screenshot = driver.save_screenshot("screenshot.png")

# driver.find_element(By.XPATH, '//*[@class= "chart-content chart-content-sp rb"]/*').screenshot("temp.png")

# download the data

driver.find_element(By.XPATH, '//*[@class = "bi-download-btn"]/button').click()

def download_wait(path_to_downloads):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 20:
        time.sleep(1)
        dl_wait = False
        for fname in os.listdir(path_to_downloads):
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds

download_wait("./")

def get_xls_file_name():
    for fname in os.listdir("./"):
        if fname.endswith(".xls"):
            return fname
    return None



# Expand the window full screen
S = lambda X: driver.execute_script('return document.body.parentNode.scroll' + X)
driver.set_window_size(S('Width'), S('Height'))  # May need manual adjustment


# take screenshot of the data
data = driver.find_element(By.XPATH, '//*[@class = "tab-content"]').find_elements(By.XPATH,
                                                                                  './div[@class = "tab-content-hor"]')

i = 1
time.sleep(1)
for i in range(1,6):
    # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//*[@class = "bi-board-tab"]/jhi-tab[position() = 1]//div[@class = "tab-content"]/div[@class = "tab-content-hor" and position() ={i}]'))).screenshot(f'web-element-{i}.png')
    driver.find_element(By.XPATH, f'//*[@class = "bi-board-tab"]/jhi-tab[position() = 1]//div[@class = "tab-content"]/div[position() ={i}]').screenshot(f'web-element-{i}.png')
    # element.screenshot(f'web-element-{i}.png')

# driver.find_element(By.TAG_NAME, 'body').screenshot('web_screenshot.png')


# direct to email menu
tab = driver.find_element(By.CLASS_NAME, "tabWarp")
email_tab = tab.find_element(By.XPATH, './a[text() = "Emails"]')
email_tab.click()

# Select dates
select_date(0, "Email", ["2022", "09", "01"])
select_date(1, "Email", ["2022", "09", "30"])

driver.find_elements(By.XPATH, "//div[@class= 'bi-filter-cont']//button")[6].click()

WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@class = "bi-board-tab"]/jhi-tab[position() = 7]//div[@class = "tab-content-hor"]'))).screenshot("email.png")
# data = driver.find_element(By.XPATH, '//*[@class = "bi-board-tab"]/jhi-tab[position() = 7]//div[@class = "tab-content-hor"]')
# data.screenshot("email.png")





# highlight(Astro, 3, "red", 2)
# Astro.click()
driver.quit()
# //*[@id="owl-dt-picker-0"]/div[2]/owl-date-time-calendar/div[2]/owl-date-time-year-view/table/tbody/tr[3]/td[2]/span
