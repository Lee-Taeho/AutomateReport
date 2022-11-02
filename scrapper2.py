from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import timeit
import time
import os


class Scrapper:

    def __init__(self):
        chrome_options = Options()
        chrome_options.add_argument("--headless")

        # Initialization
        self.driver = webdriver.Chrome(
            chrome_options=chrome_options,
            executable_path="resource/chromedriver")
        self.driver.get("https://talkbot.us/#/login")
        self.driver.implicitly_wait(10)

        self.Chatbot = {"Navien": "N", "Astro": "A"}

    def login(self, username, password):

        uname = driver.find_element(By.ID, "username")
        pw = driver.find_element(By.ID, "password")
        login = self.driver.find_element(By.XPATH, "//jhi-talkbot-login//form/div/button")

        uname.clear()
        uname.send_keys(username)
        pw.clear()
        pw.send_keys(password)
        login.click()

