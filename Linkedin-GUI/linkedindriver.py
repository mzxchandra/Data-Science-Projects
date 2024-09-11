from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import random


def recipient_fetch_script(username_input: str, password_input: str):
    #part 1: Logging in
    opts = webdriver.ChromeOptions()
    opts.add_experimental_option("detach", True)
    opts.add_argument("--disable-blink-features=AutomationControlled") 
    opts.add_experimental_option("excludeSwitches", ["enable-automation"]) 
    opts.add_experimental_option("useAutomationExtension", False) 
    driver = webdriver.Chrome(options=opts)

    linkedin_url = "https://www.linkedin.com/login?fromSignIn=true&trk=guest_homepage-basic_nav-header-signin"
    driver.get(linkedin_url)
    username = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH,'//*[@id="username"]')
                        ))
    cookies_yes =WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="artdeco-global-alert-container"]/div/section/div/div[2]/button[1]')))
    cookies_yes.click()

    password = driver.find_element(By.XPATH, '//*[@id="password"]')
    time.sleep(2)
    username.click()  # Focus the username field
    for character in username_input:
        username.send_keys(character)  # Add each character to the action chain
        print(character)
        time.sleep(random.uniform(0.7, 1.5))  # Random sleep to simulate human typing
    time.sleep(1)
    password.click()  # Focus the username field
    for character in password_input:
        password.send_keys(character)  # Add each character to the action chain
        print(character)
        time.sleep(random.uniform(0.3, 1))  # Random sleep to simulate human typing
    password.send_keys(Keys.ENTER)  # Send the Enter key
    print("Logged in successfully")

    #part 2:
    #Scrape all the names in messages, modify below to scrape month, and also change output
    messages_icon = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="global-nav"]/div/nav/ul/li[4]/a')))
    messages_icon.click()
    time.sleep(5)
    names_list = []
    fullnames = driver.find_elements(By.XPATH, '//div[contains(@class, "truncate")]')
    print("messages page loaded")
    for fullname in fullnames:
        names_list.append(fullname.text)
        print(fullname.text)
    section = driver.find_element(By.XPATH, '//*[@id="main"]/div/div[2]/div[1]/div[2]/ul')
    previous_height = driver.execute_script("return arguments[0].scrollHeight;", section)
    print(f"Scroll Height of the section: {previous_height}")
    while True:
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", section)
            sleep_time = random.randint(2, 5)
            time.sleep(sleep_time)
            # Extract the text content from each element
            fullnames = driver.find_elements(By.XPATH, '//div[contains(@class, "truncate")]')
            for fullname in fullnames:
                if fullname.text not in names_list:
                    names_list.append(fullname.text)
                    print(fullname.text)
            new_height = driver.execute_script("return arguments[0].scrollHeight;", section)
            if new_height == previous_height:
                break
            previous_height = new_height
    print(f"names collected:", names_list)
    driver.quit()
    return names_list
    