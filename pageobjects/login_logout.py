from utils.library import *
from locators.locators import Login_Logout
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
def Navigate_to_loginPage(driver, url):
    try:
        assert launchbrowser(driver, url)
    except:
        pass
    try:
        WebDriverWait(driver, 90).until(EC.presence_of_element_located(Login_Logout.link_login1))
    except:
        pass
    try:
        assert clickec(driver,Login_Logout.link_login)
    except:
        pass

def login(driver, userid, password):
    try:
        try:
            WebDriverWait(driver, 90).until(EC.presence_of_element_located(Login_Logout.email))
        except:
            pass
        assert inputtext(driver, Login_Logout.textbox_username, userid)
        assert inputtext(driver, Login_Logout.textbox_password, password)
        assert click(driver, Login_Logout.button_login)
        dashboard_loading(driver)
        assert verifyelementispresent(driver, Login_Logout.dashboard)
        assert True

    except Exception as e:
        assert False
def dashboard_loading(driver):
    try:
        dashboard_elemnt = driver.find_elements(Login_Logout.dashboard[0], Login_Logout.dashboard[1])
        start_time = time.time()
        # Maximum time in seconds the loop should run (1 minute = 60 seconds)
        max_run_time = 120
        if len(dashboard_elemnt) == 0:
            with allure.step("Waiting for Dashboard to load"):
                allure.attach(driver.get_screenshot_as_png(), name=f"Waiting for Dashboard to load",attachment_type=allure.attachment_type.PNG)
                while time.time() - start_time < max_run_time:
                    dashboard_elemnt = driver.find_elements(Login_Logout.dashboard[0], Login_Logout.dashboard[1])
                    # Check if the condition is met
                    if len(dashboard_elemnt) != 0:
                        break
    except:
        pass
def logout(driver):
    try:
        assert click(driver, Login_Logout.dropdown_dropdown_toggle)
        time.sleep(1.2)
        assert click(driver, Login_Logout.link_logout)
        assert True

    except Exception as e:
        assert False
