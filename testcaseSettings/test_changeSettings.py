import os, allure, pytest, datetime
from configurations.config import ReadConfig as config
from pageobjects.remote_test import *
from utils.readexcel import *
from pageobjects.login_logout import *
from pageobjects.Dashboard import *
from utils.updateexcelfile import *
from utils.library import *
from pageobjects.settings__dash import main_func_default_settings, main_func_change_settings

class Test_ChangeSettings_Driver:
    driver = None
    settings_runvalue_1 = Testrun_mode(value="Change Settings")
    if "LOADING".lower() == settings_runvalue_1[-1].strip().lower() or "Yes".lower() == settings_runvalue_1[-1].strip().lower():
        @pytest.mark.parametrize("environment,url,userid,password",fetch_enviroment())
        def test_changesettings(self,setup,environment, url, userid, password):
            driver = setup
            password = encrypte_decrypte(text=password)

            # Launch browser and Navigate to RantCell Application LoginPage
            with allure.step("Launch and navigating to RantCell Application LoginPage"):
                Navigate_to_loginPage(driver, url)

            # Login to RantCell Application
            with allure.step("Login to RantCell Application"):
                login(driver, userid, password)

            with allure.step("Change Settings"):
                main_func_change_settings(driver,environment,userid)

            # Logout from RantCell Application
            with allure.step("Logout to RantCell Application"):
                logout(driver)
