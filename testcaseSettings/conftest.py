import allure
import pytest
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from configurations.config import ReadConfig as config
from utils.library import *

active_threads = 0
threads = 0
Change_Settingsrunvalue = Testrun_mode(value="Change Settings")
default_Settingsrunvalue = Testrun_mode(value="Default Settings")

def pytest_sessionstart(session):
    if active_threads == 0:
        pass
@pytest.fixture(scope='function')
def setup(request):
    global driver
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    global active_threads
    active_threads += 1
    yield driver
    driver.quit()

def pytest_collection_modifyitems(config, items):
    config.collected_items_count = len(items)
    global threads
    threads = config.collected_items_count

def pytest_sessionfinish(session):
    item_count = threads
    if active_threads == item_count:
        if  "Yes".lower() == default_Settingsrunvalue[-1].strip().lower():
            updating_yes_to_run(remotevalue="RUNNED", types_of_test="Default Settings")
        elif  "FINISHED".lower() == default_Settingsrunvalue[-1].strip().lower():#--->testcase/conftest.py pytest_sessionfinish RUNNED to FINISHED
            updating_yes_to_run(remotevalue="Yes", types_of_test="Default Settings")

        if  "Yes".lower() == Change_Settingsrunvalue[-1].strip().lower():
            updating_yes_to_run(remotevalue="WAITING LOAD", types_of_test="Change Settings")
        elif "LOADING".lower() == Change_Settingsrunvalue[-1].strip().lower():#--->testcase/conftest.py pytest_sessionfinish WAITING LOAD to LOADING
            updating_yes_to_run(remotevalue="RUNNED", types_of_test="Change Settings")