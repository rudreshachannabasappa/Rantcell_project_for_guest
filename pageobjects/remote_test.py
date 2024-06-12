import concurrent.futures
import queue
import re
import sys
import time

import allure
import pytest
from selenium.webdriver.common.by import By
from pageobjects.Dashboard import *
from utils.library import *
from utils.readexcel import *
from locators.locators import *
import random
import string
def remote_test_(driver,device,campaign,campaigns_created,usercampaignsname,testgroup,tests,excelpath):
    Title = "REMOTE TEST"
    result_status = queue.Queue()
    test_complete = []
    test_Selected = []
    run_test_status_value = []
    test_Execution_status =[]
    campaigns_status = []
    runtest_runvalue = Testrun_mode(value="Remote Test")
    device = common(driver, device, campaign, campaigns_created, usercampaignsname, testgroup, tests, excelpath, runtest_runvalue, Title,test_complete, test_Selected, result_status, run_test_status_value, test_Execution_status, campaigns_status,"remote")
    return device

def common(driver,device,campaign,campaigns_created,usercampaignsname,testgroup,tests,excelpath,runvalue,Title,test_complete,test_Selected,result_status,run_test_status_value,test_Execution_status,campaigns_status,run):
    try:
        if "Yes".lower() == runvalue[-1].strip().lower():
            if "Multi Bparty Call Test" not in tests:
                try:
                    clickec(driver=driver,locators=remote_test.remotetest)
                    time.sleep(3)
                    if run == "remote":
                        device = remotetest_for_android_pro(driver,Title,device,campaign,usercampaignsname,testgroup,tests,result_status,test_complete,test_Selected,run_test_status_value,excelpath)
                    elif run == "schedule":
                        device = schedule_for_android_pro(driver, Title, device, campaign, usercampaignsname, testgroup,tests, result_status, test_complete, test_Selected,run_test_status_value, excelpath)
                    if test_complete == [True] and test_Selected == [True]:
                        click(driver=driver, locators=Login_Logout.dashboard_id)
                        test_name_in_table_view = (By.XPATH, f"//td[@id='loaderCamp']/following-sibling::td[normalize-space()='{usercampaignsname}']")
                        if run == "remote":
                            verifying_of_test_execution_for_runtest(driver,Title,device,campaign,usercampaignsname,test_name_in_table_view,run_test_status_value,result_status,test_Execution_status,campaigns_status)
                        elif run == "schedule":
                            verifying_of_test_execution_for_scheduletest(driver, Title, device, campaign,usercampaignsname, test_name_in_table_view,run_test_status_value, result_status,test_Execution_status, campaigns_status)
                except Exception as e:
                    pass
            elif "Multi Bparty Call Test" in tests:
                statement = "Multi Bparty Call Test is option is not present in 'run test'"
                with allure.step(statement):
                    status_df = status(Title, "Not to execute", "SKIPPED", statement)
                    result_status.put(status_df)
                    pass
        elif "No".lower() == runvalue[-1].strip().lower():
            statement = "You have selected Not to execute"
            with allure.step(statement):
                status_df = status(Title, "Not to execute", "SKIPPED", "You have selected No for execute")
                result_status.put(status_df)
                pass
        try:
            update_remote_test_result(result_status, excelpath)
        except Exception as e:
            pass
    finally:
        if "Multi Bparty Call Test" not in tests:
            if "Yes".lower() == runvalue[-1].strip().lower():
                try:
                    if ((test_complete == [] or test_complete == [False]) or (test_Selected == [] or test_Selected == [False])):
                        statement = "Test failed"
                        pytest.fail(statement)
                    if (test_Execution_status == [] or test_Execution_status == [False] or test_Execution_status == [None]) or (campaigns_status == [] or campaigns_status == [False]):
                        statement = "Test failed"
                        pytest.fail(statement)
                    campaigns_created.append(usercampaignsname)
                except Exception as e:
                    pass
        return device
def Schedule_test_(driver,device,campaign,campaigns_created,usercampaignsname,testgroup,tests,excelpath):
    Title = "SCHEDULE TEST"
    result_status = queue.Queue()
    test_complete = []
    test_Selected = []
    run_test_status_value = []
    test_Execution_status =[]
    campaigns_status = []
    scheduletest_runvalue = Testrun_mode(value="Schedule Test")
    device = common(driver, device, campaign, campaigns_created, usercampaignsname, testgroup, tests, excelpath, scheduletest_runvalue, Title,test_complete, test_Selected, result_status, run_test_status_value, test_Execution_status, campaigns_status, "schedule")
    return device
def Continuous_Test_(driver,device,campaign,campaigns_created,usercampaignsname,testgroup,continuous_campaigns,tests,excelpath):
    Title = "CONTINUOUS TEST"
    result_status = queue.Queue()
    test_complete = []
    test_Selected = []
    run_test_status_value = []
    test_Execution_status =[]
    campaigns_status = []
    usercampaignsname_list = []
    multi_bparty_call_test_flag = []
    continuoustest_runvalue = Testrun_mode(value="Continuous Test")
    try:
        if "Yes".lower() == continuoustest_runvalue[-1].strip().lower():
            try:
                clickec(driver=driver,locators=remote_test.remotetest)
                time.sleep(3)
                device = continuoustest_for_android_pro(driver,Title,device,campaign,usercampaignsname,testgroup,tests,result_status,test_complete,test_Selected,run_test_status_value,multi_bparty_call_test_flag,excelpath)
                if test_complete == [True] and test_Selected == [True]:
                    click(driver=driver, locators=Login_Logout.dashboard_id)
                    test_name_in_table_view = (By.XPATH,f"//td[@id='loaderCamp']/following-sibling::td[contains(.,'{device}')]/preceding-sibling::td[starts-with(normalize-space(),'{continuous_campaigns}')]")
                    verifying_of_test_execution_for_continuous(driver,Title,device,campaign,usercampaignsname,testgroup,continuous_campaigns,test_name_in_table_view,run_test_status_value,result_status,test_Execution_status,campaigns_status,usercampaignsname_list,multi_bparty_call_test_flag)
            except Exception as e:
                pass
        elif "No".lower() == continuoustest_runvalue[-1].strip().lower() and "Multi Bparty Call Test" in tests:
            statement = f"Multi Bparty Call Test is option is present in 'Continuous Test' for this campaign:-{campaign} give 'Yes' in 'CAMPAIGNS_TOTEST' sheet but give 'No' in 'TEST_RUN'sheet for 'Continuous Test'"
            with allure.step(statement):
                status_df = status(Title, "Not executed", "SKIPPED", statement)
                result_status.put(status_df)
                pass
        elif "No".lower() == continuoustest_runvalue[-1].strip().lower():
            statement = "You have selected Not to execute"
            with allure.step(statement):
                status_df = status(Title, "Not to execute", "SKIPPED", "You have selected No for execute")
                result_status.put(status_df)
                pass
        try:
            update_remote_test_result(result_status, excelpath)
        except Exception as e:
            pass
    finally:
        if "Yes".lower() == continuoustest_runvalue[-1].strip().lower():
            try:
                if ((test_complete == [] or test_complete == [False]) or (test_Selected == [] or test_Selected == [False])):
                    statement = "Test failed"
                    pytest.fail(statement)
                if (test_Execution_status == [] or test_Execution_status == [False] or test_Execution_status == [None]) or (campaigns_status == [] or campaigns_status == [False]):
                    statement = "Test failed"
                    pytest.fail(statement)
                usercampaignsname_list = list(set(usercampaignsname_list))
                for usercampaignsname2 in usercampaignsname_list:
                    campaigns_created.append(usercampaignsname2)
            except Exception as e:
                pass
        return device

def stop_pytest():
    sys.exit("Stopping pytest")
def verifying_of_test_execution_for_runtest(driver,Title,device,campaign,usercampaignsname,test_name_in_table_view,run_test_status_value,result_status,test_Execution_status,campaigns_status):
    try:
        with allure.step("Verification of the success or failed status from the application after clicking on start button in run test of remote test"):
            action = ActionChains(driver)
            try:
                table_view_refresh_element = driver.find_element(*remote_test.table_view_refresh[:2])
                action.move_to_element(table_view_refresh_element).perform()
            except Exception as e:
                pass
            execution_time = 0
            total_time = 331 #in min
            for i in range(0,total_time):
                try:
                    clickec(driver=driver,locators=remote_test.table_view_refresh)
                    WebDriverWait(driver,10).until(EC.visibility_of_element_located(test_name_in_table_view))
                    execution_time = total_time - i
                    break
                except Exception as e:
                    continue
            allure.attach(driver.get_screenshot_as_png(), name=f"verifying_of_test_execution",attachment_type=allure.attachment_type.PNG)
            try:
                Page_Down(driver=driver)
                test_name_siblings = WebDriverWait(driver,10).until(EC.visibility_of_element_located(test_name_in_table_view))
                campaigns_status.append(True)
                if run_test_status_value == [True]:
                    status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="PASSED", comments=f"Success")
                    result_status.put(status_df)
                    waiting_for_complete_or_Aborted_status_for_runtest(Title,device,campaign,usercampaignsname,result_status,execution_time,driver=driver,test_Execution_status=test_Execution_status)
                elif run_test_status_value == [False]:
                    status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="WARNING", comments=f"Remote Test config failed but campaign available")
                    result_status.put(status_df)
                    waiting_for_complete_or_Aborted_status_for_runtest(Title,device,campaign,usercampaignsname,result_status,execution_time,driver=driver,test_Execution_status=test_Execution_status)
            except Exception as e:
                campaigns_status.append(False)
                test_Execution_status.append(None)
                if run_test_status_value == [True]:
                    status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED", comments=f"Remote Test config successful but campaign not available")
                    result_status.put(status_df)
                elif run_test_status_value == [False]:
                    status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED", comments=f"	Remote Test config failed to run")
                    result_status.put(status_df)
                pass
    except Exception as e:
        pass
def waiting_for_complete_or_Aborted_status_for_runtest(Title,device,campaign,usercampaignsname,result_status,execution_time,driver,test_Execution_status):
    with allure.step("waiting for complete or Aborted status"):
        try:
            action = ActionChains(driver)
            for i in range(0, execution_time):
                try:
                    clickec(driver=driver, locators=remote_test.table_view_refresh)
                    test_execution = (By.XPATH,f"//td[@id='loaderCamp']/following-sibling::td[normalize-space()='{usercampaignsname}']/following-sibling::td[contains(.,'COMPLETED') or contains(.,'ABORTED')]")
                    WebDriverWait(driver,10).until(EC.visibility_of_element_located(test_execution))
                    break
                except Exception as e:
                    continue
            for i in range(0, 1):
                try:
                    clickec(driver=driver, locators=remote_test.table_view_refresh)
                    test_execution = (By.XPATH,f"//td[@id='loaderCamp']/following-sibling::td[normalize-space()='{usercampaignsname}']/following-sibling::td[contains(.,'COMPLETED') or contains(.,'ABORTED')]")
                    WebDriverWait(driver,10).until(EC.visibility_of_element_located(test_execution))
                    try:
                        test_execution_element = driver.find_element(*test_execution)
                        action.move_to_element(test_execution_element).perform()
                    except Exception as e:
                        pass
                    test_Execution_status.append(True)
                    status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="PASSED", comments=f"Test Campaigns execution status is Completed/Aborted/Uploaded")
                    result_status.put(status_df)
                    break
                except Exception as e:
                    test_Execution_status.append(False)
                    status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED", comments=f"Test Campaigns execution status is Executing")
                    result_status.put(status_df)
                    try:
                        test_executing = (By.XPATH,f"//td[@id='loaderCamp']/following-sibling::td[normalize-space()='{usercampaignsname}']/following-sibling::td[contains(.,'EXECUTING')]")
                        try:
                            test_executing_element = driver.find_element(*test_executing)
                            action.move_to_element(test_executing_element).perform()
                        except Exception as e:
                            pass
                        WebDriverWait(driver, 0.1).until(EC.invisibility_of_element_located(test_executing))
                    except Exception as e:
                        pass
            allure.attach(driver.get_screenshot_as_png(), name=f"waiting for complete or Aborted status",attachment_type=allure.attachment_type.PNG)
        except Exception as e:
            pass
def Updating_automation_data_to_excel(worksheet,dataframe):
    try:
        # Find the last used row in the sheet
        last_row = worksheet.max_row
        # Append DataFrame data to the worksheet
        for index, row in dataframe.iterrows():
            worksheet.append(row.tolist())
    except Exception as e:
        pass
def update_automation_data(automation_data_dict,automation_data_execel_path,Sheet):
    try:
        dataframe_automation_data = []
        automation_data_df = "None"
        df_automation_data = pd.DataFrame(automation_data_dict)
        dataframe_automation_data.append(df_automation_data)
        if len(dataframe_automation_data) != 0:
            automation_data_df = pd.concat(dataframe_automation_data, ignore_index=True)
        workbook = openpyxl.load_workbook(automation_data_execel_path)
        worksheet_componentstatus = workbook[Sheet]
        if len(dataframe_automation_data) != 0:
            Updating_automation_data_to_excel(worksheet=worksheet_componentstatus, dataframe=automation_data_df)
        workbook.save(automation_data_execel_path)
        workbook.close()
    except Exception as e:
        pass
def update_remote_test_result(result_status,excelpath):
    try:
        dataframe_status = []
        combined_status_df = "None"
        while not result_status.empty():
            updatecomponentstatus2 = result_status.get()
            df_status = pd.DataFrame(updatecomponentstatus2)
            dataframe_status.append(df_status)
        if len(dataframe_status) != 0:
            combined_status_df = pd.concat(dataframe_status, ignore_index=True)
        workbook = openpyxl.load_workbook(excelpath)
        worksheet_componentstatus = workbook["COMPONENTSTATUS"]
        if len(dataframe_status) != 0:
            update_component_status_openpyxl(worksheet=worksheet_componentstatus, dataframe=combined_status_df)
        workbook.save(excelpath)
        workbook.close()
    except Exception as e:
        pass
def remotetest_for_android_pro(driver,Title,device,campaign,usercampaignsname,testgroup,tests,result_status,test_complete,test_Selected,run_test_status_value,excelpath):
    try:
        with allure.step("remotetest for android pro"):
            flag_test_group = []
            alert_text = None
            check_android_pro_is_active_in_remotetest(driver)
            verify_test_group_is_present(driver,Title,device,campaign,usercampaignsname,result_status,testgroup,flag_test_group)
            if flag_test_group == [True]:
                remotetest_runvalue = Testrun_mode(value="Remote Test")
                scheduletest_runvalue = Testrun_mode(value="Schedule test")
                continuoustest_runvalue = Testrun_mode(value="Continuous Test")
                if "Yes".lower() == remotetest_runvalue[-1].strip().lower() and "Yes".lower() != scheduletest_runvalue[-1].strip().lower() and "Yes".lower() != continuoustest_runvalue[-1].strip().lower():
                    device_button_dropdown = click_on_test_group_button_to_open_dropdown(driver, testgroup)
                    click_on_the_check_devices(driver=driver,driver1=device_button_dropdown)
                try:
                    alert_text = alert_accept(driver=driver)
                except Exception as e:
                    pass
                if alert_text !=None:
                    with allure.step(f"{alert_text}"):
                        status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED",comments=f"{alert_text}")
                        result_status.put(status_df)
                elif alert_text == None:
                    with allure.step(f"no alert found, device is registered"):
                        status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}", status="PASSED", comments=f"no alert found, device is registered")
                        result_status.put(status_df)
                    flag_status_value = []
                    remotetest_runvalue = Testrun_mode(value="Remote Test")
                    scheduletest_runvalue = Testrun_mode(value="Schedule test")
                    continuoustest_runvalue = Testrun_mode(value="Continuous Test")
                    if "Yes".lower() == remotetest_runvalue[-1].strip().lower() and "Yes".lower() != scheduletest_runvalue[-1].strip().lower() and "Yes".lower() != continuoustest_runvalue[-1].strip().lower():
                        device = check_device_status(driver,Title,device,campaign,usercampaignsname,flag_status_value,result_status)
                    device_button_dropdown = click_on_test_group_button_to_open_dropdown(driver, testgroup)
                    click_on_the_run_test(driver=driver,driver1=device_button_dropdown)
                    waiting_for_run_test_tab_for_loading(driver)
                    run_test_form(driver,Title,usercampaignsname,tests,test_complete,test_Selected,result_status,excelpath)
                    if test_complete == [True] and test_Selected == [True]:
                        statement = "Successfully entered test data for a particular type of test and clicked on the start button."
                        with allure.step(statement):
                            allure.attach(driver.get_screenshot_as_png(), name=f"{statement}",attachment_type=allure.attachment_type.PNG)
                            click_on_start_button_of_run_test(driver=driver)
                            status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="PASSED", comments=statement)
                            result_status.put(status_df)
                            check_status_of_test(driver, Title, device, campaign, usercampaignsname,run_test_status_value,result_status)
                    elif (test_complete == [] or test_complete == [False]) and (test_Selected == [False] or test_Selected == []):
                        statement = "Test data was not entered successfully for a particular type of test,so clicked on the close button."
                        with allure.step(statement):
                            allure.attach(driver.get_screenshot_as_png(), name=f"{statement}",attachment_type=allure.attachment_type.PNG)
                            status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED", comments=statement)
                            result_status.put(status_df)
                            click_on_close_button_of_run_test(driver)
                        time.sleep(0.1)
    except Exception as e:
        pass
    finally:
        return device
def update_device_name_for_test_data(driver,device):
    try:
        devices_element = driver.find_element(*remote_test.device_name)
        device = devices_element.get_attribute("innerText")
        return str(device).strip()
    except Exception as e:
        pass
def check_android_pro_is_active_in_remotetest(driver):
    try:
        with allure.step("checking android pro button is selected and is active in remotetest"):
            # Find the active and inactive elements based on the class
            inactive_element = driver.find_element(*remote_test.android_pro_is_inactive)
            # Check if the element is not active
            if inactive_element:
                # Click on the inactive element to make it active
                inactive_element.click()
            allure.attach(driver.get_screenshot_as_png(), name=f"android pro tab open",attachment_type=allure.attachment_type.PNG)
    except Exception as e:
        pass
def verify_test_group_is_present(driver,Title,device,campaign,usercampaignsname,result_status,testgroup,flag_test_group):
    with allure.step("verifying test group is present"):
        try:
            WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,f"//p[normalize-space()='{testgroup}']/ancestor::div[@class='deviceCtldiv']/following-sibling::div//div//button")))
            flag_test_group.append(True)
        except Exception as e:
            flag_test_group.append(False)
            with allure.step(f"Check the test group name is present in remote test as user input from test data"):
                status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED", comments=f"Check the test group name is present in remote test as user input from test data")
                result_status.put(status_df)
            pass
        allure.attach(driver.get_screenshot_as_png(), name=f"verifying test group is present",attachment_type=allure.attachment_type.PNG)

def click_on_test_group_button_to_open_dropdown(driver,testgroup):
    try:
        with allure.step("click on the test group device button to open the dropdown"):
            device_button_dropdown_path =(By.XPATH,f"//p[normalize-space()='{testgroup}']/ancestor::div[@class='deviceCtldiv']/following-sibling::div//div//button","device_button_dropdown")
            device_button_dropdown = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,f"//p[normalize-space()='{testgroup}']/ancestor::div[@class='deviceCtldiv']/following-sibling::div//div//button")))
            element = driver.find_element(By.XPATH,f"//p[normalize-space()='{testgroup}']/ancestor::div[@class='deviceCtldiv']/following-sibling::div//div//button")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            clickec(driver=driver,locators=device_button_dropdown_path)
            allure.attach(driver.get_screenshot_as_png(), name=f"click on the test group device button to open the dropdown",attachment_type=allure.attachment_type.PNG)
            return device_button_dropdown
    except Exception as e:
        pass
def click_on_the_check_devices(driver,driver1):
    try:
        with allure.step("click on the check devices"):
            check_devices_element = driver1.find_element(*remote_test.check_devices)
            check_devices_element.click()
            with allure.step("Waiting for loading check device"):
                try:
                    WebDriverWait(driver, 10).until(EC.visibility_of_element_located(remote_test.check_device_tab))
                except Exception as e:
                    pass
            allure.attach(driver.get_screenshot_as_png(), name=f"click on the check devices.",attachment_type=allure.attachment_type.PNG)
    except Exception as e:
        pass
def wait_for_timer_to_reach_zero_in_check_device_popup(driver):
    try:
        with allure.step("waiting for the timer to reach zero in check device popup"):
            timer_element = driver.find_element(*remote_test.timer_xpath)
            # Extract the countdown attribute value
            countdown_value = timer_element.get_attribute("countdown")
            WebDriverWait(driver, int(countdown_value)+2).until(EC.visibility_of_element_located(remote_test.online_or_offline_status))
            allure.attach(driver.get_screenshot_as_png(), name=f"Timer reached 0 seconds.",attachment_type=allure.attachment_type.PNG)
    except Exception as e:
        pass
def wait_for_status_of_run_test_popup(driver):
    try:
        with allure.step("waiting for the status of run test popup"):
            countdown_value = 60
            WebDriverWait(driver, int(countdown_value)+2).until(EC.visibility_of_element_located((remote_test.run_test_start_status[:2])))
            allure.attach(driver.get_screenshot_as_png(), name=f"waiting for the status of run test popup.",attachment_type=allure.attachment_type.PNG)
    except Exception as e:
        print(f"Timer did not reach 0 within the specified time. {e}")
def check_status_of_test(driver,Title,device,campaign,usercampaignsname,run_test_status_value,result_status):
    try:
        with allure.step("checking the status of test"):
            try:
                WebDriverWait(driver, 10).until(EC.visibility_of_element_located(remote_test.run_test_start_status_popup))
            except Exception as e:
                pass
            wait_for_status_of_run_test_popup(driver)
            status_of_runtests = driver.find_elements(*remote_test.run_test_start_status[:2])
            for status_of_runtest in status_of_runtests:
                status_value = status_of_runtest.text
                if status_value.lower().replace(" ", "") == "Failed".lower():
                    with allure.step("Test execution didnt started may be due to device is offline"):
                        run_test_status_value.append(False)
                        allure.attach(driver.get_screenshot_as_png(), name=f"offline",attachment_type=allure.attachment_type.PNG)
                        status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED", comments="Test execution didnt started may be due to device is offline")
                        result_status.put(status_df)
                elif status_value.lower().replace(" ", "") == "Success".lower():
                    with allure.step("Test execution didnt started may be due to device is offline"):
                        run_test_status_value.append(True)
                        allure.attach(driver.get_screenshot_as_png(), name=f"offline",attachment_type=allure.attachment_type.PNG)
                        status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="PASSED",comments="Test execution started and device is online")
                        result_status.put(status_df)
            clickec(driver=driver,locators=remote_test.run_test_start_statusclose_btn)
    except Exception as e:
        pass
def check_device_status(driver,Title,device,campaign,usercampaignsname,flag_status_value,result_status):
    try:
        with allure.step("check whether the device is online/offline "):
            wait_for_timer_to_reach_zero_in_check_device_popup(driver)
            device = update_device_name_for_test_data(driver, device)
            status_of_devices = driver.find_elements(*remote_test.status_of_devices_Offline)
            offline_flag = False
            for status_of_device in status_of_devices:
                offline_flag = True
                status_value = status_of_device.text
                if status_value.lower().replace(" ", "") == "Offline".lower():
                    with allure.step("device is offline"):
                        a = "device is offline"
                        flag_status_value.append(False)
                        allure.attach(driver.get_screenshot_as_png(), name=f"offline",attachment_type=allure.attachment_type.PNG)
                        status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}", status="FAILED", comments="device is offline")
                        result_status.put(status_df)
            if offline_flag == False:
                status_of_devices = driver.find_elements(*remote_test.status_of_devices_Online)
                for status_of_device in status_of_devices:
                    status_value = status_of_device.text
                    if status_value.lower().replace(" ","") == 'Online'.lower():
                        with allure.step("device is Online"):
                            flag_status_value.append(True)
                            status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}", status="PASSED", comments=f"device is Online")
                            result_status.put(status_df)
                            allure.attach(driver.get_screenshot_as_png(), name=f"online",attachment_type=allure.attachment_type.PNG)
                            break
        click_on_close_button_of_check_device(driver)
    except Exception as e:
        try:
            click_on_close_button_of_check_device(driver)
        except Exception as e:
            pass
        pass
    finally:
        return device
def click_on_the_run_test(driver,driver1):
    try:
        with allure.step("click on run test option from the dropdown"):
            Run_Test_element = driver1.find_element(*remote_test.Run_Test)
            Run_Test_element.click()
            allure.attach(driver.get_screenshot_as_png(), name=f"click on  run test option from the dropdown.",attachment_type=allure.attachment_type.PNG)
    except Exception as e:
        pass
def click_on_close_button_of_check_device(driver):
    try:
        with allure.step("click on close button of check device"):
            clickec(driver=driver,locators=remote_test.closebtn_ofcheckdevice)
            allure.attach(driver.get_screenshot_as_png(), name=f"click on close button of check device.",attachment_type=allure.attachment_type.PNG)
    except Exception as e:
        pass
def click_on_close_button_of_run_test(driver):
    try:
        with allure.step("click on the close button of run test"):
            click(driver=driver,locators=remote_test.closebtn_ofrun_test)
            time.sleep(0.1)
            allure.attach(driver.get_screenshot_as_png(), name=f"click on the close button of run test.",attachment_type=allure.attachment_type.PNG)
    except Exception as e:
        pass

def click_on_start_button_of_run_test(driver):
    try:
        with allure.step("click on start button of run test"):
            start_button = driver.find_element(*remote_test.startbtn_ofrun_test[:2])
            if start_button.is_enabled():
                click(driver=driver,locators=remote_test.startbtn_ofrun_test)
                time.sleep(0.1)
            allure.attach(driver.get_screenshot_as_png(), name=f"click on start button of run test.",attachment_type=allure.attachment_type.PNG)
    except Exception as e:
        pass

def waiting_for_run_test_tab_for_loading(driver):
    try:
        with allure.step("waiting for run test tab for loading"):
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located(remote_test.run_test_tab))
            allure.attach(driver.get_screenshot_as_png(), name=f"waiting for run test tab for loading.",attachment_type=allure.attachment_type.PNG)
    except Exception as e:
        pass

def ping_test(driver,Title,test_data,result_status,type_of_test):
    try:
        with allure.step("Ping test"):
            if type_of_test == "runtest":
                clickec(driver=driver, locators=remote_test.ping_test_checkbox)
            elif type_of_test == "continuoustest":
                clickec(driver=driver, locators=continuous.ping_checkbox)
            elif type_of_test == "scheduletest":
                clickec(driver=driver, locators=schedule_test.ping_test_checkbox)
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located(remote_test.ping_test_form))
            Ping_test = None
            ping_test_form_fields_names = driver.find_elements(*remote_test.ping_test_form)
            Ping_test = True
            for ping_test_form_fields_name in ping_test_form_fields_names:
                try:
                    field_name = str(ping_test_form_fields_name.text).lower().replace(" ","")
                    if re.search("Host".lower(),field_name,re.IGNORECASE):
                        try:
                            ping_data = test_data["Host"]
                        except Exception as e:
                            statement = f"Check parameter is present in remote test sheet test data excel"
                            with allure.step(statement):
                                status_df = status(Title=Title, component=f"Ping_test:-'Host'", status="FAILED", comments=statement)
                                result_status.put(status_df)
                                raise e
                        inputtext(driver=ping_test_form_fields_name,locators=remote_test.host_textbox,value=ping_data)
                    elif re.search("Packet Size".lower().replace(" ",""),field_name,re.IGNORECASE):
                        try:
                            ping_data = test_data["Packet Size"]
                        except Exception as e:
                            statement = f"Check parameter is present in remote_test sheet test data excel"
                            with allure.step(statement):
                                status_df = status(Title=Title, component=f"Ping_test:-'Packet Size'", status="FAILED",comments=statement)
                                result_status.put(status_df)
                                raise e
                        if is_numeric(ping_data):
                            if 32 <= int(ping_data) <= 65500:
                                inputtext(driver=ping_test_form_fields_name,locators=remote_test.packetsize_textbox,value=ping_data)
                            else:
                                statement = f"Check the value given within the range"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Ping_test:-'Packet Size'", status="FAILED",comments=statement)
                                    result_status.put(status_df)
                                    e = Exception
                                    raise e
                        else:
                            statement = f"Check the entered value is numeric"
                            with allure.step(statement):
                                status_df = status(Title=Title, component=f"Ping_test:-'Packet Size'", status="FAILED",comments=statement)
                                result_status.put(status_df)
                                e = Exception
                                raise e
                except Exception as e:
                    Ping_test = False
                    continue
            allure.attach(driver.get_screenshot_as_png(), name=f"Ping_test.",attachment_type=allure.attachment_type.PNG)
            if Ping_test == True:
                okbtn = driver.find_element(*remote_test.pingtest_okbtn[:2])
                if okbtn.is_enabled():
                    clickec(driver=driver,locators=remote_test.pingtest_okbtn)
                elif not okbtn.is_enabled():
                    clickec(driver=driver, locators=remote_test.pingtest_closebtn)
                    e = Exception
                    raise e
            elif Ping_test == False or Ping_test == None:
                clickec(driver=driver, locators=remote_test.pingtest_closebtn)
                e = Exception
                raise e
    except Exception as e:
        raise e

def run_test_form(driver,Title,usercampaignsname,tests,test_complete,test_Selected,result_status,excelpath):
    try:
        with allure.step("run test form"):
            inputtext(driver=driver,locators=remote_test.test_name,value=f"{usercampaignsname}")
            inputtext(driver=driver,locators=remote_test.iteration_textbox,value="1")
            inputtext(driver=driver, locators=remote_test.delays_bw_tests, value="5")
            allure.attach(driver.get_screenshot_as_png(), name=f"run_test_form.",attachment_type=allure.attachment_type.PNG)
            try:
                df_remote_test = pd.read_excel(config.test_data_path,sheet_name="Remote_Test")
            except Exception as e:
                with allure.step(f"Check {config.test_data_path}"):
                    print(f"Check {config.test_data_path}")
                    assert False
            txt = []
            if tests.__len__() == 0:
                statement = f"{Title}  --  Nothing is marked 'Yes' in {str(config.test_data_path)}"
                with allure.step(f"Nothing is marked 'Yes' in {str(config.test_data_path)} for '{Title}'"):
                    updatecomponentstatus(Title, "", "FAILED", f"Nothing marked in {str(config.test_data_path)}",excelpath)
                    e = Exception
                    raise e
            else:
                test_complete.append(True)
                for test in tests:
                    try:
                        if test.lower().replace(" ", "") == "tcp-iperftest".lower().replace(" ", "") or test.lower().replace(" ", "") == "udp-iperftest".lower().replace(" ", ""):
                             testmodified = test.replace("TCP-", "").replace("UDP-", "")
                             remote_test_datas = df_remote_test[df_remote_test['Test_Type'].str.contains(testmodified,case=False, na=False)].groupby('Test_Type').apply(lambda x: x[['Parameter', 'Value']].to_dict(orient='records')).tolist()
                        else:
                            remote_test_datas = df_remote_test[df_remote_test['Test_Type'].str.contains(test, case=False, na=False)].groupby('Test_Type').apply(lambda x: x[['Parameter', 'Value']].to_dict(orient='records')).tolist()
                        remote_test_dict = {str(record['Parameter']).strip(): record['Value'] for record in remote_test_datas[0]}
                        test = test.strip().lower().replace(" ", "")
                        test_functions = {
                            "pingtest": ping_test,
                            "calltest": call_test,
                            "smstest": sms_test,
                            "speed_test": speed_test,
                            "httpspeedtest": http_speed_test,
                            "webtest": web_test,
                            "streamtest": stream_test,
                            "tcp-iperftest": iperf_test,
                            "udp-iperftest": iperf_test
                        }
                        test_name = test.lower().replace(" ", "")
                        for name, func in test_functions.items():
                            if re.fullmatch(name, test_name):
                                if name in ["webtest", "streamtest"]:
                                    func(driver=driver, test_data=remote_test_dict, type_of_test="runtest")
                                elif name in ["tcp-iperftest", "udp-iperftest"]:
                                    func(driver=driver, Title=Title, test_data=remote_test_dict,result_status=result_status, type_of_test="runtest",test_name=test_name)
                                else:
                                    func(driver=driver, Title=Title, test_data=remote_test_dict,result_status=result_status, type_of_test="runtest")
                                break
                    except Exception as e:
                        try:
                            test_complete.remove(True)
                        except Exception as e:
                            pass
                        test_complete.append(False)
                        test_complete = list(set(test_complete))
                        continue
                test_Selected.append(True)
                for test in tests:
                    try:
                        test = test.strip().lower().replace(" ", "")
                        test_checkboxes = {
                            "pingtest": remote_test.ping_test_checkbox,
                            "calltest": remote_test.call_test_checkbox,
                            "smstest": remote_test.sms_test_checkbox,
                            "speed_test": remote_test.speed_test_checkbox,
                            "httpspeedtest": remote_test.http_speed_test_checkbox,
                            "webtest": remote_test.webtest_checkbox,
                            "streamtest": remote_test.stream_checkbox,
                            "tcp-iperftest": remote_test.iperf_testcheckbox,
                            "udp-iperftest": remote_test.iperf_testcheckbox
                        }
                        test_name = test.lower().replace(" ", "")
                        for name, checkbox in test_checkboxes.items():
                            if re.fullmatch(name, test_name):
                                typeoftest_is_selected(driver, checkbox)
                                break
                    except Exception as e:
                        try:
                            test_Selected.remove(True)
                        except Exception as e:
                            pass
                        test_Selected.append(False)
                        test_Selected = list(set(test_complete))
                        continue
    except Exception as e:
        pass
def typeoftest_is_selected(driver,locator):
    checkbox_input = driver.find_element(*locator[:2])
    try:
        if checkbox_input.is_selected():
            pass
        elif not checkbox_input.is_selected():
            raise
    except Exception as e:
        if not checkbox_input.is_selected():
            raise e

def call_test(driver,Title,test_data,result_status,type_of_test):
    try:
        with allure.step("Call test"):
            if type_of_test == "runtest":
                clickec(driver=driver, locators=remote_test.call_test_checkbox)
            elif type_of_test == "continuoustest":
                clickec(driver=driver, locators=continuous.call_checkbox)
            elif type_of_test == "scheduletest":
                clickec(driver=driver, locators=schedule_test.call_test_checkbox)
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located(remote_test.call_test_form))
            Call_test = None
            call_test_form_fields_names = driver.find_elements(*remote_test.call_test_form)
            Call_test = True
            for call_test_form_fields_name in call_test_form_fields_names:
                try:
                    field_name = str(call_test_form_fields_name.text).lower().replace(" ","")
                    if re.search("B Party Phone Number".lower().replace(" ",""),field_name,re.IGNORECASE):
                        try:
                            call_data = test_data["B Party Phone Number"]
                        except Exception as e:
                            statement = f"Check parameter is present in remote test sheet test data excel"
                            with allure.step(statement):
                                status_df = status(Title=Title, component=f"Call_test:-'B Party Phone Number'", status="FAILED",comments=statement)
                                result_status.put(status_df)
                                raise e
                        inputtext(driver=call_test_form_fields_name,locators=remote_test.Call_B_Party_Phone_Number_textbox,value=call_data)
                    elif re.search("Call Duration".lower().replace(" ",""),field_name,re.IGNORECASE):
                        try:
                            call_data = test_data["Call Duration"]
                        except Exception as e:
                            statement = f"Check parameter is present in remote_test sheet test data excel"
                            with allure.step(statement):
                                status_df = status(Title=Title, component=f"Call_test:-'Call Duration'", status="FAILED",comments=statement)
                                result_status.put(status_df)
                                raise e
                        if is_numeric(call_data):
                            if 1 <= int(call_data) <= 5400:
                                inputtext(driver=call_test_form_fields_name, locators=remote_test.Call_Duration_textbox,value=call_data)
                            else:
                                statement = f"Check the value given within the range"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Call_test:-'Call Duration'", status="FAILED",comments=statement)
                                    result_status.put(status_df)
                                    e = Exception
                                    raise e
                        else:
                            statement = f"Check the entered value is numeric"
                            with allure.step(statement):
                                status_df = status(Title=Title, component=f"Call_test:-'Call Duration'", status="FAILED",comments=statement)
                                result_status.put(status_df)
                                e = Exception
                                raise e
                except Exception as e:
                    Call_test = False
                    continue
            allure.attach(driver.get_screenshot_as_png(), name=f"Call_test.",attachment_type=allure.attachment_type.PNG)
            if Call_test == True:
                okbtn = driver.find_element(*remote_test.calltest_okbtn[:2])
                if okbtn.is_enabled():
                    clickec(driver=driver,locators=remote_test.calltest_okbtn)
                elif not okbtn.is_enabled():
                    clickec(driver=driver, locators=remote_test.calltest_closebtn)
                    e = Exception
                    raise e
            elif Call_test == False or Call_test == None:
                clickec(driver=driver, locators=remote_test.calltest_closebtn)
                e = Exception
                raise e
    except Exception as e:
        raise e
def sms_test(driver,Title,test_data,result_status,type_of_test):
    try:
        with allure.step("Sms test"):
            if type_of_test == "runtest":
                clickec(driver=driver, locators=remote_test.sms_test_checkbox)
            elif type_of_test == "continuoustest":
                clickec(driver=driver, locators=continuous.sms_checkbox)
            elif type_of_test == "scheduletest":
                clickec(driver=driver, locators=schedule_test.sms_test_checkbox)

            WebDriverWait(driver, 10).until(EC.visibility_of_element_located(remote_test.sms_test_form))
            Sms_test = None
            sms_test_form_fields_names = driver.find_elements(*remote_test.sms_test_form)
            Sms_test = True
            for sms_test_form_fields_name in sms_test_form_fields_names:
                try:
                    field_name = str(sms_test_form_fields_name.text).lower().replace(" ","")
                    if re.search("B Party Phone Number".lower().replace(" ",""),field_name,re.IGNORECASE):
                        try:
                            sms_data = test_data["B Party Phone Number"]
                        except Exception as e:
                            statement = f"Check parameter is present in remote test sheet test data excel"
                            with allure.step(statement):
                                status_df = status(Title=Title, component=f"Sms_test:-'B Party Phone Number'", status="FAILED",comments=statement)
                                result_status.put(status_df)
                                raise e
                        inputtext(driver=sms_test_form_fields_name,locators=remote_test.sms_B_Party_Phone_Number_textbox,value=sms_data)
                    elif re.search("Wait Duration".lower().replace(" ",""),field_name,re.IGNORECASE):
                        try:
                            sms_data = test_data["Wait Duration"]
                        except Exception as e:
                            statement = f"Check parameter is present in remote test sheet test data excel"
                            with allure.step(statement):
                                status_df = status(Title=Title, component=f"Sms_test:-'Wait Duration'", status="FAILED",comments=statement)
                                result_status.put(status_df)
                                raise e
                        if is_numeric(sms_data):
                            if 30 <= int(sms_data) <= 180:
                                inputtext(driver=sms_test_form_fields_name,locators=remote_test.sms_Wait_Duration_textbox,value=sms_data)
                            else:
                                statement = f"Check the value given within the range"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Sms_test:-'Wait Duration'", status="FAILED",comments=statement)
                                    result_status.put(status_df)
                                    e = Exception
                                    raise e
                        else:
                            statement = f"Check the entered value is numeric"
                            with allure.step(statement):
                                status_df = status(Title=Title, component=f"Sms_test:-'Wait Duration'", status="FAILED",comments=statement)
                                result_status.put(status_df)
                                e = Exception
                                raise e
                except Exception as e:
                    Sms_test = False
                    continue
            allure.attach(driver.get_screenshot_as_png(), name=f"Sms_test.",attachment_type=allure.attachment_type.PNG)
            if Sms_test == True:
                okbtn = driver.find_element(*remote_test.smstest_okbtn[:2])
                if okbtn.is_enabled():
                    clickec(driver=driver,locators=remote_test.smstest_okbtn)
                elif not okbtn.is_enabled():
                    clickec(driver=driver, locators=remote_test.smstest_closebtn)
                    e = Exception
                    raise e
            elif Sms_test == False or Sms_test == None:
                clickec(driver=driver, locators=remote_test.smstest_closebtn)
                e = Exception
                raise e
    except Exception as e:
        raise e
def speed_test(driver,Title,test_data,result_status,type_of_test):
    try:
        with allure.step("speed test"):
            if type_of_test == "runtest":
                clickec(driver=driver, locators=remote_test.speed_test_checkbox)
            elif type_of_test == "continuoustest":
                clickec(driver=driver, locators=continuous.speed_checkbox)
            elif type_of_test == "scheduletest":
                clickec(driver=driver, locators=schedule_test.speed_test_checkbox)

            WebDriverWait(driver, 10).until(EC.visibility_of_element_located(remote_test.speed_test_form))
            Speed_test = None
            speed_test_form_fields_names = driver.find_elements(*remote_test.speed_test_form)
            Speed_test = True
            try:
                Use_Default_Server_speed_data = test_data["Use Default Server"]
            except Exception as e:
                statement = f"Check parameter is present in remote test sheet test data excel"
                with allure.step(statement):
                    status_df = status(Title=Title, component=f"Speed_test:-'Use Default Server'", status="FAILED", comments=statement)
                    result_status.put(status_df)
            try:
                Enable_Upload_Test_speed_data = test_data["Enable Upload Test"]
            except Exception as e:
                statement = f"Check parameter is present in remote test sheet test data excel"
                with allure.step(statement):
                    status_df = status(Title=Title, component=f"Speed_test:-'Enable Upload Test'", status="FAILED",comments=statement)
                    result_status.put(status_df)
            try:
                Enable_FTP_Stop_Timer_speed_data = test_data["Enable FTP Stop Timer"]
            except Exception as e:
                statement = f"Check parameter is present in remote test sheet test data excel"
                with allure.step(statement):
                    status_df = status(Title=Title, component=f"Speed_test:-'Enable FTP Stop Timer'", status="FAILED",comments=statement)
                    result_status.put(status_df)
            for speed_test_form_fields_name in speed_test_form_fields_names:
                try:
                    field_name = str(speed_test_form_fields_name.text).lower().replace(" ","")
                    if re.search("Parallel Connections".lower().replace(" ",""),field_name,re.IGNORECASE):
                        pass
                    elif re.search("Use Default Server".lower().replace(" ",""),field_name,re.IGNORECASE):
                        checkbox_input = speed_test_form_fields_name.find_element(*remote_test.speed_Use_Default_Server_checkbox[:2])
                        if (Use_Default_Server_speed_data == "No" and checkbox_input.is_selected()) or (Use_Default_Server_speed_data == "Yes" and not checkbox_input.is_selected()):
                            clickec(driver=speed_test_form_fields_name,locators=remote_test.speed_Use_Default_Server_checkbox)

                    elif re.search("Select Download Test File Size".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Use_Default_Server_speed_data == "Yes" and Enable_FTP_Stop_Timer_speed_data == "No":
                            try:
                                speed_data = test_data["Select Download Test File Size"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote_test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Speed_test:-'Select Download Test File Size'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    raise e
                            Select_Download_Test_File_Size_option = (By.XPATH,f"//select[@id='downloadfilesize']//option[normalize-space()={speed_data}]","Select_Download_Test_File_Size_option")
                            clickec(driver=driver,locators=remote_test.Select_Download_Test_File_Size_dropdown)
                            clickec(driver=driver,locators=Select_Download_Test_File_Size_option)
                    elif re.search("FTP Server".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Use_Default_Server_speed_data == "No":
                            try:
                                speed_data = test_data["FTP Server"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote_test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Speed_test:-'FTP Server'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    raise e
                            inputtext(driver=speed_test_form_fields_name,locators=remote_test.FTP_Server_textbox,value=speed_data)
                    elif re.search("UserName".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Use_Default_Server_speed_data == "No":
                            try:
                                speed_data = test_data["UserName"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Speed_test:-'UserName'", status="FAILED",comments=statement)
                                    result_status.put(status_df)
                                    raise e
                            inputtext(driver=speed_test_form_fields_name,locators=remote_test.UserName_textbox,value=speed_data)
                    elif re.search("Password".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Use_Default_Server_speed_data == "No":
                            try:
                                speed_data = test_data["Select Download Test File Size"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote_test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Speed_test:-'Password'", status="FAILED",comments=statement)
                                    result_status.put(status_df)
                                    raise e
                            inputtext(driver=speed_test_form_fields_name,locators=remote_test.Password_textbox,value=speed_data)
                    elif re.search("Download File Name".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Use_Default_Server_speed_data == "No":
                            try:
                                speed_data = test_data["Download File Name"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Speed_test:-'Download File Name'", status="FAILED",comments=statement)
                                    result_status.put(status_df)
                                    raise e
                        inputtext(driver=speed_test_form_fields_name,locators=remote_test.Download_File_Name_textbox,value=speed_data)
                    elif re.search("Enable Upload Test".lower().replace(" ",""),field_name,re.IGNORECASE):
                        checkbox_input = speed_test_form_fields_name.find_element(*remote_test.Enable_Upload_Test_checkbox[:2])
                        if (Enable_Upload_Test_speed_data == "Yes" and not checkbox_input.is_selected()) or (Enable_Upload_Test_speed_data == "No" and checkbox_input.is_selected()):
                            clickec(driver=speed_test_form_fields_name, locators=remote_test.Enable_Upload_Test_checkbox)

                    elif re.search("File Size".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Enable_Upload_Test_speed_data == "Yes" and Enable_FTP_Stop_Timer_speed_data == "No":
                            try:
                                speed_data = test_data["File Size"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Speed_test:-'File Size'", status="FAILED",comments=statement)
                                    result_status.put(status_df)
                                    raise e
                            # input_field = speed_test_form_fields_name.find_element(*remote_test.speed_File_Size_textbox[:2])
                            # min_value = input_field.get_attribute("min")
                            # max_value = input_field.get_attribute("max")
                            if is_numeric(speed_data):
                                if 10 <= int(speed_data) <= 9999:
                                    inputtext(driver=speed_test_form_fields_name,locators=remote_test.speed_File_Size_textbox,value=speed_data)
                                else:
                                    statement = f"Check the value given within the range"
                                    with allure.step(statement):
                                        status_df = status(Title=Title, component=f"Speed_test:-'File Size'", status="FAILED",comments=statement)
                                        result_status.put(status_df)
                                        e = Exception
                                        raise e
                            else:
                                statement = f"Check the entered value is numeric"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Speed_test:-'File Size'", status="FAILED",comments=statement)
                                    result_status.put(status_df)
                                    e = Exception
                                    raise e
                                # speed_data = max(int(min_value), min(int(max_value), int(speed_data)))
                    elif re.search("Enable FTP Stop Timer".lower().replace(" ",""),field_name,re.IGNORECASE):
                        checkbox_input = speed_test_form_fields_name.find_element(*remote_test.Enable_FTP_Stop_Timer_checkbox[:2])
                        if (Enable_FTP_Stop_Timer_speed_data == "Yes" and not checkbox_input.is_selected()) or (Enable_FTP_Stop_Timer_speed_data == "No" and checkbox_input.is_selected()):
                            clickec(driver=speed_test_form_fields_name, locators=remote_test.Enable_FTP_Stop_Timer_checkbox)

                    elif re.search("Set Timeout".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Enable_FTP_Stop_Timer_speed_data == "Yes":
                            try:
                                speed_data = test_data["Set Timeout"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Speed_test:-'Set Timeout'", status="FAILED",comments=statement)
                                    result_status.put(status_df)
                                    raise e
                            if is_numeric(speed_data):
                                if 60 <= int(speed_data) <= 250:
                                    inputtext(driver=speed_test_form_fields_name,locators=remote_test.speed_Wait_Duration_textbox,value=speed_data)
                                else:
                                    statement = f"Check the value given within the range"
                                    with allure.step(statement):
                                        status_df = status(Title=Title, component=f"Speed_test:-'Set Timeout'", status="FAILED",comments=statement)
                                        result_status.put(status_df)
                                        e = Exception
                                        raise e
                            else:
                                statement = f"Check the entered value is numeric"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Speed_test:-'Set Timeout'", status="FAILED",comments=statement)
                                    result_status.put(status_df)
                                    e = Exception
                                    raise e
                except Exception as e:
                    Speed_test = False
                    continue
                            # speed_data = max(int(min_value), min(int(max_value), int(speed_data)))
            allure.attach(driver.get_screenshot_as_png(), name=f"Speed_test.",attachment_type=allure.attachment_type.PNG)
            if Speed_test == True:
                okbtn = driver.find_element(*remote_test.speedtest_okbtn[:2])
                if okbtn.is_enabled():
                    clickec(driver=driver,locators=remote_test.speedtest_okbtn)
                elif not okbtn.is_enabled():
                    clickec(driver=driver, locators=remote_test.speedtest_closebtn)
                    e = Exception
                    raise e
            elif Speed_test == False or Speed_test == None:
                clickec(driver=driver, locators=remote_test.speedtest_closebtn)
                e = Exception
                raise e
    except Exception as e:
        raise e
def is_numeric(value):
    try:
        float(value)
        return True
    except ValueError:
        return False
def http_speed_test(driver,Title,test_data,result_status,type_of_test):
    try:
        with allure.step("http speed test"):
            if type_of_test == "runtest":
                clickec(driver=driver, locators=remote_test.http_speed_test_checkbox)
            elif type_of_test == "continuoustest":
                clickec(driver=driver, locators=continuous.http_checkbox)
            elif type_of_test == "scheduletest":
                clickec(driver=driver, locators=schedule_test.http_speed_test_checkbox)

            WebDriverWait(driver, 10).until(EC.visibility_of_element_located(remote_test.http_speed_test_form))
            Http_speed_test = None
            http_speed_test_form_fields_names = driver.find_elements(*remote_test.http_speed_test_form)
            Http_speed_test = True
            try:
                Enter_custom_URL_http_speed_data = test_data["Enter custom URL"]
            except Exception as e:
                statement = f"Check parameter is present in remote test sheet test data excel"
                with allure.step(statement):
                    status_df = status(Title=Title, component=f"Http_speed_test:-'Enter custom URL'", status="FAILED",comments=statement)
                    result_status.put(status_df)
            try:
                Enable_HTTP_Speed_Test_Upload_Test_http_speed_data = test_data["Enable HTTP Speed Test Upload Test"]
            except Exception as e:
                statement = f"Check parameter is present in remote test sheet test data excel"
                with allure.step(statement):
                    status_df = status(Title=Title, component=f"Http_speed_test:-'Enable HTTP Speed Test Upload Test'", status="FAILED",comments=statement)
                    result_status.put(status_df)
            try:
                Enable_HTTP_Speed_Test_stop_timer_http_speed_data = test_data["Enable HTTP Speed Test stop timer"]
            except Exception as e:
                statement = f"Check parameter is present in remote test sheet test data excel"
                with allure.step(statement):
                    status_df = status(Title=Title, component=f"Http_speed_test:-'Enable HTTP Speed Test stop timer'",status="FAILED", comments=statement)
                    result_status.put(status_df)
            try:
                Enter_Custom_Upload_URL_http_speed_data = test_data["Enter Custom Upload URL"]
            except Exception as e:
                statement = f"Check parameter is present in remote test sheet test data excel"
                with allure.step(statement):
                    status_df = status(Title=Title, component=f"Http_speed_test:-'Enter Custom Upload URL'",status="FAILED", comments=statement)
                    result_status.put(status_df)
            for http_speed_test_form_fields_name in http_speed_test_form_fields_names:
                try:
                    field_name = str(http_speed_test_form_fields_name.text).lower().replace(" ","")
                    if re.search("Parallel Connections".lower().replace(" ",""),field_name,re.IGNORECASE):
                        pass
                    elif re.search("Enter custom URL".lower().replace(" ",""),field_name,re.IGNORECASE):
                        checkbox_input = driver.find_element(*remote_test.Enter_custom_URL_checkbox[:2])
                        if (Enter_custom_URL_http_speed_data == "Yes" and not checkbox_input.is_selected()) or (Enter_custom_URL_http_speed_data == "No" and checkbox_input.is_selected()):
                            clickec(driver=driver, locators=remote_test.Enter_custom_URL_checkbox)

                    elif re.search("Enter URL".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Enter_custom_URL_http_speed_data == "Yes":
                            try:
                                http_speed_data = test_data["Enter URL"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Http_speed_test:-'Enter Custom Upload URL'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    raise e
                            inputtext(driver=driver,locators=remote_test.Enter_URL_textbox,value=http_speed_data)
                    elif re.search("HTTP Speed Download Test File Size".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Enter_custom_URL_http_speed_data == "No" and Enable_HTTP_Speed_Test_stop_timer_http_speed_data == "No":
                            try:
                                http_speed_data = test_data["HTTP Speed Download Test File Size"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote_test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Http_speed_test:-'HTTP Speed Download Test File Size'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    raise e
                            HTTP_Speed_Download_Test_File_Size_dropdown_option = (By.XPATH,f"//div[@id='httpspeedtest']//option[normalize-space()='{http_speed_data}']","HTTP_Speed_Download_Test_File_Size_dropdown_option")
                            clickec(driver=driver,locators=remote_test.HTTP_Speed_Download_Test_File_Size_dropdown)
                            clickec(driver=driver,locators=HTTP_Speed_Download_Test_File_Size_dropdown_option)
                    elif re.search("Enable HTTP Speed Test Upload Test".lower().replace(" ",""),field_name,re.IGNORECASE):
                        checkbox_input = driver.find_element(*remote_test.Enable_HTTP_Speed_Test_Upload_Test_checkbox[:2])
                        if (Enable_HTTP_Speed_Test_Upload_Test_http_speed_data == "Yes" and not checkbox_input.is_selected()) or (Enable_HTTP_Speed_Test_Upload_Test_http_speed_data == "No" and checkbox_input.is_selected()):
                            clickec(driver=driver, locators=remote_test.Enable_HTTP_Speed_Test_Upload_Test_checkbox)

                    elif re.search("File Size".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Enable_HTTP_Speed_Test_Upload_Test_http_speed_data == "Yes" and Enable_HTTP_Speed_Test_stop_timer_http_speed_data == "No":
                            try:
                                http_speed_data = test_data["File Size"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title,component=f"Http_speed_test:-'File Size'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    raise e

                            if is_numeric(http_speed_data):
                                if 10 <= int(http_speed_data) <= 9999:
                                    inputtext(driver=driver,locators=remote_test.http_speed_File_Size_textbox,value=http_speed_data)
                                else:
                                    statement = f"Check the value given within the range"
                                    with allure.step(statement):
                                        status_df = status(Title=Title, component=f"Http_speed_test:-'File Size'", status="FAILED",comments=statement)
                                        result_status.put(status_df)
                                        e = Exception
                                        raise e
                            else:
                                statement = f"Check the entered value is numeric"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Http_speed_test:-'File Size'", status="FAILED",comments=statement)
                                    result_status.put(status_df)
                                    e = Exception
                                    raise e

                    elif re.search("Enable HTTP Speed Test stop timer".lower().replace(" ",""),field_name,re.IGNORECASE):
                        checkbox_input = driver.find_element(*remote_test.Enable_HTTP_Speed_Test_stop_timer_checkbox[:2])
                        if (Enable_HTTP_Speed_Test_stop_timer_http_speed_data == "Yes" and not checkbox_input.is_selected()) or (Enable_HTTP_Speed_Test_stop_timer_http_speed_data == "No" and checkbox_input.is_selected()):
                            clickec(driver=driver, locators=remote_test.Enable_HTTP_Speed_Test_stop_timer_checkbox)

                    elif re.search("Set Timeout".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Enable_HTTP_Speed_Test_stop_timer_http_speed_data == "Yes":
                            try:
                                http_speed_data = test_data["Set Timeout"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Http_speed_test:-'Set Timeout'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    raise e

                            if is_numeric(http_speed_data):
                                if 60 <= int(http_speed_data) <= 200:
                                    inputtext(driver=driver, locators=remote_test.HTTP_Speed_Set_Timeout_textbox,value=http_speed_data)
                                else:
                                    statement = f"Check the value given within the range"
                                    with allure.step(statement):
                                        status_df = status(Title=Title, component=f"Http_speed_test:-'Set Timeout'",status="FAILED", comments=statement)
                                        result_status.put(status_df)
                                        e = Exception
                                        raise e
                            else:
                                statement = f"Check the entered value is numeric"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Http_speed_test:-'Set Timeout'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    e = Exception
                                    raise e
                    elif re.search("Enter Custom Upload URL".lower().replace(" ",""),field_name,re.IGNORECASE):
                        checkbox_input = driver.find_element(*remote_test.Enter_Custom_Upload_URL_checkbox[:2])
                        if (Enter_Custom_Upload_URL_http_speed_data == "Yes" and not checkbox_input.is_selected()) or (Enter_Custom_Upload_URL_http_speed_data == "No" and checkbox_input.is_selected()):
                            clickec(driver=driver, locators=remote_test.Enter_Custom_Upload_URL_checkbox)
                    elif re.search("Upload URL".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Enter_Custom_Upload_URL_http_speed_data == "Yes":
                            try:
                                http_speed_data = test_data["Upload URL"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Http_speed_test:-'Upload URL'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    raise e
                            inputtext(driver=driver,locators=remote_test.Upload_URL_textbox,value=http_speed_data)
                except Exception as e:
                    Http_speed_test = False
                    continue
            allure.attach(driver.get_screenshot_as_png(), name=f"Speed_test.",attachment_type=allure.attachment_type.PNG)
            if Http_speed_test == True:
                okbtn = driver.find_element(*remote_test.http_speedtest_okbtn[:2])
                if okbtn.is_enabled():
                    clickec(driver=driver,locators=remote_test.http_speedtest_okbtn)
                elif not okbtn.is_enabled():
                    clickec(driver=driver, locators=remote_test.http_speedtest_closebtn)
                    e = Exception
                    raise e
            elif Http_speed_test == False or Http_speed_test == None:
                clickec(driver=driver, locators=remote_test.http_speedtest_closebtn)
                e = Exception
                raise e
    except Exception as e:
        raise e
def iperf_test(driver,Title,test_data,result_status,type_of_test,test_name):
    try:
        with allure.step("iperf test"):
            if type_of_test == "runtest":
                clickec(driver=driver, locators=remote_test.iperf_testcheckbox)
            elif type_of_test == "continuoustest":
                clickec(driver=driver, locators=continuous.iperf_checkbox)
            elif type_of_test == "scheduletest":
                clickec(driver=driver, locators=schedule_test.iperf_test_checkbox)

            WebDriverWait(driver, 10).until(EC.visibility_of_element_located(remote_test.iperf_test_form))
            Iperf_test = None
            iperf_test_form_fields_names = driver.find_elements(*remote_test.iperf_test_form)
            Iperf_test = True
            try:
                Use_Default_Server_iperf_data = test_data["Use Default Server"]
            except Exception as e:
                statement = f"Check parameter is present in remote test sheet test data excel"
                with allure.step(statement):
                    status_df = status(Title=Title, component=f"Iperf_test:-'Use Default Server'", status="FAILED",comments=statement)
                    result_status.put(status_df)
            try:
                Enable_Iperf_Upload_Test_iperf_data = test_data["Enable Iperf Upload Test"]
            except Exception as e:
                statement = f"Check parameter is present in remote test sheet test data excel"
                with allure.step(statement):
                    status_df = status(Title=Title, component=f"Iperf_test:-'Enable Iperf Upload Test'", status="FAILED",comments=statement)
                    result_status.put(status_df)
            try:
                Enable_Iperf_Upload_Test_iperf_data = test_data["Enable Iperf Upload Test"]
            except Exception as e:
                statement = f"Check parameter is present in remote test sheet test data excel"
                with allure.step(statement):
                    status_df = status(Title=Title, component=f"Iperf_test:-'Enable Iperf Upload Test'", status="FAILED",comments=statement)
                    result_status.put(status_df)
            try:
                TCP_Mode_iperf_data = test_data["TCP Mode"]
            except Exception as e:
                statement = f"Check parameter is present in remote test sheet test data excel"
                with allure.step(statement):
                    status_df = status(Title=Title, component=f"Iperf_test:-'TCP Mode'", status="FAILED",comments=statement)
                    result_status.put(status_df)
            try:
                UDP_Mode_iperf_data = test_data["UDP Mode"]
            except Exception as e:
                statement = f"Check parameter is present in remote test sheet test data excel"
                with allure.step(statement):
                    status_df = status(Title=Title, component=f"Iperf_test:-'UDP Mode'", status="FAILED",comments=statement)
                    result_status.put(status_df)
            for iperf_test_form_fields_name in iperf_test_form_fields_names:
                try:
                    field_name = str(iperf_test_form_fields_name.text).lower().replace(" ","")
                    if re.search("Use Default Server".lower().replace(" ",""),field_name,re.IGNORECASE):
                        checkbox_input = driver.find_element(*remote_test.Iperf_Use_Default_Server_checkbox[:2])
                        if (Use_Default_Server_iperf_data == "Yes" and not checkbox_input.is_selected()) or (Use_Default_Server_iperf_data == "No" and checkbox_input.is_selected()):
                            clickec(driver=driver, locators=remote_test.Iperf_Use_Default_Server_checkbox)
                    elif re.search("Host Name".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if Use_Default_Server_iperf_data == "No":
                            try:
                                iperf_data = test_data["Host Name"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Iperf_test:-'Host Name'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    raise  e
                            inputtext(driver=driver,locators=remote_test.Host_Name_textbox,value=iperf_data)
                    elif re.search("Test Duration".lower().replace(" ",""),field_name,re.IGNORECASE):
                            try:
                                iperf_data = test_data["Test Duration(sec)"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Iperf_test:-'Test Duration(sec)'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    raise e
                            if is_numeric(iperf_data):
                                if 10 <= int(iperf_data) <= 120:
                                    inputtext(driver=driver,locators=remote_test.Test_Duration_textbox,value=iperf_data)
                                else:
                                    statement = f"Check the value given within the range"
                                    with allure.step(statement):
                                        status_df = status(Title=Title, component=f"Iperf_test:-'Test Duration'",status="FAILED", comments=statement)
                                        result_status.put(status_df)
                                        e = Exception
                                        raise e
                            else:
                                statement = f"Check the entered value is numeric"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Iperf_test:-'Test Duration'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    e = Exception
                                    raise e
                    elif re.search("Enable Iperf Upload Test".lower().replace(" ",""),field_name,re.IGNORECASE):
                        checkbox_input = driver.find_element(*remote_test.Enable_Iperf_Upload_Test_checkbox[:2])
                        if (Enable_Iperf_Upload_Test_iperf_data == "Yes" and not checkbox_input.is_selected()) or (Enable_Iperf_Upload_Test_iperf_data == "No" and checkbox_input.is_selected()):
                            clickec(driver=driver, locators=remote_test.Enable_Iperf_Upload_Test_checkbox)
                    elif re.search("TCP Mode".lower().replace(" ",""),field_name,re.IGNORECASE) and test_name.lower().replace(" ", "") == "TCP-Iperf Test".lower().replace(" ", ""):
                        checkbox_input = driver.find_element(*remote_test.TCP_Mode_checkbox[:2])
                        checkbox_input1 = driver.find_element(*remote_test.UDP_Mode_checkbox[:2])
                        if (TCP_Mode_iperf_data == "Yes" and not checkbox_input.is_selected()) or (TCP_Mode_iperf_data == "No" and checkbox_input.is_selected()):
                            if checkbox_input1.is_selected():
                                clickec(driver=driver, locators=remote_test.UDP_Mode_checkbox)
                            clickec(driver=driver, locators=remote_test.TCP_Mode_checkbox)
                    elif re.search("UDP Mode".lower().replace(" ",""),field_name,re.IGNORECASE) and test_name.lower().replace(" ", "") == "UDP-Iperf Test".lower().replace(" ", ""):
                        checkbox_input = driver.find_element(*remote_test.UDP_Mode_checkbox[:2])
                        checkbox_input1 = driver.find_element(*remote_test.TCP_Mode_checkbox[:2])
                        if (UDP_Mode_iperf_data == "Yes" and not checkbox_input.is_selected()) or (UDP_Mode_iperf_data == "No" and checkbox_input.is_selected()):
                            if checkbox_input1.is_selected():
                                clickec(driver=driver, locators=remote_test.TCP_Mode_checkbox)
                            clickec(driver=driver, locators=remote_test.UDP_Mode_checkbox)
                    elif re.search("Enter the Bandwidth".lower().replace(" ",""),field_name,re.IGNORECASE):
                        if UDP_Mode_iperf_data == "Yes":
                            try:
                                iperf_data = test_data["Enter the Bandwidth"]
                            except Exception as e:
                                statement = f"Check parameter is present in remote_test sheet test data excel"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Iperf_test:-'Enter the Bandwidth'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    raise e
                            if is_numeric(iperf_data):
                                if 1 <= int(iperf_data) <= 9999:
                                    inputtext(driver=driver,locators=remote_test.Enter_the_Bandwidth_textbox,value=iperf_data)
                                else:
                                    statement = f"Check the value given within the range"
                                    with allure.step(statement):
                                        status_df = status(Title=Title, component=f"Iperf_test:-'Test Duration'",status="FAILED", comments=statement)
                                        result_status.put(status_df)
                                        e = Exception
                                        raise e
                            else:
                                statement = f"Check the entered value is numeric"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Iperf_test:-'Enter the Bandwidth'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    e = Exception
                                    raise e
                except Exception as e:
                    Iperf_test = False
                    continue
            allure.attach(driver.get_screenshot_as_png(), name=f"Speed_test.",attachment_type=allure.attachment_type.PNG)
            if Iperf_test == True:
                okbtn = driver.find_element(*remote_test.iperf_test_okbtn[:2])
                if okbtn.is_enabled():
                    clickec(driver=driver,locators=remote_test.iperf_test_okbtn)
                elif not okbtn.is_enabled():
                    clickec(driver=driver, locators=remote_test.iperf_test_closebtn)
                    e = Exception
                    raise e
            elif Iperf_test == False or Iperf_test == None:
                clickec(driver=driver, locators=remote_test.iperf_test_closebtn)
                e = Exception
                raise e
    except Exception as e:
        raise e
def web_test(driver,test_data,type_of_test):
    try:
        with allure.step("Web test"):
            if type_of_test == "runtest":
                clickec(driver=driver, locators=remote_test.webtest_checkbox)
            elif type_of_test == "continuoustest":
                clickec(driver=driver, locators=continuous.web_checkbox)
            elif type_of_test == "scheduletest":
                clickec(driver=driver, locators=schedule_test.web_test_checkbox)

            web_data = None
            try:
                web_data = test_data["URL"]
            except Exception as e:
                raise
            # Interaction with URL input field and "OK" button
            clickec(driver=driver, locators=remote_test.web_url)
            inputtext(driver=driver, locators=remote_test.web_url, value=web_data)
            allure.attach(driver.get_screenshot_as_png(), name=f"Web Test.", attachment_type=allure.attachment_type.PNG)
            clickec(driver=driver, locators=remote_test.web_test_okbtn)
    except Exception as e:
        print(e)
        raise e
def stream_test(driver,test_data,type_of_test):
    try:
        with allure.step("stream test"):
            if type_of_test == "runtest":
                clickec(driver=driver, locators=remote_test.stream_checkbox)
            elif type_of_test == "continuoustest":
                clickec(driver=driver, locators=continuous.stream_checkbox)
            elif type_of_test == "scheduletest":
                clickec(driver=driver, locators=schedule_test.stream_test_checkbox)
            stream_data = None
            try:
                stream_data = test_data["Enter_Video_URL"]
            except Exception as e:
                raise
            clickec(driver,remote_test.enter_url_checkbox)
            inputtext(driver,locators=remote_test.txt_box_url,value=stream_data)
            allure.attach(driver.get_screenshot_as_png(), name=f"Enter video URL value", attachment_type=allure.attachment_type.PNG)
            clickec(driver,remote_test.submit_ok_btn)
    except Exception as e:
        raise e

##################################################################### Phase - 6 ####################################################################################################
def group_for_remotetest(driver,device,excelpath,campaign):
    Title = "GROUP"
    result_status = queue.Queue()
    Group_runvalue = Testrun_mode(value="Group")
    flag_delete_device = []
    try:
        if "Yes".lower() == Group_runvalue[-1].strip().lower():
            timestamp_value = datetime.datetime.now().strftime("%d%m%Y%H%M%S")
            random_length = random.randint(3, 5)
            random_alphabet = generate_random_alphabet(random_length)
            timestamp_value = str(random_alphabet + timestamp_value)
            clickec(driver, remote_test.remotetest)
            time.sleep(3)
            check_android_pro_is_active_in_remotetest(driver)
            testgroup = timestamp_value
            add_test_group(driver,Title,device,campaign,result_status,timestamp_value)
            Page_Down(driver=driver)
            time.sleep(3)
            try:
                add_device(driver,testgroup,device,Title,campaign,result_status)
                delete_device(driver, testgroup, device, result_status, Title, campaign,flag_delete_device)
            except Exception as e:
                pass
            delete_group(driver, testgroup, result_status, Title, device, campaign,flag_delete_device)
        elif "No".lower() == Group_runvalue[-1].strip().lower():
            statement = "You have selected Not to execute"
            with allure.step(statement):
                status_df = status(Title, "Not to execute", "SKIPPED", "You have selected No for execute")
                result_status.put(status_df)
                pass
    except Exception as e:
        pass
    finally:
        if "Yes".lower() == Group_runvalue[-1].strip().lower():
            click(driver=driver, locators=Login_Logout.dashboard_id)
    try:
        update_remote_test_result(result_status, excelpath)
    except Exception as e:
        pass
def add_device(driver,testgroup,device,Title,campaign,result_status):
    click_on_test_group_button_to_open_dropdown(driver, testgroup)
    edit_group_locator = driver.find_element(*edit_group.edit_group_btn)
    action_chain = ActionChains(driver)
    action_chain.move_to_element(edit_group_locator).perform()
    clickec(driver,edit_group.add_device_btn)
    add_device_selection = (By.XPATH,f"//div[@class='table-responsive']//td[normalize-space()='{device}']/preceding-sibling::td//input[@name ='adddeviceChk']")
    add_devicetab_close = (By.XPATH, "//div[@ng-form='addDeviceForm']//button[@type='button'][normalize-space()='Close']","add_devicetab_close")
    try:
        WebDriverWait(driver, 10).until(EC.invisibility_of_element_located(add_device_selection))
        statement = "Device is added during the 'Test Group Creation'"
        with allure.step(statement):
            status_df = status(Title,component=f"device:- {device},campaign:- {campaign},",status="PASSED", comments=statement)
            result_status.put(status_df)
            allure.attach(driver.get_screenshot_as_png(), name=statement,attachment_type=allure.attachment_type.PNG)
        clickec(driver,add_devicetab_close)
    except Exception as e:
        clickec(driver,add_devicetab_close)
        statement = "Test group is created but device is not added during Test Group Creation.Hence 'Delete Device' action is not performed"
        with allure.step(statement):
            status_df = status(Title,component=f"device:- {device},campaign:- {campaign},",status="PASSED", comments=statement)
            result_status.put(status_df)
            allure.attach(driver.get_screenshot_as_png(), name=statement,attachment_type=allure.attachment_type.PNG)
            raise e

def add_test_group(driver,Title,device,campaign,result_status,timestamp_value):
    try:
        with allure.step("Addition of Test Group"):
            time.sleep(5)
            clickec(driver, add_test_group_1.test_group_btn)
            inputtext(driver, locators=add_test_group_1.group_name, value=timestamp_value)
            allure.attach(driver.get_screenshot_as_png(), name=f"Test Group Name",attachment_type=allure.attachment_type.PNG)
            clickec(driver, add_test_group_1.next_button)
            device_selection = (By.XPATH,f"//div[@class='table-responsive']//td[normalize-space()='{device}']/preceding-sibling::td//input[@name ='listOfdeviceChk']","Device Name Selection")
            clickec(driver, device_selection)
            clickec(driver, add_test_group_1.add_button)
            verify_group(driver,Title,device,campaign,result_status,testgroup=timestamp_value,passed_statement="Test Group is created successfully by selecting device",failed_statement="Failed to create Test Group",locator_toverify=(By.XPATH,f"//p[normalize-space()='{timestamp_value}']/ancestor::div[@class='deviceCtldiv']/following-sibling::div//div//button"))
    except Exception as e:
       raise

def delete_device(driver,testgroup,device,result_status,Title,campaign,flag_delete_device):
    alert_text = None
    try:
        with allure.step("REMOTE TEST DELETE DEVICE"):
            time.sleep(10)
            click_on_test_group_button_to_open_dropdown(driver, testgroup)
            edit_group_locator = driver.find_element(*edit_group.edit_group_btn)
            action_chain = ActionChains(driver)
            action_chain.move_to_element(edit_group_locator).perform()
            time.sleep(1)
            clickec(driver=driver, locators=edit_group.delete_device_btn)
            delete_device_checkbox = (By.XPATH,f"//table[@class='table table-bordered table-hover table-striped']//td[normalize-space()='{device}']/preceding-sibling::td//input[@name='deleteChk']","device delete checkbox")
            clickec(driver=driver, locators=delete_device_checkbox)
            clickec(driver=driver, locators=edit_group.delete_btn)
            alert_locator = (By.XPATH,"//div[@role='alert']")
            try:
                try:
                    WebDriverWait(driver, 20).until(EC.visibility_of_element_located(alert_locator))
                except Exception as e:
                    pass
                alert = driver.find_element(*alert_locator)
                alert_text = alert.get_attribute("innerText")
                removeitems = ["\n","×"]
                for a in removeitems:
                    alert_text = str(str(alert_text).replace(a, "")).lower()
            except:
                pass
            if re.search(str(alert_text).replace(" ",""),'Devices successfully deleted from group'.replace(" ","").lower(),re.IGNORECASE):
                statement = f"{alert_text}"
                with allure.step(statement):
                    allure.attach(driver.get_screenshot_as_png(), name=statement,attachment_type=allure.attachment_type.PNG)
                    status_df = status(Title,component=f"device:- {device},campaign:- {campaign}",status="PASSED", comments=statement)
                    result_status.put(status_df)
                    flag_delete_device.append(alert_text)
            elif not re.search(str(alert_text).replace(" ",""),'Devices successfully deleted from group'.replace(" ","").lower(),re.IGNORECASE):
                statement = f"{alert_text}"
                with allure.step(statement):
                    allure.attach(driver.get_screenshot_as_png(), name=statement,attachment_type=allure.attachment_type.PNG)
                    status_df = status(Title,component=f"device:- {device},campaign:- {campaign}",status="FAILED", comments=statement)
                    result_status.put(status_df)
                    flag_delete_device.append(alert_text)
    except Exception as e:
        pass

def delete_group(driver,testgroup,result_status,Title,device,campaign,flag_delete_device):
    alert_text = None
    try:
        with allure.step("REMOTE TEST DELETE GROUP"):
            time.sleep(10)
            click_on_test_group_button_to_open_dropdown(driver, testgroup)
            edit_group_locator = driver.find_element(*edit_group.edit_group_btn)
            action_chain = ActionChains(driver)
            action_chain.move_to_element(edit_group_locator).perform()
            clickec(driver=driver, locators=edit_group.delete_group_btn_in_edit_group)
            clickec(driver=driver, locators=edit_group.delete_group_btn)
            alert_locator = (By.XPATH, "//div[@role='alert']")
            try:
                try:
                    WebDriverWait(driver, 20).until(EC.visibility_of_element_located(alert_locator))
                except Exception as e:
                    pass
                alert = driver.find_element(*alert_locator)
                alert_text = alert.get_attribute("innerText")
                removeitems = ["\n", "×"]
                for a in removeitems:
                    alert_text = str(str(alert_text).replace(a, "")).lower()
            except:
                pass
            if re.search(str(alert_text).replace(" ",""), 'Group has been deleted.'.replace(" ","").lower(),re.IGNORECASE):
                statement = f"{flag_delete_device[-1]} and {alert_text}"
                with allure.step(statement):
                    allure.attach(driver.get_screenshot_as_png(), name=statement,attachment_type=allure.attachment_type.PNG)
                    status_df = status(Title,component=f"device:- {device},campaign:- {campaign}",status="PASSED", comments=statement)
                    result_status.put(status_df)
            elif not re.search(str(alert_text).replace(" ",""), 'Group has been deleted.'.replace(" ","").lower(),re.IGNORECASE):
                statement = f"{flag_delete_device[-1]} and {alert_text}"
                with allure.step(statement):
                    allure.attach(driver.get_screenshot_as_png(), name=statement,attachment_type=allure.attachment_type.PNG)
                    status_df = status(Title,component=f"device:- {device},campaign:- {campaign}",status="FAILED", comments=statement)
                    result_status.put(status_df)
            time.sleep(10)
    except Exception as e:
        pass
#####################################################################################################################################################################################################################
def verify_group(driver,Title,device,campaign,result_status,testgroup,passed_statement,failed_statement,locator_toverify):
    with allure.step(passed_statement):
        try:
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located(locator_toverify))
            status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},testgroup:- {testgroup}",status="PASSED", comments=passed_statement)
            result_status.put(status_df)
            allure.attach(driver.get_screenshot_as_png(), name=passed_statement, attachment_type=allure.attachment_type.PNG)
        except Exception as e:
            with allure.step(failed_statement):
                status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},testgroup:- {testgroup}",status="FAILED", comments=failed_statement)
                result_status.put(status_df)
                allure.attach(driver.get_screenshot_as_png(), name=failed_statement, attachment_type=allure.attachment_type.PNG)
            raise failed_statement

#########################################################################################################################################################################################################################################
def continuous_test(driver,testgroup,continuous_option):
    clickec(driver, remote_test.remotetest)
    click_on_test_group_button_to_open_dropdown(driver, testgroup)
    time.sleep(3)
    action_chain = ActionChains(driver)
    last_option_restart = (By.XPATH,"//div[@class='btn-group open']//a[@data-target='#restartdevice'][contains(text(),'Restart')]")
    try:
        last_option_restart_element = driver.find_element(*last_option_restart)
        action_chain.move_to_element(last_option_restart_element).perform()
    except:
        pass
    continuous_element_locator = driver.find_element(*continuous.continuous_hover)
    action_chain.move_to_element(continuous_element_locator).perform()
    time.sleep(1)
    if continuous_option == "add continuous Test":
        click(continuous_element_locator,continuous.add_continuous_button)
    elif continuous_option == "Delete continuous Test":
        click(continuous_element_locator,continuous.delete_continuous_button)
    time.sleep(2)

def continuoustest_for_android_pro(driver,Title,device,campaign,usercampaignsname,testgroup,tests,result_status,test_complete,test_Selected,run_test_status_value,multi_bparty_call_test_flag,excelpath):
    try:
        with allure.step("continuous test for android pro"):
            flag_test_group = []
            alert_text = None
            check_android_pro_is_active_in_remotetest(driver)
            verify_test_group_is_present(driver,Title,device,campaign,usercampaignsname,result_status,testgroup,flag_test_group)
            if flag_test_group == [True]:
                scheduletest_runvalue = Testrun_mode(value="Schedule test")
                continuoustest_runvalue = Testrun_mode(value="Continuous Test")
                if "Yes".lower() != scheduletest_runvalue[-1].strip().lower() and "Yes".lower() == continuoustest_runvalue[-1].strip().lower():
                    device_button_dropdown = click_on_test_group_button_to_open_dropdown(driver, testgroup)
                    click_on_the_check_devices(driver=driver,driver1=device_button_dropdown)
                try:
                    alert_text = alert_accept(driver=driver)
                except Exception as e:
                    pass
                if alert_text !=None:
                    with allure.step(f"{alert_text}"):
                        status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED",comments=f"{alert_text}")
                        result_status.put(status_df)
                elif alert_text == None:
                    with allure.step(f"no alert found, device is registered"):
                        status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}", status="PASSED", comments=f"no alert found, device is registered")
                        result_status.put(status_df)
                    flag_status_value = []
                    scheduletest_runvalue = Testrun_mode(value="Schedule test")
                    continuoustest_runvalue = Testrun_mode(value="Continuous Test")
                    if "Yes".lower() != scheduletest_runvalue[-1].strip().lower() and "Yes".lower() == continuoustest_runvalue[-1].strip().lower():
                        device = check_device_status(driver,Title,device,campaign,usercampaignsname,flag_status_value,result_status)
                    continuous_test(driver,testgroup,continuous_option="add continuous Test")
                    continuous_test_form(driver,Title,usercampaignsname,tests,test_complete,test_Selected,result_status,excelpath,testgroup,multi_bparty_call_test_flag)
                    if test_complete == [True] and test_Selected == [True]:
                        statement = "Successfully entered test data for a particular type of test and clicked on the start button."
                        with allure.step(statement):
                            allure.attach(driver.get_screenshot_as_png(), name=f"{statement}",attachment_type=allure.attachment_type.PNG)
                            click(driver=driver,locators=continuous.run_button)
                            status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="PASSED", comments=statement)
                            result_status.put(status_df)
                            check_status_of_test(driver, Title, device, campaign, usercampaignsname,run_test_status_value,result_status)
                    elif (test_complete == [] or test_complete == [False]) and (test_Selected == [False] or test_Selected == []):
                        statement = "Test data was not entered successfully for a particular type of test,so clicked on the close button."
                        with allure.step(statement):
                            allure.attach(driver.get_screenshot_as_png(), name=f"{statement}",attachment_type=allure.attachment_type.PNG)
                            status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED", comments=statement)
                            result_status.put(status_df)
                            click(driver=driver, locators=continuous.close_button)
                        time.sleep(0.1)
    except Exception as e:
        pass
    finally:
        return device
def continuous_test_form(driver,Title,usercampaignsname,tests,test_complete,test_Selected,result_status,excelpath,testgroup,multi_bparty_call_test_flag):
    try:
        with allure.step("continuous test form"):
            inputtext(driver=driver, locators=continuous.test_name_field, value=f"{usercampaignsname}")
            allure.attach(driver.get_screenshot_as_png(), name=f"continuous test form.",attachment_type=allure.attachment_type.PNG)
            try:
                df_remote_test = pd.read_excel(config.test_data_path,sheet_name="Remote_Test")
            except Exception as e:
                with allure.step(f"Check {config.test_data_path}"):
                    print(f"Check {config.test_data_path}")
                    assert False
            txt = []
            if tests.__len__() == 0:
                statement = f"{Title}  --  Nothing is marked 'Yes' in {str(config.test_data_path)}"
                with allure.step(f"Nothing is marked 'Yes' in {str(config.test_data_path)} for '{Title}'"):
                    updatecomponentstatus(Title, "", "FAILED", f"Nothing marked in {str(config.test_data_path)}",excelpath)
                    e = Exception
                    raise e
            else:
                test_complete.append(True)
                for test in tests:
                    try:
                        if test.lower().replace(" ", "") == "tcp-iperftest".lower().replace(" ", "") or test.lower().replace(" ", "") == "udp-iperftest".lower().replace(" ", ""):
                             testmodified = test.replace("TCP-", "").replace("UDP-", "")
                             remote_test_datas = df_remote_test[df_remote_test['Test_Type'].str.contains(testmodified,case=False, na=False)].groupby('Test_Type').apply(lambda x: x[['Parameter', 'Value']].to_dict(orient='records')).tolist()
                        else:
                            remote_test_datas = df_remote_test[df_remote_test['Test_Type'].str.contains(test, case=False, na=False)].groupby('Test_Type').apply(lambda x: x[['Parameter', 'Value']].to_dict(orient='records')).tolist()
                        remote_test_dict = {str(record['Parameter']).strip(): record['Value'] for record in remote_test_datas[0]}
                        test = test.strip().lower().replace(" ", "")
                        test_functions = {
                            "pingtest": ping_test,
                            "calltest": call_test,
                            "smstest": sms_test,
                            "speed_test": speed_test,
                            "httpspeedtest": http_speed_test,
                            "tcp-iperftest": iperf_test,
                            "udp-iperftest": iperf_test,
                            "webtest": web_test,
                            "streamtest": stream_test
                        }
                        test_name = test.lower().replace(" ", "")
                        if re.fullmatch("multi.*b.*party.*call.*test", test_name, re.IGNORECASE):
                            multi_bparty_call_test_flag.append(True)
                            multi_bparty_call_test(driver=driver, Title=Title, test_data=remote_test_dict,result_status=result_status)

                        elif test_name in test_functions:
                            func = test_functions[test_name]
                            if test_name in ["webtest", "streamtest"]:
                                func(driver=driver, test_data=remote_test_dict, type_of_test="continuoustest")
                            elif test_name in ["tcp-iperftest", "udp-iperftest"]:
                                func(driver=driver, Title=Title, test_data=remote_test_dict,result_status=result_status, type_of_test="continuoustest", test_name=test_name)
                            else:
                                func(driver=driver, Title=Title, test_data=remote_test_dict,result_status=result_status, type_of_test="continuoustest")
                    except Exception as e:
                        try:
                            test_complete.remove(True)
                        except Exception as e:
                            pass
                        test_complete.append(False)
                        test_complete = list(set(test_complete))
                        continue
                test_Selected.append(True)
                for test in tests:
                    try:
                        test = test.strip().lower().replace(" ", "")
                        checkboxes = {
                            "pingtest": continuous.ping_checkbox,
                            "calltest": continuous.call_checkbox,
                            "smstest": continuous.sms_checkbox,
                            "speed_test": continuous.speed_checkbox,
                            "httpspeedtest": continuous.http_checkbox,
                            "webtest": continuous.web_checkbox,
                            "streamtest": continuous.stream_checkbox,
                            "tcp-iperftest": continuous.iperf_checkbox,
                            "udp-iperftest": continuous.iperf_checkbox,
                            "multi_bparty_calltest": continuous.multi_bparty_checkbox
                        }
                        test_name = test.lower().replace(" ", "")
                        if re.fullmatch("multi bparty call test", test_name, re.IGNORECASE):
                            typeoftest_is_selected(driver, continuous.multi_bparty_checkbox)
                        else:
                            for name, checkbox in checkboxes.items():
                                if re.fullmatch(name, test_name):
                                    typeoftest_is_selected(driver, checkbox)
                                    break
                    except Exception as e:
                        try:
                            test_Selected.remove(True)
                        except Exception as e:
                            pass
                        test_Selected.append(False)
                        test_Selected = list(set(test_complete))
                        continue
    except Exception as e:
        pass
def verifying_of_test_execution_for_continuous(driver,Title,device,campaign,usercampaignsname,testgroup,continuous_campaigns,test_name_in_table_view,run_test_status_value,result_status,test_Execution_status,campaigns_status,usercampaignsname_list,multi_bparty_call_test_flag):
    try:
        with allure.step("Verification of the success or failed status from the application after clicking on start button in run test of remote test"):
            action = ActionChains(driver)
            try:
                table_view_refresh_element = driver.find_element(*remote_test.table_view_refresh[:2])
                action.move_to_element(table_view_refresh_element).perform()
            except Exception as e:
                pass
            execution_time = 0
            total_time = 211
            for i in range(0,total_time):
                try:
                    clickec(driver=driver,locators=remote_test.table_view_refresh)
                    if multi_bparty_call_test_flag == []:
                        WebDriverWait(driver, 10).until(min_max_elements_present(test_name_in_table_view, min_count=2, max_count=3))
                    elif multi_bparty_call_test_flag == [True]:
                        WebDriverWait(driver, 10).until(min_max_elements_present(test_name_in_table_view, min_count=4, max_count=6))
                    execution_time = total_time - i
                    break
                except Exception as e:
                    continue
            time.sleep(180)
            continuous_test(driver, testgroup, continuous_option="Delete continuous Test")
            click(driver=driver, locators=Login_Logout.dashboard_id)
            time.sleep(2)
            try:
                usercampaignsname_elements = driver.find_elements(*test_name_in_table_view)
                for usercampaignsname_element in usercampaignsname_elements:
                    usercampaignsname1 = usercampaignsname_element.text
                    usercampaignsname_list.append(usercampaignsname1)
            except Exception as e:
                pass
            allure.attach(driver.get_screenshot_as_png(), name=f"verifying_of_test_execution",attachment_type=allure.attachment_type.PNG)
            try:
                Page_Down(driver=driver)
                if multi_bparty_call_test_flag == []:
                    test_name_siblings = WebDriverWait(driver, 10).until(min_max_elements_present(test_name_in_table_view, min_count=2, max_count=3))
                elif multi_bparty_call_test_flag == [True]:
                    test_name_siblings = WebDriverWait(driver, 10).until(min_max_elements_present(test_name_in_table_view, min_count=4, max_count=6))
                campaigns_status.append(True)
                if run_test_status_value == [True]:
                    status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="PASSED", comments=f"Success")
                    result_status.put(status_df)
                    waiting_for_complete_or_Aborted_status_for_continuous_test(Title,device,campaign,usercampaignsname,continuous_campaigns,result_status,execution_time,driver=driver,test_Execution_status=test_Execution_status,multi_bparty_call_test_flag=multi_bparty_call_test_flag)
                elif run_test_status_value == [False]:
                    status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="WARNING", comments=f"Remote Test config failed but campaign available")
                    result_status.put(status_df)
                    waiting_for_complete_or_Aborted_status_for_continuous_test(Title,device,campaign,usercampaignsname,continuous_campaigns,result_status,execution_time,driver=driver,test_Execution_status=test_Execution_status,multi_bparty_call_test_flag=multi_bparty_call_test_flag)
            except Exception as e:
                campaigns_status.append(False)
                test_Execution_status.append(None)
                if run_test_status_value == [True]:
                    status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED", comments=f"Remote Test config successful but campaign not available")
                    result_status.put(status_df)
                elif run_test_status_value == [False]:
                    status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED", comments=f"	Remote Test config failed to run")
                    result_status.put(status_df)
                pass
    except Exception as e:
        pass
def waiting_for_complete_or_Aborted_status_for_continuous_test(Title,device,campaign,usercampaignsname,continuous_campaigns,result_status,execution_time,driver,test_Execution_status,multi_bparty_call_test_flag):
    with allure.step("waiting for complete or Aborted status"):
        try:
            action = ActionChains(driver)
            for i in range(0, execution_time):
                try:
                    clickec(driver=driver, locators=remote_test.table_view_refresh)
                    test_execution = (By.XPATH,f"//td[@id='loaderCamp']/following-sibling::td[contains(.,'{device}')]/preceding-sibling::td[starts-with(normalize-space(),'{continuous_campaigns}')]/following-sibling::td[contains(.,'COMPLETED') or contains(.,'ABORTED')]")
                    if multi_bparty_call_test_flag == []:
                        WebDriverWait(driver, 10).until(min_max_elements_present(test_execution, min_count=2, max_count=3))
                    elif multi_bparty_call_test_flag == [True]:
                        WebDriverWait(driver, 10).until(min_max_elements_present(test_execution, min_count=4, max_count=6))
                    break
                except Exception as e:
                    continue
            for i in range(0, 1):
                try:
                    clickec(driver=driver, locators=remote_test.table_view_refresh)
                    test_execution = (By.XPATH,f"//td[@id='loaderCamp']/following-sibling::td[contains(.,'{device}')]/preceding-sibling::td[starts-with(normalize-space(),'{continuous_campaigns}')]/following-sibling::td[contains(.,'COMPLETED') or contains(.,'ABORTED')]")
                    if multi_bparty_call_test_flag == []:
                        WebDriverWait(driver, 10).until(min_max_elements_present(test_execution, min_count=2, max_count=3))
                    elif multi_bparty_call_test_flag == [True]:
                        WebDriverWait(driver, 10).until(min_max_elements_present(test_execution, min_count=4, max_count=6))
                    try:
                        test_execution_element = driver.find_element(*test_execution)
                        action.move_to_element(test_execution_element).perform()
                    except Exception as e:
                        pass
                    test_Execution_status.append(True)
                    status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="PASSED", comments=f"Test Campaigns execution status is Completed/Aborted/Uploaded")
                    result_status.put(status_df)
                    break
                except Exception as e:
                    test_Execution_status.append(False)
                    status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED", comments=f"Test Campaigns execution status is Executing")
                    result_status.put(status_df)
                    try:
                        test_executing = (By.XPATH,f"//td[@id='loaderCamp']/following-sibling::td[contains(.,'{device}')]/preceding-sibling::td[starts-with(normalize-space(),'{continuous_campaigns}')]/following-sibling::td[contains(.,'EXECUTING')]")
                        try:
                            test_executing_element = driver.find_element(*test_executing)
                            action.move_to_element(test_executing_element).perform()
                        except Exception as e:
                            pass
                        WebDriverWait(driver, 0.1).until(EC.invisibility_of_element_located(test_executing))
                    except Exception as e:
                        pass
            allure.attach(driver.get_screenshot_as_png(), name=f"waiting for complete or Aborted status",attachment_type=allure.attachment_type.PNG)
        except Exception as e:
            pass
def schedultest_form(driver,Title,usercampaignsname,tests,test_complete,test_Selected,result_status,excelpath):
    try:
        with allure.step("schedule test form"):
            inputtext(driver=driver,locators=schedule_test.test_name,value=f"{usercampaignsname}")
            inputtext(driver=driver,locators=schedule_test.iteration_textbox,value="1")
            inputtext(driver=driver, locators=schedule_test.delays_bw_tests, value="5")
            calendar_trigger = driver.find_element(By.XPATH,"//div[@id='datetimePicker']//span[@class='glyphicon glyphicon-calendar']")
            calendar_trigger.click()
            # Wait for the calendar to be visible
            wait = WebDriverWait(driver, 10)
            calendar_popup=wait.until(EC.visibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[1]/div[1]/div[2]/section[1]/div[1]/section[1]/div[9]/div[1]/div[1]/div[2]/div[1]/div[1]/form[1]/div[6]/div[1]/div[1]/div[1]")))
            # Find and click on the desired date element
            desired_date = calendar_popup.find_element(By.XPATH, "//span[@class='glyphicon glyphicon-time']")
            desired_date.click()
            time.sleep(2)
            click_multiple_times(driver,schedule_test.increse_time_up_arrow_btn,2)
            time.sleep(2)
            calendar_trigger.click()
            allure.attach(driver.get_screenshot_as_png(), name=f"run_test_form.",attachment_type=allure.attachment_type.PNG)
            try:
                df_remote_test = pd.read_excel(config.test_data_path,sheet_name="Remote_Test")
            except Exception as e:
                with allure.step(f"Check {config.test_data_path}"):
                    print(f"Check {config.test_data_path}")
                    assert False
            txt = []
            if tests.__len__() == 0:
                statement = f"{Title}  --  Nothing is marked 'Yes' in {str(config.test_data_path)}"
                with allure.step(f"Nothing is marked 'Yes' in {str(config.test_data_path)} for '{Title}'"):
                    updatecomponentstatus(Title, "", "FAILED", f"Nothing marked in {str(config.test_data_path)}",excelpath)
                    e = Exception
                    raise e
            else:
                test_complete.append(True)
                for test in tests:
                    try:
                        if test.lower().replace(" ", "") == "tcp-iperftest".lower().replace(" ", "") or test.lower().replace(" ", "") == "udp-iperftest".lower().replace(" ", ""):
                             testmodified = test.replace("TCP-", "").replace("UDP-", "")
                             remote_test_datas = df_remote_test[df_remote_test['Test_Type'].str.contains(testmodified,case=False, na=False)].groupby('Test_Type').apply(lambda x: x[['Parameter', 'Value']].to_dict(orient='records')).tolist()
                        else:
                            remote_test_datas = df_remote_test[df_remote_test['Test_Type'].str.contains(test, case=False, na=False)].groupby('Test_Type').apply(lambda x: x[['Parameter', 'Value']].to_dict(orient='records')).tolist()
                        remote_test_dict = {str(record['Parameter']).strip(): record['Value'] for record in remote_test_datas[0]}
                        test = test.strip().lower().replace(" ", "")
                        test_functions = {
                            "pingtest": ping_test,
                            "calltest": call_test,
                            "smstest": sms_test,
                            "speed_test": speed_test,
                            "httpspeedtest": http_speed_test,
                            "webtest": web_test,
                            "streamtest": stream_test,
                            "tcp-iperftest": iperf_test,
                            "udp-iperftest": iperf_test
                        }
                        test_name = test.lower().replace(" ", "")
                        for name, func in test_functions.items():
                            if re.fullmatch(name, test_name):
                                if name in ["webtest", "streamtest"]:
                                    func(driver=driver, test_data=remote_test_dict, type_of_test="scheduletest")
                                elif name in ["tcp-iperftest", "udp-iperftest"]:
                                    func(driver=driver, Title=Title, test_data=remote_test_dict,result_status=result_status, type_of_test="scheduletest",test_name=test_name)
                                else:
                                    func(driver=driver, Title=Title, test_data=remote_test_dict,result_status=result_status, type_of_test="scheduletest")
                                break
                    except Exception as e:
                        try:
                            test_complete.remove(True)
                        except Exception as e:
                            pass
                        test_complete.append(False)
                        test_complete = list(set(test_complete))
                        continue
                test_Selected.append(True)
                for test in tests:
                    try:
                        test = test.strip().lower().replace(" ", "")
                        test_checkboxes = {
                            "pingtest": schedule_test.ping_test_checkbox,
                            "calltest": schedule_test.call_test_checkbox,
                            "smstest": schedule_test.sms_test_checkbox,
                            "speed_test": schedule_test.speed_test_checkbox,
                            "httpspeedtest": schedule_test.http_speed_test_checkbox,
                            "tcp-iperftest": schedule_test.iperf_test_checkbox,
                            "udp-iperftest": schedule_test.iperf_test_checkbox,
                            "webtest": schedule_test.web_test_checkbox,
                            "streamtest": schedule_test.stream_test_checkbox
                        }
                        test_name = test.lower().replace(" ", "")
                        for name, checkbox in test_checkboxes.items():
                            if re.fullmatch(name, test_name):
                                typeoftest_is_selected(driver, checkbox)
                                break
                    except Exception as e:
                        try:
                            test_Selected.remove(True)
                        except Exception as e:
                            pass
                        test_Selected.append(False)
                        test_Selected = list(set(test_complete))
                        continue
    except Exception as e:
        pass
def click_multiple_times(driver, locators, times):
    for i in range(times):
        clickec(driver=driver, locators=locators)
def schedule_for_android_pro(driver,Title,device,campaign,usercampaignsname,testgroup,tests,result_status,test_complete,test_Selected,run_test_status_value,excelpath):
    try:
        with allure.step("schedule test for android pro"):
            flag_test_group = []
            alert_text = None
            check_android_pro_is_active_in_remotetest(driver)
            verify_test_group_is_present(driver,Title,device,campaign,usercampaignsname,result_status,testgroup,flag_test_group)
            if flag_test_group == [True]:
                device_button_dropdown = click_on_test_group_button_to_open_dropdown(driver, testgroup)
                click_on_the_check_devices(driver=driver,driver1=device_button_dropdown)
                try:
                    alert_text = alert_accept(driver=driver)
                except Exception as e:
                    pass
                if alert_text !=None:
                    with allure.step(f"{alert_text}"):
                        status_df = status(Title=Title,component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED",comments=f"{alert_text}")
                        result_status.put(status_df)
                elif alert_text == None:
                    with allure.step(f"no alert found, device is registered"):
                        status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}", status="PASSED", comments=f"no alert found, device is registered")
                        result_status.put(status_df)
                    flag_status_value = []
                    device = check_device_status(driver,Title,device,campaign,usercampaignsname,flag_status_value,result_status)
                    device_button_dropdown = click_on_test_group_button_to_open_dropdown(driver, testgroup)
                    clickec(driver=driver, locators=schedule_test.schedule_test_btn)
                    schedultest_form(driver, Title, usercampaignsname, tests, test_complete, test_Selected,result_status,excelpath)
                    if test_complete == [True] and test_Selected == [True]:
                        statement = "Successfully entered test data for a particular type of test and clicked on the start button."
                        with allure.step(statement):
                            allure.attach(driver.get_screenshot_as_png(), name=f"{statement}",attachment_type=allure.attachment_type.PNG)
                            click(driver=driver,locators=schedule_test.run_button)
                            status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="PASSED", comments=statement)
                            result_status.put(status_df)
                            try:
                                WebDriverWait(driver, 120).until(EC.visibility_of_element_located(remote_test.run_test_start_status))
                            except Exception as e:
                                pass
                            check_status_of_test(driver, Title, device, campaign, usercampaignsname,run_test_status_value,result_status)
                    elif (test_complete == [] or test_complete == [False]) and (test_Selected == [False] or test_Selected == []):
                        statement = "Test data was not entered successfully for a particular type of test,so clicked on the close button."
                        with allure.step(statement):
                            allure.attach(driver.get_screenshot_as_png(), name=f"{statement}",attachment_type=allure.attachment_type.PNG)
                            status_df = status(Title=Title, component=f"device:- {device},campaign:- {campaign},usercampaignsname:-{usercampaignsname}",status="FAILED", comments=statement)
                            result_status.put(status_df)
                            click(driver=driver, locators=schedule_test.close_button)
                        time.sleep(0.1)
    except Exception as e:
        pass
    finally:
        return device

verifying_of_test_execution_for_scheduletest= verifying_of_test_execution_for_runtest

def creating_excel_file_for_contiunoustest_for_each_campaigns_generated(original_file_path, usercampaignsname_list,copied_file_paths):
    # Check if the original file exists
    if os.path.isfile(original_file_path):
        # Get the directory and base filename without extension
        directory, base_filename = os.path.split(original_file_path)
        filename, extension = os.path.splitext(base_filename)

        # Iterate over items and make copies
        for item in usercampaignsname_list[:-1]:
            # Create a new filename with the item appended
            new_filename = f"{filename}_{item}_{extension}"

            # Construct the full path for the new file
            new_file_path = os.path.join(directory, new_filename)

            # Copy the original file to the new path
            shutil.copy(original_file_path, new_file_path)

            # Append the new file path to the list
            copied_file_paths.append(new_file_path)

        # Rename the original file with the last item from the list
        last_item = usercampaignsname_list[-1]
        renamed_original_file = os.path.join(directory, f"{filename}_{last_item}_{extension}")
        os.rename(original_file_path, renamed_original_file)

        # Append the renamed original file path to the list
        copied_file_paths.append(renamed_original_file)

        print("Files copied and renamed successfully.")
        print("List of file paths:", copied_file_paths)
    else:
        print("Original file not found.")

def multi_bparty_call_test(driver,Title,test_data,result_status):
    try:
        with allure.step("Sequential Multi BParty Number Call Test"):
            clickec(driver=driver, locators=continuous.multi_bparty_checkbox)
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located(continuous.multi_bparty_test_form))
            multi_bparty_Call_test = None
            multi_bparty_Call_test_form_fields_names = driver.find_elements(*continuous.multi_bparty_test_form)
            multi_bparty_Call_test = True
            for multi_bparty_Call_test_form_fields_name in multi_bparty_Call_test_form_fields_names:
                try:
                    field_name = str(multi_bparty_Call_test_form_fields_name.text).lower().replace(" ", "")
                    if re.search("B Party Phone Number".lower().replace(" ", ""), field_name, re.IGNORECASE):
                        try:
                            multi_bparty_call_data = test_data["Multi B Party Phone Number"]
                        except Exception as e:
                            statement = f"Check parameter is present in remote test sheet test data excel"
                            with allure.step(statement):
                                status_df = status(Title=Title, component=f"Multi BParty Number Call Test:-'Multi B Party Phone Number'",status="FAILED", comments=statement)
                                result_status.put(status_df)
                                raise e
                        inputtext(driver=multi_bparty_Call_test_form_fields_name,locators=continuous.multi_bparty_phone_number, value=multi_bparty_call_data)
                    elif re.search("Call Duration".lower().replace(" ", ""), field_name, re.IGNORECASE):
                        try:
                            multi_bparty_call_data = test_data["Call Duration"]
                        except Exception as e:
                            statement = f"Check parameter is present in remote_test sheet test data excel"
                            with allure.step(statement):
                                status_df = status(Title=Title, component=f"Multi BParty Number Call Test:-'Call Duration'",status="FAILED", comments=statement)
                                result_status.put(status_df)
                                raise e
                        if is_numeric(multi_bparty_call_data):
                            if 1 <= int(multi_bparty_call_data) <= 90:
                                inputtext(driver=multi_bparty_Call_test_form_fields_name, locators=continuous.call_duration,value=multi_bparty_call_data)
                            else:
                                statement = f"Check the value given within the range"
                                with allure.step(statement):
                                    status_df = status(Title=Title, component=f"Multi BParty Number Call Test:-'Call Duration'",status="FAILED", comments=statement)
                                    result_status.put(status_df)
                                    e = Exception
                                    raise e
                        else:
                            statement = f"Check the entered value is numeric"
                            with allure.step(statement):
                                status_df = status(Title=Title, component=f"Multi BParty Number Call Test:-'Call Duration'",status="FAILED", comments=statement)
                                result_status.put(status_df)
                                e = Exception
                                raise e
                except Exception as e:
                    multi_bparty_Call_test = False
                    continue
            allure.attach(driver.get_screenshot_as_png(), name=f"Multi BParty Number Call Test.",attachment_type=allure.attachment_type.PNG)
            if multi_bparty_Call_test == True:
                okbtn = driver.find_element(*continuous.multi_bparty_ok_btn[:2])
                if okbtn.is_enabled():
                    clickec(driver=driver, locators=continuous.multi_bparty_ok_btn)
                elif not okbtn.is_enabled():
                    clickec(driver=driver, locators=continuous.multi_bparty_closebtn)
                    e = Exception
                    raise e
            elif multi_bparty_Call_test == False or multi_bparty_Call_test == None:
                clickec(driver=driver, locators=continuous.multi_bparty_closebtn)
                e = Exception
                raise e
    except Exception as e:
        raise e
