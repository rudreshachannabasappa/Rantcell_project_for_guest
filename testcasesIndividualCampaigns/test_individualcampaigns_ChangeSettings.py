
import os, allure, pytest, datetime
from configurations.config import ReadConfig as config
from pageobjects.remote_test import *
from utils.createxl import create_workbook
from pageobjects.settings__dash import *
from utils.readexcel import *
from pageobjects.login_logout import *
from pageobjects.Dashboard import *
from utils.updateexcelfile import *
from utils.library import *

class Test_Campaign_ChangeSettings_Driver:
    driver = None
    keys_to_check = ['Map_View', 'Graph_View', 'Exports', 'PDF Data Export and Validation',
                     'Table Summary Export Validation', 'NW Freeze',
                     'Combined Export Data Validation', 'Individual Popup window data validation',
                     'NQC table data validation','Default Settings','Change Settings']
    androidtest = side_bar_to_run_for_androidtest(keys_to_check)
    print(str(androidtest))
    remotetest_runvalue = Testrun_mode(value="Remote Test")
    scheduletest_runvalue = Testrun_mode(value="Schedule test")
    continuoustest_runvalue = Testrun_mode(value="Continuous Test")
    group_runvalue = Testrun_mode(value="Group")
    Default_settings_runvalue = Testrun_mode(value="Default Settings")
    Change_settings_runvalue = Testrun_mode(value="Change Settings")
    # ////////////////////////////////////CHANGE/////////////////////////////////////////////
    if ("FINISHED".lower() == remotetest_runvalue[-1].strip().lower() or "FINISHED".lower() == scheduletest_runvalue[-1].strip().lower() or "FINISHED".lower() == continuoustest_runvalue[-1].strip().lower()) and ((androidtest == True or "Yes".lower() == group_runvalue[-1].strip().lower()) and ("Yes".lower() == Default_settings_runvalue[-1].strip().lower() or "RUNNED".lower() == Change_settings_runvalue[-1].strip().lower())):
        @pytest.mark.parametrize("device,campaign,usercampaignsname,testgroup,environment,url,userid,password,excelreportfilepath,testdownloadpath",fetch_camapaigns_enviroment())
        def test_campaign_created_in_remotetest_and_also_for_changesettings(self, setup, device, campaign, usercampaignsname, testgroup, environment, url, userid,password, excelreportfilepath, testdownloadpath):
            global Excel_report_file_path, campaign1, campaign2
            driver, test_case_downloading_files_path = setup
            password = encrypte_decrypte(text=password)
            allure.dynamic.title(str(campaign))
            excelpath = excelreportfilepath
            protestdata_runvalue = Testrun_mode(value="Pro TestData")
            litetestdata_runvalue = Testrun_mode(value="LITE TestData")
            typeoftest = None
            if "Yes".lower() == protestdata_runvalue[-1].strip().lower():
                typeoftest = "ProTest data"
            elif "Yes".lower() == litetestdata_runvalue[-1].strip().lower():
                typeoftest = "LiteTest data"
            # Create XL file to capture data points for each component
            timestamp = datetime.datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
            timestamp1 = datetime.datetime.now().strftime("%d%m%Y%H%M%S")
            if compare_values(usercampaignsname, 'None'):
                campaign1 = campaign
                campaign2 = campaign
            elif not compare_values(usercampaignsname, 'None'):
                campaign1 = usercampaignsname
                campaign2 = usercampaignsname + campaign
                # runvalue = Testrun_mode(value="Remote Test")
                # if "Yes".lower() == runvalue[-1].strip().lower():
                #     campaign1 = campaign1 + timestamp1
            # Update Test Details in the Excel sheet
            updateiteration( campaign, device, environment, url, excelpath)
            add_headers_and_data(file_path=excelpath, headers=['Component Type', 'Data validation'],sheet_name="DATA_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['Component Type', 'Data validation'],sheet_name="DATA_NOT_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['Data validation'],sheet_name="TABLESUMMARY_DATA_NOT_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['Data validation'], sheet_name="TABLESUMMARY_DATA_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['File', 'ParameterType', 'Data validation'],sheet_name="CBE_vs_CE_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['File', 'ParameterType', 'Data validation'],sheet_name="CBE_vs_CE_DONOT_MATCH")
            add_headers_and_data(file_path=excelpath,headers=['File', "Individual pop up headers", "Individual pop up value","combine export value", 'Data validation'], sheet_name="IPU_vs_CE_DATA_MATCH")
            add_headers_and_data(file_path=excelpath,headers=['File', "Individual pop up headers", "Individual pop up value","combine export value", 'Data validation'],sheet_name="IPU_vs_CE_DATA_NOT_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['File', "map view Operator", "map view Operator value","calculated csv value", 'Data validation'],sheet_name="NQC_vs_OC_DATA_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['File', "map view Operator", "map view Operator value","calculated csv value", 'Data validation'],sheet_name="NQC_vs_OC_DATA_NOT_MATCH")
            list_of_testCases_Ignored = ["T010", "T051", "T092", "T133", "T174"]
            if str(campaign) in list_of_testCases_Ignored:
                statement = f"It was agreed not to perform test-case {str(campaign)} hence greyed-out in Test_Data.xlsx perform next 4 consecutive for same"
                with allure.step(statement):
                    Failupdatename(excelpath, statement)
                    updatecomponentstatus(f"Complete View of {str(campaign)}", str(campaign), "FAILED", statement,excelpath)
                    format_workbook(excelpath)
                    assert False
            else:
                # Launch browser and Navigate to RantCell Application LoginPage
                with allure.step("Launch and navigating to RantCell Application LoginPage"):
                    Navigate_to_loginPage(driver, url)

                # Login to RantCell Application
                with allure.step("Login to RantCell Application"):
                    login(driver, userid, password)

                keys_to_check = ['Map_View', 'Graph_View', 'Exports', 'PDF Data Export and Validation',
                                 'Table Summary Export Validation',
                                 'NW Freeze', 'Combined Export Data Validation',
                                 'Individual Popup window data validation', 'NQC table data validation',
                                 'Default Settings', 'Change Settings']
                androidtest = side_bar_to_run_for_androidtest(keys_to_check)
                if androidtest == True:
                    # Select campaign/classifier by navigating via sidebar menu : Android TestData -> ProTest data -> Device
                    with allure.step(f"Navigating to [Android TestData  >>>  {typeoftest}  >>>  Device  >>>  {str(campaign)}]"):
                        side_menu_Components_(driver, device, campaign1, userid, password, excelpath)

                    with allure.step(" Change Settings"):
                        main_change_settings(driver,campaign,environment,userid,excelpath)

                    # Select Map View Components
                    with allure.step("Select Map View Components"):
                        select_Map_View_Components_(driver, campaign, device, excelpath)

                    with allure.step("PDF export with operator comparsion"):
                        pdf_export_file_with_operator_comparsion_(driver, campaign, excelpath,test_case_downloading_files_path + "\\")

                    with allure.step("NQC table data vs operator comparsion"):
                        Nqc(driver, test_case_downloading_files_path, excelpath)
                try:
                    finishcomponentstatus_test_case_((campaign2 + device + environment), excelpath)
                except Exception as e:
                    pass
                # Logout from RantCell Application
                with allure.step("Logout to RantCell Application"):
                    logout(driver)

                # folder_path = './target/allure-results'
                # compress_png_files(folder_path)

                # Read the component statues from Excel Report
                status = readcomponentstatus(excelpath)
                format_workbook(excelpath)

                # Mark the test case as failed if any component is field
                if 'FAILED' not in status:
                    driver.quit()
                    assert True
                else:
                    driver.quit()
                    assert False

    elif ((androidtest == True or "Yes".lower() == group_runvalue[-1].strip().lower()) and ("Yes".lower() == Default_settings_runvalue[-1].strip().lower() or "RUNNED".lower() == Change_settings_runvalue[-1].strip().lower())) and ("No".lower() == remotetest_runvalue[-1].strip().lower() and "No".lower() == scheduletest_runvalue[-1].strip().lower() and "No".lower() == continuoustest_runvalue[-1].strip().lower()):
        @pytest.mark.parametrize("device,campaign,usercampaignsname,testgroup,environment,url,userid,password,excelreportfilepath,testdownloadpath",fetch_camapaigns_enviroment())
        def test_campaign_already_created_and_also_for_changesettings(self, setup, device, campaign, usercampaignsname, testgroup, environment, url, userid,password, excelreportfilepath, testdownloadpath):
            global Excel_report_file_path, campaign1, campaign2
            driver, test_case_downloading_files_path = setup
            password = encrypte_decrypte(text=password)
            allure.dynamic.title(str(campaign))
            excelpath = excelreportfilepath
            protestdata_runvalue = Testrun_mode(value="Pro TestData")
            litetestdata_runvalue = Testrun_mode(value="LITE TestData")
            typeoftest = None
            if "Yes".lower() == protestdata_runvalue[-1].strip().lower():
                typeoftest = "ProTest data"
            elif "Yes".lower() == litetestdata_runvalue[-1].strip().lower():
                typeoftest = "LiteTest data"
            # Create XL file to capture data points for each component
            timestamp = datetime.datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
            timestamp1 = datetime.datetime.now().strftime("%d%m%Y%H%M%S")
            if compare_values(usercampaignsname, 'None'):
                campaign1 = campaign
                campaign2 = campaign
            elif not compare_values(usercampaignsname, 'None'):
                campaign1 = usercampaignsname
                campaign2 = usercampaignsname + campaign
                # runvalue = Testrun_mode(value="Remote Test")
                # if "Yes".lower() == runvalue[-1].strip().lower():
                #     campaign1 = campaign1 + timestamp1
            # Update Test Details in the Excel sheet
            updateiteration( campaign, device, environment, url, excelpath)
            startcomponentstatus_test_case_((campaign2 + device + environment), excelpath)
            add_headers_and_data(file_path=excelpath,headers = ['Title', 'Componentname', 'Status', 'Comments'],sheet_name='COMPONENTSTATUS')
            add_headers_and_data(file_path=excelpath, headers=['Component Type', 'Data validation'],sheet_name="DATA_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['Component Type', 'Data validation'],sheet_name="DATA_NOT_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['Data validation'], sheet_name="TABLESUMMARY_DATA_NOT_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['Data validation'], sheet_name="TABLESUMMARY_DATA_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['File', 'ParameterType', 'Data validation'],sheet_name="CBE_vs_CE_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['File', 'ParameterType', 'Data validation'],sheet_name="CBE_vs_CE_DONOT_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['File',"Individual pop up headers","Individual pop up value","combine export value", 'Data validation'],sheet_name="IPU_vs_CE_DATA_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['File',"Individual pop up headers","Individual pop up value","combine export value", 'Data validation'],sheet_name="IPU_vs_CE_DATA_NOT_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['File',"map view Operator","map view Operator value","calculated csv value",'Data validation'],sheet_name="NQC_vs_OC_DATA_MATCH")
            add_headers_and_data(file_path=excelpath, headers=['File',"map view Operator","map view Operator value","calculated csv value",'Data validation'],sheet_name="NQC_vs_OC_DATA_NOT_MATCH")
            list_of_testCases_Ignored = ["T010", "T051", "T092", "T133", "T174"]
            if str(campaign) in list_of_testCases_Ignored:
                statement = f"It was agreed not to perform test-case {str(campaign)} hence greyed-out in Test_Data.xlsx perform next 4 consecutive for same"
                with allure.step(statement):
                    Failupdatename(excelpath, statement)
                    updatecomponentstatus(f"Complete View of {str(campaign)}", str(campaign), "FAILED", statement,excelpath)
                    format_workbook(excelpath)
                    assert False
            else:
                # Launch browser and Navigate to RantCell Application LoginPage
                with allure.step("Launch and navigating to RantCell Application LoginPage"):
                    Navigate_to_loginPage(driver, url)

                # Login to RantCell Application
                with allure.step("Login to RantCell Application"):
                    login(driver, userid, password)

                keys_to_check = ['Map_View', 'Graph_View', 'Exports', 'PDF Data Export and Validation',
                                 'Table Summary Export Validation',
                                 'NW Freeze', 'Combined Export Data Validation',
                                 'Individual Popup window data validation', 'NQC table data validation',
                                 'Default Settings', 'Change Settings']
                androidtest = side_bar_to_run_for_androidtest(keys_to_check)
                if androidtest == True:
                    # Select campaign/classifier by navigating via sidebar menu : Android TestData -> ProTest data -> Device
                    with allure.step(f"Navigating to [Android TestData  >>> {typeoftest} >>>  Device  >>>  {str(campaign)}]"):
                        side_menu_Components_(driver, device, campaign1, userid, password, excelpath)

                    with allure.step(" Change Settings"):
                        main_change_settings(driver,campaign,environment,userid,excelpath)

                    # Select Map View Components
                    with allure.step("Select Map View Components"):
                        select_Map_View_Components_(driver, campaign, device, excelpath)

                    with allure.step("PDF export with operator comparsion"):
                        pdf_export_file_with_operator_comparsion_(driver, campaign, excelpath,test_case_downloading_files_path + "\\")

                    with allure.step("NQC table data vs operator comparsion"):
                        Nqc(driver, test_case_downloading_files_path, excelpath)
                try:
                    finishcomponentstatus_test_case_((campaign2 + device + environment), excelpath)
                except Exception as e:
                    pass
                # Logout from RantCell Application
                with allure.step("Logout to RantCell Application"):
                    logout(driver)

                # folder_path = './target/allure-results'
                # compress_png_files(folder_path)

                # Read the component statues from Excel Report
                status = readcomponentstatus(excelpath)
                format_workbook(excelpath)

                # Mark the test case as failed if any component is field
                if 'FAILED' not in status:
                    driver.quit()
                    assert True
                else:
                    driver.quit()
                    assert False