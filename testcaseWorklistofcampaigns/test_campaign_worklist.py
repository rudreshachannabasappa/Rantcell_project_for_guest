import os, allure, pytest, datetime
from configurations.config import ReadConfig as config
from pageobjects.remote_test import *
from utils.createxl import create_workbook
from utils.readexcel import *
from pageobjects.login_logout import *
from pageobjects.Dashboard import *
from utils.updateexcelfile import *
from utils.library import *

class Test_List_Of_Campaign_Driver:
    driver = None
    @pytest.mark.parametrize("environment,url,userid,password",fetch_enviroment())
    def test_campaign_list_of_campaigns(self,setup,environment,url,userid,password):
        global Excel_report_file_path
        driver, test_case_downloading_files_path = setup
        f1 = open(config.test_run_excelreportdata_path, "r")
        testrunexcelfolder = f1.read()
        f1.close()
        password = encrypte_decrypte(text=password)
        protestdata_runvalue = Testrun_mode(value="Pro TestData")
        litetestdata_runvalue = Testrun_mode(value="LITE TestData")
        typeoftest = None
        if "Yes".lower() == protestdata_runvalue[-1].strip().lower():
            typeoftest = "ProTest data"
        elif "Yes".lower() == litetestdata_runvalue[-1].strip().lower():
            typeoftest = "LiteTest data"
        # Create XL file to capture data points for each component
        timestamp = datetime.datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
        try:
            Excel_report_file_path = config.excel_report_path + testrunexcelfolder
            if os.path.exists(Excel_report_file_path):
                print("test run excel folder is exist")
            if not os.path.exists(Excel_report_file_path):
                pytest.fail("test run excel folder is not exist")
            excelpath = Excel_report_file_path +"\\"+"Worklist_of_campaigns"+"_" + environment + timestamp +".xlsx"
            create_workbook(excelpath)
        except Exception as e:
            with allure.step(f"Check {Excel_report_file_path}{e}"):
                print(f"Check {Excel_report_file_path}{e}")
                assert False
        campaigns_data = fetch_camapaigns()
        # Update Test Details in the Excel sheet
        startcomponentstatus_test_case_(("Worklist_of_campaigns" + environment), excelpath)
        add_headers_and_data(file_path=excelpath, headers=['Title', 'Componentname', 'Status', 'Comments'],sheet_name='COMPONENTSTATUS')
        add_headers_and_data(file_path=excelpath, headers=['Component Type', 'Data validation'],sheet_name="DATA_MATCH")
        add_headers_and_data(file_path=excelpath, headers=['Component Type', 'Data validation'],sheet_name="DATA_NOT_MATCH")
        add_headers_and_data(file_path=excelpath, headers=['Data validation'],sheet_name="TABLESUMMARY_DATA_NOT_MATCH")
        add_headers_and_data(file_path=excelpath, headers=['Data validation'],sheet_name="TABLESUMMARY_DATA_MATCH")
        add_headers_and_data(file_path=excelpath, headers=['File', 'ParameterType', 'Data validation'],sheet_name="CBE_vs_CE_MATCH")
        add_headers_and_data(file_path=excelpath, headers=['File', 'ParameterType', 'Data validation'],sheet_name="CBE_vs_CE_DONOT_MATCH")
        add_headers_and_data(file_path=excelpath, headers=['File', "Individual pop up headers", "Individual pop up value","combine export value", 'Data validation'],sheet_name="IPU_vs_CE_DATA_MATCH")
        add_headers_and_data(file_path=excelpath, headers=['File', "Individual pop up headers", "Individual pop up value","combine export value", 'Data validation'],sheet_name="IPU_vs_CE_DATA_NOT_MATCH")
        add_headers_and_data(file_path=excelpath, headers=['File', "map view Operator", "map view Operator value","calculated csv value", 'Data validation'],sheet_name="NQC_vs_OC_DATA_MATCH")
        add_headers_and_data(file_path=excelpath, headers=['File', "map view Operator", "map view Operator value","calculated csv value", 'Data validation'],sheet_name="NQC_vs_OC_DATA_NOT_MATCH")

        # Launch browser and Navigate to RantCell Application LoginPage
        with allure.step("Launch and navigating to RantCell Application LoginPage"):
            Navigate_to_loginPage(driver, url)

        # Login to RantCell Application
        with allure.step("Login to RantCell Application"):
            login(driver, userid, password)

        with allure.step(f"List Of Campaigns ==> {typeoftest}"):
            side_bar_menu_for_work_list_campaigns(driver,userid, password,campaigns_data,excelpath)

        with allure.step("Select Map View Components"):
            Map_view_for_work_list_campaigns(driver,campaigns_data,excelpath)

        # Select Graph View Components
        with allure.step("Select Graph View Components"):
            Graph_View_for_work_list_campaigns(driver, campaigns_data, excelpath)

        # Select Table View Components
        with allure.step("Expand List of Campaign's Table-View and Verify Pop-Up"):
            expand_tableView_verify_popUp_(driver)

        # Download CSV files from Exports
        with allure.step("List Of Campaigns Export in Dashboard"):
            Exports_view_work_list_campaigns(driver, excelpath, campaigns_data, test_case_downloading_files_path + "\\")

        with allure.step("Table summary data validation"):
            table_summary_(driver, downloadfilespath=test_case_downloading_files_path + "\\", excelpath=excelpath)

        with allure.step("Combine export vs Combine binary export"):
            combine_export_vs_combine_binary_export(driver, test_case_downloading_files_path, excelpath)

        with allure.step("PDF export with operator comparsion"):
            pdf_export_for_work_list_campaigns(driver, campaigns_data, excelpath,test_case_downloading_files_path + "\\")

        with allure.step("Combine binary export nw freeze"):
            combine_binary_export_nw_freeze(driver, excelpath, test_case_downloading_files_path + "\\")

        with allure.step("individual pop up table vs Combine export"):
            individual_popup_table_vs_ce(driver, test_case_downloading_files_path, excelpath)

        with allure.step("NQC table data vs operator comparsion"):
            Nqc(driver, test_case_downloading_files_path, excelpath)

        try:
            finishcomponentstatus_test_case_(("Datetime_query"+environment), excelpath)
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
