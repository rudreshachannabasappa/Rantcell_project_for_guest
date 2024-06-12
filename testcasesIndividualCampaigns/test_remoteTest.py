import os, allure, pytest, datetime
from configurations.config import ReadConfig as config
from pageobjects.remote_test import *
from utils.readexcel import *
from pageobjects.login_logout import *
from pageobjects.Dashboard import *
from utils.updateexcelfile import *
from utils.createxl import create_workbook
from utils.library import *

# if Test_Data.xlsx Excel file is not present in the path mentioned in the config file 'config.json',then execution will be stopped
class Test_RemoteTest_Driver:
    driver = None
    remotetest_runvalue = Testrun_mode(value="Remote Test")
    group_runvalue = Testrun_mode(value="Group")
    scheduletest_runvalue = Testrun_mode(value="Schedule test")
    continuoustest_runvalue = Testrun_mode(value="Continuous Test")
    if "Yes".lower() == remotetest_runvalue[-1].strip().lower() or "Yes".lower() == scheduletest_runvalue[-1].strip().lower() or "Yes".lower() == continuoustest_runvalue[-1].strip().lower():
        @pytest.mark.parametrize("device,campaign,usercampaignsname,testgroup,environment,url,userid,password",fetch_camapaigns_enviroment())
        def test_remotetest(self,setup, device, campaign,usercampaignsname,testgroup,environment, url, userid, password):
            global Excel_report_file_path,campaign1,campaign2,continuous_campaigns
            driver,test_case_downloading_files_path= setup
            continuous_campaigns="None"
            campaign_scheduletest = "None"
            campaign_continuoustest = "None"
            campaign_runtest = "None"
            campaigns_created = []
            f1 = open(config.test_run_excelreportdata_path, "r")
            testrunexcelfolder = f1.read()
            f1.close()
            password = encrypte_decrypte(text=password)
            allure.dynamic.title(str(campaign))
            remotetest_runvalue = Testrun_mode(value="Remote Test")
            scheduletest_runvalue = Testrun_mode(value="Schedule test")
            continuoustest_runvalue = Testrun_mode(value="Continuous Test")
            keys_to_check = ['Map_View', 'Graph_View', 'Exports', 'PDF Data Export and Validation',
                             'Table Summary Export Validation', 'NW Freeze',
                             'Combined Export Data Validation', 'Individual Popup window data validation',
                             'NQC table data validation']
            androidtest = side_bar_to_run_for_androidtest(keys_to_check)
            # Create XL file to capture data points for each component
            timestamp = datetime.datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
            timestamp1 = datetime.datetime.now().strftime("%d%m%Y%H%M%S")
            if compare_values(usercampaignsname, 'None'):
                campaign1 = campaign
                campaign2 = campaign
            elif not compare_values(usercampaignsname, 'None'):
                campaign1 = usercampaignsname
                campaign2 = usercampaignsname + campaign
                runvalue = Testrun_mode(value="Remote Test")
                if "Yes".lower() == runvalue[-1].strip().lower():
                    campaign_runtest = campaign1 + timestamp1
                if "Yes".lower() == continuoustest_runvalue[-1].strip().lower():
                    random_length = random.randint(3, 5)
                    random_alphabet = generate_random_alphabet(random_length)
                    random_alphabet = random_alphabet.upper()
                    campaign_continuoustest = random_alphabet + "_" + campaign1
                    continuous_campaigns = random_alphabet + "_"
                if "Yes".lower() == scheduletest_runvalue[-1].strip().lower():
                    campaign_scheduletest = campaign1 + timestamp1 + "schedule"
            try:
                Excel_report_file_path = config.excel_report_path + testrunexcelfolder
                if os.path.exists(Excel_report_file_path):
                    print("test run excel folder is exist")
                if not os.path.exists(Excel_report_file_path):
                    pytest.fail("test run excel folder is not exist")
                excelpath = Excel_report_file_path +"\\"+ campaign2 + "_" + device + "_" + environment + timestamp +".xlsx"
                create_workbook(excelpath)
            except Exception as e:
                with allure.step(f"Check {Excel_report_file_path}{e}"):
                    print(f"Check {Excel_report_file_path}{e}")
                    assert False
            # Fetch components based on the campaign/classifier "T001","T002" etc
            remote_test_point, map_start_point, graph_start_point, export_start_point, load_start_point, PDF_Export_index_start_point, END_index = fetch_input_points()
            tests = fetch_components(campaign, remote_test_point, map_start_point)
            # Update Test Details in the Excel sheet
            startcomponentstatus_test_case_((campaign2 + device + environment), excelpath)
            add_headers_and_data(file_path=excelpath,headers = ['Title', 'Componentname', 'Status', 'Comments'],sheet_name='COMPONENTSTATUS')
            list_of_testCases_Ignored = ["T010", "T051", "T092", "T133", "T174"]
            if str(campaign) in list_of_testCases_Ignored:
                statement = f"It was agreed not to perform test-case {str(campaign)} hence greyed-out in Test_Data.xlsx perform next 4 consecutive for same"
                with allure.step(statement):
                    Failupdatename(excelpath, statement)
                    updatecomponentstatus(f"Complete View of {str(campaign)}", str(campaign), "FAILED", statement, excelpath)
                    format_workbook(excelpath)
                    assert False
            else:
                # Launch browser and Navigate to RantCell Application LoginPage
                with allure.step("Launch and navigating to RantCell Application LoginPage"):
                    Navigate_to_loginPage(driver, url)

                # Login to RantCell Application
                with allure.step("Login to RantCell Application"):
                    login(driver, userid, password)

                with allure.step("Schedule Test"):
                    device = Schedule_test_(driver,device,campaign,campaigns_created,usercampaignsname=campaign_scheduletest,testgroup=testgroup,tests=tests,excelpath= excelpath)

                with allure.step("Continuous Test"):
                    device = Continuous_Test_(driver,device,campaign,campaigns_created,usercampaignsname=campaign_continuoustest,testgroup=testgroup,continuous_campaigns = continuous_campaigns,tests=tests,excelpath= excelpath)

                with allure.step("Remote Test"):
                    device = remote_test_(driver,device,campaign,campaigns_created,usercampaignsname=campaign_runtest,testgroup=testgroup,tests=tests,excelpath= excelpath)

                if ("Yes".lower() == remotetest_runvalue[-1].strip().lower() or "Yes".lower() == scheduletest_runvalue[-1].strip().lower() or "Yes".lower() == continuoustest_runvalue[-1].strip().lower()) and androidtest == False:
                    try:
                        finishcomponentstatus_test_case_((campaign2 + device + environment), excelpath)
                    except Exception as e:
                        pass

                copied_file_paths = []
                creating_excel_file_for_contiunoustest_for_each_campaigns_generated(original_file_path=excelpath,usercampaignsname_list=campaigns_created,copied_file_paths=copied_file_paths)
                for usercampaignsname2,excelpath1 in zip(campaigns_created,copied_file_paths):
                    # campaigns_created.append(usercampaignsname2)
                    automation_data_dict = {"DEVICES":[device],"CAMPAIGNS":[campaign],"EXECUTE":["Yes"],"USERCAMPAIGNSNAME":[usercampaignsname2],"TEST GROUP":[testgroup],"Environment":[environment],"URL":[url],"Login":[userid],"Password":[password],"Excel report file path":[excelpath1],"Test Download path":[test_case_downloading_files_path]}
                    update_automation_data(automation_data_dict=automation_data_dict, automation_data_execel_path=config.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx",Sheet="AUTOMATION_DATA")
                    excelpath = excelpath1

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





