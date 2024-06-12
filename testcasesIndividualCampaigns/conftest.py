import os
import queue
import pytest
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from pageobjects.remote_test import update_automation_data
from utils.createFolderforRantcell_automation_DataandReports import create_folder_for_rantcell_data_and_ExcelReport, \
    Updating_source_folder, \
    create_folder_for_excelreport, excel_report_path_, testRun_downloadfile_path, \
    create_folder_for_downloads, text_file_update_for_remote_test
from configurations.config import ReadConfig
from utils.library import *
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
# Counter for active parallel threads
active_threads = 0
threads =0
Remotetest_runvalue = Testrun_mode(value="Remote Test")
scheduletest_runvalue = Testrun_mode(value="Schedule test")
continuoustest_runvalue = Testrun_mode(value="Continuous Test")
Change_Settingsrunvalue = Testrun_mode(value="Change Settings")
default_Settingsrunvalue = Testrun_mode(value="Default Settings")
# time.sleep(10)
def pytest_cmdline_main(config):
    target_argument = 'test_remoteTest.py'
    args = config.args
    inicfg = config.inicfg
    def create_workbook_for_automation_data(path):
        workbook = Workbook()
        workbook.create_sheet("AUTOMATION_DATA", 0)
        workbook.create_sheet("CHANGE AUTOMATION_DATA", 1)
        workbook.save(path)
    for arg in args:
        if target_argument in arg:
            if "RUNNED".lower() != Remotetest_runvalue[-1].strip().lower() and "RUNNED".lower() != scheduletest_runvalue[-1].strip().lower() and "RUNNED".lower() != continuoustest_runvalue[-1].strip().lower() :
                try:
                    if os.path.exists(ReadConfig.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx"):
                        os.remove(ReadConfig.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx")
                except Exception as e:
                    pass
                # Check if the -n option is present in the command line
                if "Yes".lower() == Remotetest_runvalue[-1].strip().lower() or "Yes".lower() == scheduletest_runvalue[-1].strip().lower() or "Yes".lower() == continuoustest_runvalue[-1].strip().lower():
                    if "-n" in inicfg:
                        inicfg['addopts'] = '-n 1'
                    if "Yes".lower() == Remotetest_runvalue[-1].strip().lower() or "Yes".lower() == scheduletest_runvalue[-1].strip().lower() or "Yes".lower() == continuoustest_runvalue[-1].strip().lower():
                        create_workbook_for_automation_data(path=ReadConfig.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx")
                        add_headers_and_data(file_path=ReadConfig.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx",headers=["DEVICES", "CAMPAIGNS", "EXECUTE", "USERCAMPAIGNSNAME", "TEST GROUP","Environment", "URL", "Login", "Password", "Excel report file path","Test Download path"], sheet_name="AUTOMATION_DATA")
                        add_headers_and_data(file_path=ReadConfig.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx",headers=["DEVICES", "CAMPAIGNS", "EXECUTE", "USERCAMPAIGNSNAME", "TEST GROUP","Environment", "URL", "Login", "Password", "Excel report file path","Test Download path"], sheet_name="CHANGE AUTOMATION_DATA")
                elif "No".lower() == Remotetest_runvalue[-1].strip().lower() and "No".lower() == scheduletest_runvalue[-1].strip().lower() and "No".lower() == continuoustest_runvalue[-1].strip().lower():
                    pytest.exit("Test is not to execute")
            break
    if ("RUNNED".lower() == default_Settingsrunvalue[-1].strip().lower() or "WAITING LOAD".lower() == Change_Settingsrunvalue[-1].strip().lower()):
        target_argument = 'test_individualcampaigns.py'
        for arg in args:
            if target_argument in arg:
                if ("RUNNED".lower() == default_Settingsrunvalue[-1].strip().lower() or "WAITING LOAD".lower() ==Change_Settingsrunvalue[-1].strip().lower()) and ("RUNNED".lower() == Remotetest_runvalue[-1].strip().lower() or "RUNNED".lower() == scheduletest_runvalue[-1].strip().lower() or "RUNNED".lower() == continuoustest_runvalue[-1].strip().lower()):
                    pass
                    # add_headers_and_data(file_path=ReadConfig.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx",headers=["DEVICES", "CAMPAIGNS", "EXECUTE", "USERCAMPAIGNSNAME", "TEST GROUP", "Environment", "URL","Login", "Password", "Excel report file path", "Test Download path"],sheet_name="CHANGE AUTOMATION_DATA")
                elif ("RUNNED".lower() == default_Settingsrunvalue[-1].strip().lower() or "WAITING LOAD".lower() ==Change_Settingsrunvalue[-1].strip().lower()):
                    try:
                        if os.path.exists(ReadConfig.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx"):
                            os.remove(ReadConfig.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx")
                    except Exception as e:
                        pass
                    time.sleep(1)
                    create_workbook_for_automation_data(path=ReadConfig.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx")
                    add_headers_and_data(file_path=ReadConfig.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx",headers=["DEVICES", "CAMPAIGNS", "EXECUTE", "USERCAMPAIGNSNAME", "TEST GROUP", "Environment","URL", "Login", "Password", "Excel report file path", "Test Download path"],sheet_name="CHANGE AUTOMATION_DATA")
                print("RUNNED",str(arg))
                break
    else:
        pass

def pytest_sessionstart(session):
    if active_threads == 0:
        try:
            if not os.path.exists(ReadConfig.test_data_folder_rootpath):
                e = Exception
                raise e
            if (os.path.exists(ReadConfig.test_data_folder_rootpath)) and (Remotetest_runvalue[-1].strip().lower() != 'RUNNED'.strip().lower() and "RUNNED".lower() != scheduletest_runvalue[-1].strip().lower() and "RUNNED".lower() != continuoustest_runvalue[-1].strip().lower()) and (Remotetest_runvalue[-1].strip().lower() != 'FINISHED'.strip().lower() and "FINISHED".lower() != scheduletest_runvalue[-1].strip().lower() and "FINISHED".lower() != continuoustest_runvalue[-1].strip().lower()) and ("RUNNED".lower() != Change_Settingsrunvalue[-1].strip().lower()):
                try:
                    create_folder_for_rantcell_data_and_ExcelReport(ReadConfig.test_data_folder_rootpath, ReadConfig.source_dest)
                except Exception:
                    print(str(Exception))
                finally:
                    testRun_downloadfile_path(ReadConfig.test_run_download_file_path)
                    time.sleep(2)
                    excel_report_path_(ReadConfig.test_run_excelreportdata_path)
                    f1 = open(ReadConfig.test_run_download_file_path, "r")
                    testrundownloadfolder = f1.read()
                    f1.close()
                    test_case_downloading_files_path_timestamps = ReadConfig.test_case_downloading_files_path_timestamp + testrundownloadfolder
                    create_folder_for_downloads(destination_folder=test_case_downloading_files_path_timestamps)
                    f1 = open(ReadConfig.test_run_excelreportdata_path, "r")
                    testrunexcelfolder = f1.read()
                    f1.close()
                    test_run_excel_report_pathtimestamp = ReadConfig.excel_report_path + testrunexcelfolder
                    create_folder_for_excelreport(destination_folder=test_run_excel_report_pathtimestamp)
                    Updating_source_folder(ReadConfig.updating_source_folders, ReadConfig.test_data_folder_rootpath)
            else:
                pass
        except Exception as e:
            with allure.step(f"Goto the directory {ReadConfig.test_data_folder_rootpath} and update the file as per your requirement, then re-run"):
                if not os.path.exists(ReadConfig.test_data_folder_rootpath):
                    create_folder_for_rantcell_data_and_ExcelReport(ReadConfig.test_data_folder_rootpath, ReadConfig.source_dest)
                    Updating_source_folder(ReadConfig.updating_source_folders, ReadConfig.test_data_folder_rootpath)
                    pytest.fail(f"Goto the directory {ReadConfig.test_data_folder_rootpath} and update the file as per your requirement")
def sanitize_folder_name(name):
    # Remove characters that are not valid for a folder name
    forbidden_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*','[',']','(',')','{','}']
    for char in forbidden_chars:
        name = name.replace(char, '')
    # Truncate or pad the name to have a length of 30
    sanitized_name = name[:40].ljust(40)
    return sanitized_name
@pytest.fixture(scope="function", autouse=True)
def check_previous_test_status(request):
    global active_threads
    # Get the previous item's outcome
    previous_outcome = request.session.testsfailed
    # If the previous test failed, skip the remaining tests
    if previous_outcome >= 1:
        Remotetest_runvalue = Testrun_mode(value="Remote Test")
        scheduletest_runvalue = Testrun_mode(value="Schedule test")
        continuoustest_runvalue = Testrun_mode(value="Continuous Test")
        if "Yes".lower() == Remotetest_runvalue[-1].strip().lower() or "Yes".lower() == scheduletest_runvalue[-1].strip().lower() or "Yes".lower() == continuoustest_runvalue[-1].strip().lower():
            active_threads = threads
            pytest.skip("Skipping remaining tests as a previous test failed")
@pytest.fixture(scope='function')
def setup(request):
    global driver,test_case_downloading_files_path
    chrome_options = Options()
    f1 = open(ReadConfig.test_run_download_file_path, "r")
    testrundownloadfolder = f1.read()
    f1.close()
    # Set the root path for test data and reports
    test_data_folder_rootpath = ReadConfig.test_case_downloading_files_path_timestamp + testrundownloadfolder
    if os.path.exists(test_data_folder_rootpath):
        print("test run downloading files folder path")
    if not os.path.exists(test_data_folder_rootpath):
        pytest.fail("test run downloading files folder path is not exist")
    if "Yes".lower() != Remotetest_runvalue[-1].strip().lower() and "Yes".lower() != scheduletest_runvalue[-1].strip().lower() and "Yes".lower() != continuoustest_runvalue[-1].strip().lower():
        random_length = random.randint(3, 5)
        random_alphabet = generate_random_alphabet(random_length)
        timestamp = datetime.datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
        if "RUNNED".lower() == continuoustest_runvalue[-1].strip().lower() or "FINISHED".lower() == continuoustest_runvalue[-1].strip().lower():
            test_case_name = sanitize_folder_name(request.node.name) + timestamp + random_alphabet
        else:
            test_case_name = sanitize_folder_name(request.node.name) + timestamp + random_alphabet
        test_case_downloading_files_path = test_data_folder_rootpath+"\\"+str(test_case_name)
        os.makedirs(test_case_downloading_files_path, exist_ok=True)
    if "Yes".lower() == Remotetest_runvalue[-1].strip().lower() or "Yes".lower() == scheduletest_runvalue[-1].strip().lower() or "Yes".lower() == continuoustest_runvalue[-1].strip().lower():
        test_case_downloading_files_path = test_data_folder_rootpath
    # Create the download folder path for the test case
    prefs = {
        'profile.default_content_setting_values.automatic_downloads': 1,
        "download.default_directory": test_case_downloading_files_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    #driver = webdriver.Chrome(executable_path=r'C:\\RantCell_Automation_Data_and_Reports\\Driver\\chromedriver.exe', options=chrome_options)
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()),options=chrome_options)
    global active_threads
    active_threads += 1
    yield driver,test_case_downloading_files_path
    driver.quit()

def pytest_collection_modifyitems(config, items):
    config.collected_items_count = len(items)
    global threads
    threads = config.collected_items_count
def pytest_sessionfinish(session):
    item_count = threads
    failed_count = session.testsfailed
    if active_threads == item_count:
        folder_path = config.test_data_folder_rootpath + "\\testdata\\"
        files = os.listdir(folder_path)
        xlsx_files = [file for file in files if file.endswith('.xlsx') and "Automationdata_" in file]
        automationdata = []
        for file in xlsx_files:
            automationpath = os.path.join(folder_path, file)
            automationdf = pd.read_excel(automationpath,sheet_name="CHANGE AUTOMATION_DATA")
            automationdata.append(automationdf)
            os.remove(automationpath)
        if len(automationdata)!=0:
            combine_df = pd.concat(automationdata,ignore_index=True)
            workbook = openpyxl.load_workbook(ReadConfig.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx")
            worksheet_automationdata = workbook["CHANGE AUTOMATION_DATA"]
            # Insert the DataFrame into the worksheet
            for index, row in combine_df.iterrows():
                worksheet_automationdata.append(row.tolist())
            workbook.save(ReadConfig.test_data_folder_rootpath + "\\testdata\\Automation_data.xlsx")
            workbook.close()
        if ("Yes".lower() == Remotetest_runvalue[-1].strip().lower() or "Yes".lower() == scheduletest_runvalue[-1].strip().lower() or "Yes".lower() == continuoustest_runvalue[-1].strip().lower()) and failed_count == 0 and item_count != 0 :
            if "Yes".lower() == Remotetest_runvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue="RUNNED",types_of_test="Remote Test")
            if "Yes".lower() == scheduletest_runvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue="RUNNED", types_of_test="Schedule Test")
            if "Yes".lower() == continuoustest_runvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue="RUNNED", types_of_test="Continuous Test")
        elif ("RUNNED".lower() == Remotetest_runvalue[-1].strip().lower() or "RUNNED".lower() == scheduletest_runvalue[-1].strip().lower() or "RUNNED".lower() == continuoustest_runvalue[-1].strip().lower()) :
            value = "Yes"
            if ("RUNNED".lower() == default_Settingsrunvalue[-1].strip().lower() or "WAITING LOAD".lower() == Change_Settingsrunvalue[-1].strip().lower()):
                value = "FINISHED"
            if "RUNNED".lower() == Remotetest_runvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue=value,types_of_test="Remote Test")
            if "RUNNED".lower() == scheduletest_runvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue=value, types_of_test="Schedule Test")
            if "RUNNED".lower() == continuoustest_runvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue=value, types_of_test="Continuous Test")
            if "RUNNED".lower() == default_Settingsrunvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue=value, types_of_test="Default Settings")
            if "WAITING LOAD".lower() == Change_Settingsrunvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue="LOADING", types_of_test="Change Settings")
        elif ("FINISHED".lower() == Remotetest_runvalue[-1].strip().lower() or "FINISHED".lower() == scheduletest_runvalue[-1].strip().lower() or "FINISHED".lower() == continuoustest_runvalue[-1].strip().lower()) and ("Yes".lower() == default_Settingsrunvalue[-1].strip().lower() or "RUNNED".lower() == Change_Settingsrunvalue[-1].strip().lower()):
            if "FINISHED".lower() == Remotetest_runvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue="Yes",types_of_test="Remote Test")
            if "FINISHED".lower() == scheduletest_runvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue="Yes", types_of_test="Schedule Test")
            if "FINISHED".lower() == continuoustest_runvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue="Yes", types_of_test="Continuous Test")
            if "RUNNED".lower() == Change_Settingsrunvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue="Yes", types_of_test="Change Settings")
        elif ("No".lower() == Remotetest_runvalue[-1].strip().lower() or "No".lower() == scheduletest_runvalue[-1].strip().lower() or "No".lower() == continuoustest_runvalue[-1].strip().lower()) and ("RUNNED".lower() == default_Settingsrunvalue[-1].strip().lower() or "WAITING LOAD".lower() == Change_Settingsrunvalue[-1].strip().lower()):
            if "RUNNED".lower() == default_Settingsrunvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue="FINISHED", types_of_test="Default Settings")
            if "WAITING LOAD".lower() == Change_Settingsrunvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue="LOADING", types_of_test="Change Settings") #----not worked
        elif ("No".lower() == Remotetest_runvalue[-1].strip().lower() or "No".lower() == scheduletest_runvalue[-1].strip().lower() or "No".lower() == continuoustest_runvalue[-1].strip().lower()) and ("Yes".lower() == default_Settingsrunvalue[-1].strip().lower() or "RUNNED".lower() == Change_Settingsrunvalue[-1].strip().lower()):
            if "RUNNED".lower() == Change_Settingsrunvalue[-1].strip().lower():
                updating_yes_to_run(remotevalue="Yes", types_of_test="Change Settings")

    Remotetest_runvalue1 = Testrun_mode(value="Remote Test")
    scheduletest_runvalue1 = Testrun_mode(value="Schedule test")
    continuoustest_runvalue1 = Testrun_mode(value="Continuous Test")
    Change_Settingsrunvalue1 = Testrun_mode(value="Change Settings")
    default_Settingsrunvalue1 = Testrun_mode(value="Default Settings")
    compare_same_hld_of_diffrentserver_runvalue1 = Testrun_mode(value="Comparative Analysis of two Server Results Within a Single HLD Excel Report")
    if active_threads == item_count and (Remotetest_runvalue1[-1].strip().lower() != 'RUNNED'.strip().lower() and "RUNNED".lower() != scheduletest_runvalue1[-1].strip().lower() and "RUNNED".lower() != continuoustest_runvalue1[-1].strip().lower()) and (Remotetest_runvalue1[-1].strip().lower() != 'FINISHED'.strip().lower() and "FINISHED".lower() != scheduletest_runvalue1[-1].strip().lower() and "FINISHED".lower() != continuoustest_runvalue1[-1].strip().lower()) and "LOADING".lower() != Change_Settingsrunvalue1[-1].strip().lower() and "RUNNED".lower() != Change_Settingsrunvalue1[-1].strip().lower() and "WAITING LOAD".lower() != Change_Settingsrunvalue1[-1].strip().lower() and "RUNNED".lower() != default_Settingsrunvalue1[-1].strip().lower() and "FINISHED".lower() != default_Settingsrunvalue1[-1].strip().lower():
        try:
            if os.path.exists(ReadConfig.test_data_folder_rootpath):
                with allure.step("Data file is present"):
                    with allure.step(f"Item count: {item_count}"):
                        f1 = open(ReadConfig.test_run_excelreportdata_path, "r")
                        testrunexcelfolder = f1.read()
                        test_run_excel_report_pathtimestamp = ReadConfig.excel_report_path + testrunexcelfolder
                        if os.path.exists(test_run_excel_report_pathtimestamp):
                            print("test run excel folder is exist")
                        if not os.path.exists(test_run_excel_report_pathtimestamp):
                            pytest.fail("test run excel folder is not exist")
                        high_level_excel_report_path = ReadConfig.excel_report_path + testrunexcelfolder
                        highlevelExcelReport_folder = high_level_excel_report_path + "\\highlevelExcelReport"
                        highlevelExcelReport_path = create_folder_for_excelreport(highlevelExcelReport_folder)
                        highlevelExcelReport(test_run_excel_report_pathtimestamp,highlevelExcelReport_path)
                        if compare_same_hld_of_diffrentserver_runvalue1[-1].strip().lower() == 'Yes'.strip().lower():
                            # Load the Excel file
                            excel_file = ReadConfig.test_data_path  # Update with the path to your Excel file
                            df = pd.read_excel(excel_file, sheet_name="ENVIRONMENTS_USERINPUT_LOGIN")
                            # Filter the "Environment" column where "Execute" column equals 'Yes'
                            filtered_env = df.loc[df['Execute'] == 'Yes', 'Environment'].tolist()  # Convert to list
                            if len(filtered_env) == 2:
                                excelhld1 = highlevelExcelReport_path + "\\HLD.xlsx"
                                excelhld2 = ""
                                if not os.path.exists(ReadConfig.test_data_folder_rootpath + "\\HLD_Comparision_Input_Files"):
                                    create_folder_for_excelreport(destination_folder=ReadConfig.test_data_folder_rootpath + "\\HLD_Comparision_Input_Files")
                                timestamp = datetime.datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
                                hld_comparsion_result_folder = ReadConfig.test_data_folder_rootpath+f"\\HLD_Comparision_Input_Files\\HLD_comparsion_result_{timestamp}"
                                create_folder_for_excelreport(destination_folder=hld_comparsion_result_folder)
                                comparision(files=[excelhld1],directory=config.test_data_folder_rootpath + "\\HLD_Comparision_Input_Files",timestamped_folder=hld_comparsion_result_folder)
                                # comparing_hld_result(excelhld1, excelhld2,comparing_2hld_file_value="No",excelpath_for_ouput=hld_comparsion_result_folder+f"\\HLD_comparsion_result_{timestamp}.xlsx",server1=filtered_env[0],server2=filtered_env[1])
                        f1.close()

        except Exception as e:
            with allure.step(f"Item count: {item_count}"):
                with allure.step(f"Goto the directory {ReadConfig.test_data_folder_rootpath} and update the file as per your requirement, then re-run"):
                    if not os.path.exists(ReadConfig.test_data_folder_rootpath):
                        create_folder_for_rantcell_data_and_ExcelReport(ReadConfig.test_data_folder_rootpath, ReadConfig.source_dest)
                        Updating_source_folder(ReadConfig.updating_source_folders, ReadConfig.test_data_folder_rootpath)
                        pytest.fail(f"Goto the directory {ReadConfig.test_data_folder_rootpath} and update the file as per your requirement")
