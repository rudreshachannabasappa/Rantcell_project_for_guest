import time
from utils.library import *
from locators.locators import *
from pageobjects.login_logout import *
from pageobjects.Dashboard import *
import pandas as pd
from utils.updateexcelfile import *
import re
################################################ For all 26 parameters and it's data in Settings section [Default] ###################################################################################################
def dashboard_default_setting(driver,combine_dict):
    Page_Down(driver)
    clickec(driver, settings_1.btn_setting)
    time.sleep(2)
    Page_Down(driver)
    try:
        WebDriverWait(driver, 90).until(EC.presence_of_element_located(settings_1.default_settings1))
    except:
        pass
    clickec(driver, settings_1.default_settings_btn)
    time.sleep(2)
    clickec(driver, settings_1.save_settings_btn)
    time.sleep(2)
    data_extraction_settings(driver, combine_dict)
############################################# Extraction of data in Map Legend and NQC Table ########################################################################################################################################3
def Map_legend_and_NQC_worklist(driver,campaigns_data,excelpath,datas1,datas2):
    tests_data1 = []
    for i in range(len(campaigns_data)):
        # previous_device = []
        device, campaign, usercampaignsname, testgroup = campaigns_data[i]
        # Fetch components based on the campaign/classifier "T001","T002" etc
        remote_test_point, map_start_point, graph_start_point, export_start_point, load_start_point, PDF_Export_index_start_point, END_index = fetch_input_points()
        tests_list = fetch_components(campaign, map_start_point, graph_start_point)
        skip_tests = ['Failed Call', 'Web test', 'Sent SMS', 'Received SMS', 'Failed SMS', 'nrArfcn', 'nrPCI', 'nrCID','ECNO','BCCH_ARFCN', 'PSC', 'PCI', 'Data Type', 'Network Type', 'Arfcn', 'lteCID']
        tests = [test for test in tests_list if test not in skip_tests]
        tests_data1.append(tests)
    tests_data = [test_data1 for test_data in tests_data1 for test_data1 in test_data]
    tests_data = list(set(tests_data))
    print("tests_data-------------", tests_data)
    map_legend_nqc_components(driver, tests_data, excelpath, datas1, datas2)

def Map_legend_and_NQC(driver,campaign,excelpath,datas1,datas2):
    time.sleep(1)
    remote_test_point, map_start_point, graph_start_point, export_start_point, load_start_point, PDF_Export_index_start_point, END_index = fetch_input_points()
    tests_list = fetch_components(campaign, map_start_point, graph_start_point)
    skip_tests = ['Failed Call', 'Web test', 'Sent SMS', 'Received SMS', 'Failed SMS', 'nrArfcn', 'nrPCI', 'nrCID','ECNO', 'BCCH_ARFCN', 'PSC', 'PCI', 'Data Type', 'Network Type', 'Arfcn', 'lteCID']
    tests = [test for test in tests_list if test not in skip_tests]
    map_legend_nqc_components(driver,tests,excelpath, datas1, datas2)

def map_legend_nqc_components(driver,tests,excelpath, datas1, datas2):
    Title = "MAP VIEW"
    e_flag = None
    # # Fetch components based on the campaign/classifier "T001","T002" etc
    ## remote_test_point, map_start_point, graph_start_point, export_start_point, load_start_point, PDF_Export_index_start_point, END_index = fetch_input_points()
    ## tests_list = fetch_components(campaign, map_start_point, graph_start_point)
    #
    ## skip_tests = ['Failed Call', 'Web test','Sent SMS', 'Received SMS', 'Failed SMS','nrArfcn', 'nrPCI', 'nrCID', 'ECNO', 'BCCH_ARFCN', 'PSC', 'PCI', 'Data Type', 'Network Type', 'Arfcn','lteCID']
    ## tests = [test for test in tests_list if test not in skip_tests]
    ## tests = ['Ping Test', 'Download Test', 'Upload Test', 'HTTP Download Test', 'HTTP Upload Test', 'iperf Download Test', 'iperf Upload Test', 'Call Test','Stream Test', 'RSSI/RSCP', 'RSRP', 'RSRQ', 'nrSSRSRP', 'nrSSRSRQ', 'LteSNR', 'nrSsSinr']
    try:
        Notestdatafound_elements = driver.find_elements(*select_Map_View_Components.No_test_data_element)
        i = 0
        while len(Notestdatafound_elements) != 0:
            Notestdatafound_elements = driver.find_elements(*select_Map_View_Components.No_test_data_element)
            i += 1
            if len(Notestdatafound_elements) == 0 or i == 15:
                break
        closeFullTableView_elements = driver.find_elements(close_button.closeFullTableView[0],close_button.closeFullTableView[1])
        e_flag = 0
        Page_up(driver)
        while len(closeFullTableView_elements) == 0:
            try:
                clickec(driver, select_Map_View_Components.Expand_Map_View)
                try:
                    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((close_button.closeFullTableView[0],close_button.closeFullTableView[1])))
                except:
                    pass
                e_flag += 1
                closeFullTableView_elements = driver.find_elements(close_button.closeFullTableView[0],close_button.closeFullTableView[1])
                if len(closeFullTableView_elements) != 0 or e_flag == 5:
                    break
            except:
                click(driver, select_Map_View_Components.Expand_Map_View)
                try:
                    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((close_button.closeFullTableView[0],close_button.closeFullTableView[1])))
                except:
                    pass
                e_flag += 1
                closeFullTableView_elements = driver.find_elements(close_button.closeFullTableView[0],close_button.closeFullTableView[1])
                if len(closeFullTableView_elements) != 0 or e_flag == 5:
                    break
        WebDriverWait(driver, 2.5).until(EC.visibility_of_element_located(select_Map_View_Components.export_selection_box))
        closeFullTableView_elements = driver.find_elements(close_button.closeFullTableView[0],close_button.closeFullTableView[1])
        if len(closeFullTableView_elements) != 0:
            try:
                WebDriverWait(driver, 2.5).until(EC.presence_of_element_located(select_Map_View_Components.map_element_presence))
                element = driver.find_element(*select_Map_View_Components.map_element_verify)
                if element:
                    pass
                else:
                    with allure.step("Map is not loaded"):
                        allure.attach(driver.get_screenshot_as_png(), name="Map is not loaded",attachment_type=allure.attachment_type.PNG)
                WebDriverWait(driver, 1).until(EC.presence_of_element_located(select_Map_View_Components.satellite_element_presence))
                element = driver.find_element(*select_Map_View_Components.satellite_element_verify)
                if element:
                    pass
                else:
                    with allure.step("Map is not loaded"):
                        allure.attach(driver.get_screenshot_as_png(), name="Map is not loaded",attachment_type=allure.attachment_type.PNG)
            except:
                pass
            try:
                pattern_mapping_df = pd.read_excel(config.settings_path,sheet_name="MAP_SETTINGS")
            except Exception as e:
                with allure.step(f"Check {config.settings_path}"):
                    print(f"Check {config.settings_path}")
                    assert False
            pattern_mapping = pattern_mapping_df.set_index('TC Sheet Components').apply(lambda x: x.dropna().tolist(),axis=1).to_dict()
            txt = []
            if tests.__len__() == 0:
                statement = f"Map-View  --  Nothing is marked 'Yes' in {str(config.test_data_path)}"
                with allure.step(f"Nothing is marked 'Yes' in {str(config.test_data_path)} for 'Map-View"):
                    updatename(excelpath, statement)
                    updatecomponentstatus(Title, "", "FAILED", f"Nothing marked in {str(config.test_data_path)}",excelpath)
                    e = Exception
                    raise e
            else:
                for test in tests:
                    test = test.strip()
                    for pattern, values in pattern_mapping.items():
                        if pattern.lower() == test.lower():
                            txt = values[1:3]
                            test = values[0]
                            break
                        else:
                            txt = []
                    time.sleep(0.1)
                    try:
                        try:
                            listbox = WebDriverWait(driver, 0.1).until(EC.visibility_of_element_located(select_Map_View_Components.map_menu_dropdown))
                            if listbox.is_displayed():
                                listbox_btn = WebDriverWait(driver, 1.2).until(EC.visibility_of_element_located(select_Map_View_Components.Test_Type_Dropdown))
                                listbox_btn.click()
                        except:
                            pass
                        Map_view_Search_Box_not_visible_do_page_up_(driver)
                        data = read_map_legend(driver, select_Map_View_Components.Test_Type_Dropdown,select_Map_View_Components.nested_locators1,select_Map_View_Components.Call_Test_locator, txt, test, Title,excelpath, test)
                        if data['operator_comparsion_table_data']:
                            if test not in datas1:
                                datas1[test] = []
                            datas1[test].extend(data['operator_comparsion_table_data'])

                        if data['map_legend_data']:
                            if test not in datas2:
                                datas2[test] = []
                            datas2[test].extend(data['map_legend_data'])
                    except Exception as e:
                        continue
            click_closeButton(driver)
        elif len(closeFullTableView_elements) == 0:
            statement = f"Failed to click on the expand for {Title}"
            with allure.step(statement):
                allure.attach(driver.get_screenshot_as_png(), name=f"Expand_Map_View_screenshot",attachment_type=allure.attachment_type.PNG)
                e = Exception
                raise e
    except Exception as e:
        Notestdatafound_elements = driver.find_elements(By.XPATH,"// h3[contains(text(), 'No test data found. Please try different date and ')]")
        closeFullTableView_elements = driver.find_elements(close_button.closeFullTableView[0],close_button.closeFullTableView[1])
        if len(closeFullTableView_elements) == 0:
            statement = f"Failed to click on expand button for {Title}"
            Failupdatename(excelpath, statement)
            updatecomponentstatus(Title, "Expand_Map_View", "FAILED", statement, excelpath)
        elif e_flag == 1:
            print('select Map View Components fail')
        elif len(Notestdatafound_elements) != 0:
            statement = f"No test data found. 'Please try different date in Map View' statement is present due to that map didn't load"
            Failupdatename(excelpath, statement)
            updatecomponentstatus(Title, "No test data found. Please try different date", "FAILED", statement,excelpath)
def read_map_legend(driver, listbox_locator, nested_locators1, Call_Test_locator, option_text_list, elementname,Title, excelpath, test):
    ListboxSelectstatus = "None"
    map_legend_data = []
    operator_comparsion_table_data = []
    with allure.step(f"Map View Select '{elementname}' and Read Data"):
        try:
            time.sleep(1)
            try:
                wait_for_loading_elements(driver)
            except:
                pass
            l_flag = 0
            if option_text_list.__len__() == 0:
                with allure.step(f"In input data from {str(config.map_view_components_excelpath)} for 'Map-View for '{test}' in header of Map view Components column value against the 2nd row of headers of Map view in {str(config.test_data_path)} is mismatch/empty"):
                    l_flag = 2
                    allure.attach(driver.get_screenshot_as_png(), name=f"{elementname}_screenshot",attachment_type=allure.attachment_type.PNG)
                    e = Exception
                    raise e
            elif ["Call Test", "Call Test"] != option_text_list:
                ListboxSelectstatus, alert_text = select_from_listbox_ECs(driver, listbox_locator, nested_locators1,option_text_list, Title, excelpath)
                l_flag = 1
            elif ["Call Test", "Call Test"] == option_text_list:
                clickEC_for_listbox(driver, Map_View_Select_and_ReadData.Test_Type_Dropdown_for_call_Test, Title,excelpath)
                clickEC_for_listbox(driver, Map_View_Select_and_ReadData.Call_Test_locator2, Title, excelpath)
                ListboxSelectstatus, alert_text = clickEC_for_listbox(driver, Call_Test_locator, Title, excelpath)
                l_flag = 1
            time.sleep(1)
            try:
                wait_for_loading_elements(driver)
            except:
                pass
            if alert_text == None and l_flag == 1:
                try:
                    time.sleep(1)
                    operator_data = driver.find_elements(*settings_1.operator_comparison_data)
                    for i in range(len(operator_data)):
                        j = operator_data[i].text
                        list_a = ['>=', 'B/w', 'below', '-', '>', 'Bw', 'above', '<=']
                        list_b = ["to"]
                        list_c = ["Call"]
                        if any(a.lower() in j.lower() for a in list_a) and not any(c.lower() in j.lower() for c in list_c):
                            operator_comparsion_table_data.append(j)
                        elif any(b.lower() in j.lower() for b in list_b) and test == "Call Test" and i==0:
                            operator_comparsion_table_data.append(j)
                except:
                    pass
                try:
                    if len(operator_comparsion_table_data) != 0:
                        with allure.step("Operator Comparison Data Extraction"):
                            # updatecomponentstatus("MAP VIEW", str(test), "PASSED",f"Passed step :- In Operator comparison table for {option_text_list[-1]} data is found in table",excelpath)
                            allure.attach(driver.get_screenshot_as_png(), name="Operator Comparison Data",attachment_type=allure.attachment_type.PNG)
                    elif (len(operator_comparsion_table_data) == 0 or operator_comparsion_table_data == None) and test != "iperf Download Test" and test != "iperf Upload Test":
                        e = Exception
                        raise e
                except Exception:
                    if len(operator_comparsion_table_data) == 0 or operator_comparsion_table_data == None:
                        with allure.step(f"Failed step :- In Operator comparsion table for {option_text_list[-1]} No data in table"):
                            Failupdatename(excelpath,f"Failed step :- In Operator comparison table for {option_text_list[-1]} No data in table")
                            raise Exception
                try:
                    time.sleep(1)
                    map_legend = driver.find_elements(*settings_1.map_legend_each_elements)
                    for m in range(len(map_legend)):
                        l = map_legend[m].text
                        if test != "Call Test" and test != "Stream Test" and l.lower().replace(" ","") != "dropped packets".lower().replace(" ",""):
                            map_legend_data.append(l)
                except:
                    pass
                try:
                    if len(map_legend_data) != 0:
                        with allure.step("Map Legend Data Extraction"):
                            # updatecomponentstatus("MAP VIEW", str(test), "PASSED",f"Passed step :- In Map Legend for {option_text_list[-1]} data is found",excelpath)
                            allure.attach(driver.get_screenshot_as_png(), name="Map Legend Data",attachment_type=allure.attachment_type.PNG)
                    elif len(map_legend_data) == 0 or map_legend_data == None and test != "Call Test" and test != "Stream Test":
                        e = Exception
                        raise e
                except Exception:
                    if len(map_legend_data) == 0 or map_legend_data == None:
                        with allure.step(f"Failed step :- In Map Legend for {option_text_list[-1]}there is no data"):
                            Failupdatename(excelpath,f"Failed step :- In Map Legend for {option_text_list[-1]} there is no data")
                            raise Exception
            elif ListboxSelectstatus == 0 and alert_text != None and l_flag == 1:
                e = Exception
                with allure.step(f"failed step :- Alert Found is '{alert_text}' for Map View to select {elementname}"):
                    updatecomponentstatus(Title, elementname, "FAILED",f"Alert Found is '{alert_text}' for Map View to select {elementname}",excelpath)
                    raise e
        except Exception as e:
            print("Map View Select and Read Data fail")
            if l_flag == 0:
                statement = f"Unable to locate the element/No such element found and so error in selecting " + str(option_text_list) + " from listbox"
                with allure.step(statement):
                    allure.attach(driver.get_screenshot_as_png(), name=f"{elementname}_screenshot",attachment_type=allure.attachment_type.PNG)
                    Failupdatename(excelpath, statement)
                    raise e
            elif option_text_list.__len__() == 0 and l_flag == 2:
                statement = f"In input data from {str(config.map_view_components_excelpath)} for 'Map-View for '{test}' in header of Map view Components column value against the 2nd row of headers of Map view in {str(config.test_data_path)} is mismatch/empty"
                with allure.step(statement):
                    allure.attach(driver.get_screenshot_as_png(), name=f"{elementname}_screenshot",attachment_type=allure.attachment_type.PNG)
                    Failupdatename(excelpath, statement)
                    updatecomponentstatus("MAP VIEW", str(test), "FAILED", statement, excelpath)
                    raise e
            elif alert_text != None and l_flag == 1:
                statement = f"Alert Found is '{alert_text}' for Map View to select {elementname}"
                with allure.step(statement):
                    allure.attach(driver.get_screenshot_as_png(), name=f"{elementname}_screenshot",attachment_type=allure.attachment_type.PNG)
                    Failupdatename(excelpath, statement)
                    raise e
    return {'map_legend_data': map_legend_data, 'operator_comparsion_table_data': operator_comparsion_table_data}
def settings_pdf(driver, campaign, data_combine_dict, excelpath):
    remote_test_point, map_start_point, graph_start_point, export_start_point, load_start_point, PDF_Export_index_start_point, END_index = fetch_input_points()
    tests_list = fetch_components(campaign, PDF_Export_index_start_point, END_index)
    # List of parameters
    skip_tests = ['Failed Call', 'Web test', 'Sent SMS', 'Received SMS', 'Failed SMS', 'nrArfcn', 'nrPCI', 'nrCID','ECNO', 'BCCH_ARFCN', 'PSC', 'PCI', 'Data Type', 'Network Type', 'Arfcn', 'lteCID']
    tests = [test for test in tests_list if test not in skip_tests]
    settings_pdf_data(driver, tests, data_combine_dict, excelpath)
def settings_pdf_data_worklist(driver,campaigns_data,data_combine_dict,excelpath):
    tests_data1 = []
    for i in range(len(campaigns_data)):
        # previous_device = []
        device, campaign, usercampaignsname, testgroup = campaigns_data[i]
        # Fetch components based on the campaign/classifier "T001","T002" etc
        remote_test_point, map_start_point, graph_start_point, export_start_point, load_start_point, PDF_Export_index_start_point, END_index = fetch_input_points()
        tests_list = fetch_components(campaign, PDF_Export_index_start_point, END_index)
        skip_tests = ['Failed Call', 'Web test', 'Sent SMS', 'Received SMS', 'Failed SMS', 'nrArfcn', 'nrPCI', 'nrCID',
                      'ECNO', 'BCCH_ARFCN', 'PSC', 'PCI', 'Data Type', 'Network Type', 'Arfcn', 'lteCID']
        tests = [test for test in tests_list if test not in skip_tests]
        tests_data1.append(tests)
    tests_data = [test_data1 for test_data in tests_data1 for test_data1 in test_data]
    tests_data = list(set(tests_data))
    print("tests_data-------------", tests_data)
    settings_pdf_data(driver,tests_data,data_combine_dict,excelpath)

def settings_pdf_data(driver,tests,data_combine_dict,excelpath):
    Title = "PDF Data"
    with allure.step("PDF Data Extraction"):
        time.sleep(1)
        List_of_options_txt = ["Export As PDF"]
        allure.attach(driver.get_screenshot_as_png(), name="PDF Data Extraction",attachment_type=allure.attachment_type.PNG)
        ## remote_test_point, map_start_point, graph_start_point, export_start_point, load_start_point, PDF_Export_index_start_point, END_index = fetch_input_points()
        ## tests_list = fetch_components(campaign, PDF_Export_index_start_point, END_index)
        try:
            # if WebDriverWait(driver, 1).until(EC.visibility_of_element_located(List_Of_Campaigns_Export_Dashboard.List_Of_Campaigns_Export_Dropdown)).is_displayed():
            #     listbox_btn = WebDriverWait(driver, 1).until(EC.visibility_of_element_located(List_Of_Campaigns_Export_Dashboard.List_Of_Campaigns_Export_Dropdown))
            #     # Click on the listbox to close it
            #     listbox_btn.click()
            select_from_listbox_ECs(driver, List_Of_Campaigns_Export_Dashboard.List_Of_Campaigns_Export_Dropdown,List_Of_Campaigns_Export_Dashboard.List_Of_Campaigns_Export_Dropdown_Options,List_of_options_txt, Title, excelpath)
        except:
            pass
        time.sleep(5)
        # clickec(driver, settings_1.pdf_btn_setting)
        # time.sleep(7)
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(7)
        # List of parameters
        ## skip_tests = ['Failed Call', 'Web test', 'Sent SMS', 'Received SMS', 'Failed SMS','nrArfcn', 'nrPCI', 'nrCID', 'ECNO', 'BCCH_ARFCN', 'PSC', 'PCI', 'Data Type', 'Network Type','Arfcn','lteCID']
        ## tests = [test for test in tests_list if test not in skip_tests]
        try:
            pattern_mapping_df = pd.read_excel(config.settings_path,sheet_name="PDF_SETTINGS")
        except Exception as e:
            with allure.step(f"Check {config.settings_path}"):
                print(f"Check {config.settings_path}")
                assert False
        pattern_mapping = pattern_mapping_df.set_index('pdf_export').apply(lambda x: x.dropna().tolist(),axis=1).to_dict()
        txt = []
        if tests.__len__() == 0:
            statement = f"Map-View  --  Nothing is marked 'Yes' in {str(config.test_data_path)}"
            with allure.step(f"Nothing is marked 'Yes' in {str(config.test_data_path)} for 'Map-View"):
                updatename(excelpath, statement)
                updatecomponentstatus(Title, "", "FAILED", f"Nothing marked in {str(config.test_data_path)}",excelpath)
                e = Exception
                raise e
        else:
            for test in tests:
                test = test.strip()
                for pattern, values in pattern_mapping.items():
                    if pattern.lower() == test.lower():
                        txt = values
                        break
                    else:
                        txt = []
                if len(txt) !=0:
                    # parameters = ['Ping Test', 'Download Test', 'Upload Test', 'Http Download Test','Http Upload Test','iPerf Test','Call Test','Stream Test','RSSI/RSCP','RSRP','RSRQ','nrSsRsrp','nrSsRsrq','lteSNR','nrSsSinr']  # Add your parameters here
                    parameter = txt[0]
                    try:
                        parameter1 = parameter.replace(" ","")
                        # Find and click the checkbox for the current parameter
                        checkbox = driver.find_element(By.XPATH, f"//div[@id='checkboxes']//*[contains(., '{parameter1}')]")
                        if checkbox.is_enabled():
                            checkbox.click()
                            time.sleep(1)
                        else:
                            print(f"Checkbox corresponding to parameter '{parameter}' is not enabled. Skipping.")
                            continue  # Skip collecting PDF data for this parameter if checkbox is not enabled
                    except (NoSuchElementException, ElementNotInteractableException):
                        print(f"Checkbox corresponding to parameter '{parameter}' not found or not interactable.")
                        continue  # Skip collecting PDF data for this parameter if checkbox is not found or not interactable
                    locator = None
                    pdf_data = []  # Reset pdf_data for each parameter
                    try:
                        testtype = str(txt[1]).replace(' ','')
                        locator = (By.XPATH,f"//div[@id ='{testtype}']//td[1]//span")

                        if locator:
                            time.sleep(2)
                            elements = driver.find_elements(*locator)
                            # for element in elements:
                            for i in range(len(elements)):
                                j = elements[i].text
                                list_a = ['>=', 'B/w', 'below', '-', '>', 'Bw', 'above', '<=']
                                list_b = ["to"]
                                list_c = ["Call"]
                                if any(a.lower() in j.lower() for a in list_a) and not any(c.lower() in j.lower() for c in list_c):
                                    pdf_data.append(j)
                                elif any(a.lower() in j.lower() for a in list_b) and parameter == "Call Test" and i == 0:
                                    pdf_data.append(j)
                    except Exception as e:
                        pass
                    if pdf_data:  # Check if pdf_data is not empty
                        data_combine_dict[parameter] = pdf_data
                allure.attach(driver.get_screenshot_as_png(), name="PDF Data",attachment_type=allure.attachment_type.PNG)
        try:
            time.sleep(2)
            driver.switch_to.window(driver.window_handles[1])
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            # Print or return the combined values
            print(data_combine_dict)
        except:pass
def main_func_default_settings(driver,environment,userid):
    runvalue = Testrun_mode(value="Default Settings")
    if "Yes".lower() == runvalue[-1].strip().lower():
        combine_dict = {}
        dashboard_default_setting(driver,combine_dict)
        combine_dict1 = {}
        data = {
            'Test_Type': [],
            'Parameter': [],
        }
        for key, values in combine_dict.items():
            combine_dict1[key] = [{setting: extract_numerical_values(setting)} for setting in values]
        for test_type, values in combine_dict1.items():
            for value in values:
                for param, param_values in value.items():
                    data['Test_Type'].append(test_type)
                    data['Parameter'].append(param)
        # Convert to DataFrame
        df = pd.DataFrame(data)
        default_excel_path = config.test_data_folder_rootpath +f"\\testdata\\{environment}_{userid}_default_setting.xlsx"
        # Write DataFrame to Excel
        df.to_excel(default_excel_path, index=False)
def main_default_settings(driver,campaign,excelpath,environment,userid):
    runvalue = Testrun_mode(value="Default Settings")
    if "RUNNED".lower() == runvalue[-1].strip().lower():
        Title = "DEFAULT SETTINGS"
        datas1 = {}
        datas2 = {}
        data_combine_dict = {}
        combine_dict1 = {}
        final_result = {"Operator_Comparison": [], "Map_Legend": [], "Pdf_Data": []}
        default_excel_path = config.test_data_folder_rootpath +f"\\testdata\\{environment}_{userid}_default_setting.xlsx"
        df = pd.read_excel(default_excel_path)
        combine_dict = df.groupby("Test_Type")["Parameter"].apply(list).to_dict()
        Map_legend_and_NQC(driver,campaign,excelpath,datas1,datas2)
        settings_pdf_data(driver,campaign,data_combine_dict,excelpath)
        # Updating settings, Operator Comparison, Map Legend and PDF Data to excel
        updating_settings_data_extraction_to_excel(combine_dict, datas1, datas2, data_combine_dict, excelpath,"DATA_EXTRACTION_SETTINGS")

        # Extracting only numerical values from the "Settings" section
        for key, values in combine_dict.items():
            combine_dict1[key] = [extract_numerical_values(setting) for setting in values]
        # Calling comparison function to compare the settings data against operator comparison, map_legend and pdf data
        comparison(datas1, datas2, data_combine_dict, combine_dict1,final_result,Title,excelpath)
        updating_comparison_results_to_excel1(final_result, excelpath,"RESULTS_DEFAULT_SETTINGS")
def main_default_settings_worklist(driver,campaigns_data,excelpath,environment,userid):
    runvalue = Testrun_mode(value="Default Settings")
    if "RUNNED".lower() == runvalue[-1].strip().lower():
        Title = "DEFAULT SETTINGS"
        datas1 = {}
        datas2 = {}
        data_combine_dict = {}
        combine_dict1 = {}
        final_result = {"Operator_Comparison": [], "Map_Legend": [], "Pdf_Data": []}
        default_excel_path = config.test_data_folder_rootpath + f"\\testdata\\{environment}_{userid}_default_setting.xlsx"
        df = pd.read_excel(default_excel_path)
        combine_dict = df.groupby("Test_Type")["Parameter"].apply(list).to_dict()
        Map_legend_and_NQC_worklist(driver,campaigns_data,excelpath,datas1,datas2)
        settings_pdf_data_worklist(driver,campaigns_data,data_combine_dict,excelpath)
        # Updating settings, Operator Comparison, Map Legend and PDF Data to excel
        updating_settings_data_extraction_to_excel(combine_dict, datas1, datas2, data_combine_dict, excelpath,"DATA_EXTRACTION_SETTINGS")
        # Extracting only numerical values from the "Settings" section
        for key, values in combine_dict.items():
            combine_dict1[key] = [extract_numerical_values(setting) for setting in values]
        # Calling comparison function to compare the settings data against operator comparison, map_legend and pdf data
        comparison(datas1, datas2, data_combine_dict, combine_dict1, final_result, Title, excelpath)
        updating_comparison_results_to_excel1(final_result, excelpath, "RESULTS_DEFAULT_SETTINGS")

# function for comparison of default settings
def comparison(datas1, datas2, data_combine_dict, combine_dict,final_result,Title,excelpath):
    # Validation point to check whether the data from excel is correctly updated in the Application
    excel_data = pd.read_excel("C:\\RantCell_Automation_Data_and_Reports\\testdata\\Change_settings.xlsx",dtype={'Value': str})
    dict_exceldata = excel_data.groupby("Test_Type")["Value"].apply(list).to_dict()
    testtype =[]
    if Title == "CHANGE SETTINGS":
        for settingparameter, values in dict_exceldata.items():
            if settingparameter in combine_dict:
                flag_failed = True
                observed_values = combine_dict[settingparameter]
                extracted_settings = [numeric_value for observed_numeric_value in observed_values for numeric_value in observed_numeric_value]
                comparison_results = []
                i = 0
                for expected_value in values:
                    i += 1
                    matched = [numeric_value for numeric_value in extracted_settings if compare_values(str(numeric_value).replace("-", ""),str(expected_value).replace("-", ""))]
                    if matched:
                        comparison_results.append({"SETTINGS_PARAMETER_NAME(Reference)": settingparameter,f"EXCEL VALUE": f"{expected_value}",f"SETTINGS_APPLICATION_VALUE(Reference)": f"{matched}",f"Data validation": "The value is found"})
                    else:
                        flag_failed = False
                        comparison_results.append({"SETTINGS_PARAMETER_NAME(Reference)": settingparameter,f"EXCEL VALUE": f"{expected_value}",f"SETTINGS_APPLICATION_VALUE(Reference)": f"{extracted_settings[i]}",f"Data validation": "The value is Not Found"})
                final_result["SETTINGS"].extend(comparison_results)
                if flag_failed == True:
                    updatecomponentstatus(Title=Title,componentname=f"{settingparameter} --> Excel Data vs Change Settings Data(Application)",status="PASSED",comments=f"The values are found when comparing Excel Data vs Change Settings Data(Application)",path=excelpath)
                elif flag_failed == False:
                    testtype.append(settingparameter)
                    updatecomponentstatus(Title=Title,componentname=f"{settingparameter} --> Excel Data vs Change Settings Data(Application)",status="FAILED",comments=f"The values are Not found when comparing Excel Data vs Change Settings Data(Application)",path=excelpath)
    # Comparison function starts from here
    dict_list = {
        "Operator_Comparison": datas1,
        "Map_Legend": datas2,
        "Pdf_Data": data_combine_dict
    }
    for reference_key in combine_dict.keys():
        i = 0
        for view_key, other_dict in dict_list.items():
            for other_key, other_values in other_dict.items():
                matched_key = key_match(reference_key, other_key)
                if matched_key:
                    flag_run = []
                    comparison_results = compare_values_setting(combine_dict[reference_key], other_values, view_key, other_key, reference_key,Title,flag_run,testtype,excelpath)
                    if len(flag_run) != 0:
                        final_result[view_key].extend(comparison_results)
                    if "Download" not in reference_key and "Upload" not in reference_key:
                        break
                else:
                    if i == 3:
                        print(f"{reference_key} not found any of the views")
            i += 1
def key_match(reference_key, other_key):
    # Convert keys to lowercase for case-insensitive comparison
    reference_key_lower = reference_key.lower()
    other_key_lower = other_key.lower()
    # Split the other key into parts
    other_parts = other_key_lower.split()
    if len(other_parts) >= 3:
        if all(part in reference_key_lower for part in other_parts[:2]):
            return other_key
    elif len(other_parts) == 2:
        if all(part in reference_key_lower for part in other_parts[:1]):
            return other_key
    elif len(other_parts) == 1:
        reference_key = reference_key_lower.split()
        if all(compare_values(part,reference_key[0])for part in other_parts):
            return other_key
    return None
def contains_only_empty_strings(lst):
    return all(item == '' for item in lst)
def compare_values_setting(value_list, other_dict_list, view_key, other_key, reference_key,Title,flag_run,testtype,excelpath):
    comparison_results = []
    i =0
    flag_failed = True
    if not contains_only_empty_strings(other_dict_list):
        flag_run.append(True)
        for val in value_list:
            a = [value1 for value1 in other_dict_list if all(check_numeric_value(v, value1) for v in val)]
            b = [value1 for value1 in other_dict_list if not any(check_numeric_value(v, value1) for v in val)]
            if a and not reference_key in testtype :
                 comparison_results.append({"SETTINGS_PARAMETER_NAME(Reference)": reference_key, f"{view_key} PARAMETER": other_key,f"{view_key} VALUE": f"{a}", f"SETTINGS_APPLICATION_VALUE(Reference)": f"{val}",f"Data validation": "The value is found"})
            elif a and reference_key in testtype :
                 comparison_results.append({"SETTINGS_PARAMETER_NAME(Reference)": reference_key, f"{view_key} PARAMETER": other_key,f"{view_key} VALUE": f"{a}", f"SETTINGS_APPLICATION_VALUE(Reference)": f"{val}",f"Data validation": "The value is found,but settings application value(reference) != excel settings values"})
                 flag_failed = [True,False]
            elif b:
                flag_failed = False
                comparison_results.append({"SETTINGS_PARAMETER_NAME(Reference)": reference_key, f"{view_key} PARAMETER": other_key,f"{view_key} VALUE": f"{b[i]}", f"SETTINGS_APPLICATION_VALUE(Reference)": f"{val}",f"Data validation": "The value is Not Found"})
            i+=1
            print(i)
        if flag_failed == True and i != 0:
            updatecomponentstatus(Title=Title,componentname=f"{reference_key} == {other_key} --> {view_key} vs Settings", status="PASSED", comments=f"The values are found.", path=excelpath)
        elif flag_failed == False and i != 0:
            updatecomponentstatus(Title=Title,componentname=f"{reference_key} == {other_key} --> {view_key} vs Settings", status="FAILED", comments=f"The values are not found.", path=excelpath)
        elif flag_failed == [True,False] and i != 0:
            updatecomponentstatus(Title=Title,componentname=f"{reference_key} == {other_key} --> {view_key} vs Settings", status="FAILED", comments=f"The values are found,but settings application value(reference) != excel settings values", path=excelpath)
    return comparison_results
########################################################### Change Settings Scenario ####################################################################################################################################################################
def main_func_change_settings(driver,environment,userid):
    runvalue = Testrun_mode(value="Change Settings")
    if "LOADING".lower() == runvalue[-1].strip().lower():
        combine_dict = {}
        change_settings(driver, combine_dict)
        combine_dict1 = {}
        data = {
            'Test_Type': [],
            'Parameter': [],
        }
        for key, values in combine_dict.items():
            combine_dict1[key] = [{setting: extract_numerical_values(setting)} for setting in values]
        for test_type, values in combine_dict1.items():
            for value in values:
                for param, param_values in value.items():
                    data['Test_Type'].append(test_type)
                    data['Parameter'].append(param)
        # Convert to DataFrame
        df = pd.DataFrame(data)
        change_excel_path = config.test_data_folder_rootpath + f"\\testdata\\{environment}_{userid}_change_setting.xlsx"
        # Write DataFrame to Excel
        df.to_excel(change_excel_path, index=False)
def main_change_settings(driver,campaign,environment,userid,excelpath):
    runvalue = Testrun_mode(value="Change Settings")
    if "RUNNED".lower() == runvalue[-1].strip().lower():
        Title = "CHANGE SETTINGS"
        final_result = {"Operator_Comparison": [], "Map_Legend": [], "Pdf_Data": [], "SETTINGS":[]}
        datas1 = {}
        datas2 = {}
        data_combine_dict = {}
        combine_dict1 = {}
        change_excel_path = config.test_data_folder_rootpath + f"\\testdata\\{environment}_{userid}_change_setting.xlsx"
        df = pd.read_excel(change_excel_path)
        combine_dict = df.groupby("Test_Type")["Parameter"].apply(list).to_dict()
        Map_legend_and_NQC(driver,campaign,excelpath,datas1,datas2)
        settings_pdf_data(driver,campaign,data_combine_dict,excelpath)
        # Updating settings, Operator Comparison, Map Legend and PDF Data to excel
        updating_settings_data_extraction_to_excel(combine_dict, datas1, datas2, data_combine_dict, excelpath,"DATA_EXTRACTION_CHANGE_SETTINGS")

        # Extracting only numerical values from the "Settings" section
        for key, values in combine_dict.items():
            combine_dict1[key] = [extract_numerical_values(setting) for setting in values]
        # Calling comparison function to compare the settings against operator comparison, map_legend and pdf data
        comparison(datas1, datas2, data_combine_dict, combine_dict1,final_result,Title,excelpath)
        updating_comparison_results_to_excel1(final_result, excelpath, "RESULTS_CHANGE_SETTINGS")
def change_settings_worklist(driver,campaigns_data,environment,userid,excelpath):
    runvalue = Testrun_mode(value="Change Settings")
    if "RUNNED".lower() == runvalue[-1].strip().lower():
        Title = "CHANGE SETTINGS"
        final_result = {"Operator_Comparison": [], "Map_Legend": [], "Pdf_Data": [], "SETTINGS": []}
        datas1 = {}
        datas2 = {}
        data_combine_dict = {}
        combine_dict1 = {}
        change_excel_path = config.test_data_folder_rootpath + f"\\testdata\\{environment}_{userid}_change_setting.xlsx"
        df = pd.read_excel(change_excel_path)
        combine_dict = df.groupby("Test_Type")["Parameter"].apply(list).to_dict()
        Map_legend_and_NQC_worklist(driver, campaigns_data, excelpath, datas1, datas2)
        settings_pdf_data_worklist(driver, campaigns_data, data_combine_dict, excelpath)
        # Updating settings, Operator Comparison, Map Legend and PDF Data to excel
        updating_settings_data_extraction_to_excel(combine_dict, datas1, datas2, data_combine_dict, excelpath,"DATA_EXTRACTION_CHANGE_SETTINGS")
        # Extracting only numerical values from the "Settings" section
        for key, values in combine_dict.items():
            combine_dict1[key] = [extract_numerical_values(setting) for setting in values]
        # Calling comparison function to compare the settings against operator comparison, map_legend and pdf data
        comparison(datas1, datas2, data_combine_dict, combine_dict1, final_result, Title, excelpath)
        updating_comparison_results_to_excel1(final_result, excelpath, "RESULTS_CHANGE_SETTINGS")

def change_settings(driver, combine_dict):
    with allure.step("Change Settings scenario"):
        Page_Down(driver)
        clickec(driver, settings_1.btn_setting)
        time.sleep(2)
        excel_data = pd.read_excel("C:\\RantCell_Automation_Data_and_Reports\\testdata\\Change_settings.xlsx", dtype={'Value': str})
        # Define a dictionary to map test types to locators
        test_type_locators = {
            "RSSI/RSCP dBm setting": {
                "Greater than equal to": settings_1.rssi_rscp_1,"Range1": settings_1.rssi_rscp_2,"Range2": settings_1.rssi_rscp_3
            },
            "WIFI RSSI dBm setting": {
                "Greater than equal to": settings_1.wifi_rssi_1,"Range1": settings_1.wifi_rssi_2,"Range2": settings_1.wifi_rssi_3
            },
            "RSRP dBm setting": {
                "Greater than equal to": settings_1.rsrp_1,"Range1": settings_1.rsrp_2,"Range2": settings_1.rsrp_3
            },
            "RSRQ dBm setting": {
                "Greater than equal to": settings_1.rsrq_1,"Range1": settings_1.rsrq_2
            },
            "lteSNR dBm setting": {
                "Greater than equal to": settings_1.ltesnr_1,"Range1": settings_1.ltesnr_2,"Range2": settings_1.ltesnr_3,
                "Range3": settings_1.ltesnr_4,"Range4": settings_1.ltesnr_5,"Range5": settings_1.ltesnr_6,"Range6": settings_1.ltesnr_7
            },
            "CDMA RSSI dBm setting": {
                "Greater than equal to": settings_1.cdma_rssi_1,"Range1": settings_1.cdma_rssi_2,"Range2": settings_1.cdma_rssi_3
            },
            "3G Ec/No dBm setting": {
                "Greater than equal to": settings_1.ecno_1,"Range1": settings_1.ecno_2
            },
            "CDMA SNR dBm setting": {
                "Less than equal to": settings_1.cdma_snr_1,"Range1": settings_1.cdma_snr_2
            },
            "nrSsSINR dBm setting": {
                "Greater than equal to": settings_1.nrSsSINR_1,"Range1": settings_1.nrSsSINR_2,"Range2": settings_1.nrSsSINR_3,
                "Range3": settings_1.nrSsSINR_4,"Range4": settings_1.nrSsSINR_5, "Range5": settings_1.nrSsSINR_6,"Range6": settings_1.nrSsSINR_7,
            },
            "nrSsRSRP dBm setting": {
                "Greater than equal to": settings_1.nrSsRSRP_1,"Range1": settings_1.nrSsRSRP_2,"Range2": settings_1.nrSsRSRP_3
            },
            "nrSsRSRQ dBm setting": {
                "Greater than equal to": settings_1.nrSsRSRQ_1,"Range1": settings_1.nrSsRSRQ_2
            },
            "Ping Test": {
                "Less than equal to   ms": settings_1.ping_1
            },
            "Call Setup Time": {
                "Range1": settings_1.call_setup_time_1,"Greater than   sec and less than equal to   sec": settings_1.call_setup_time_2
            },
            "SMS Sent/Received Duration (Graph view)": {
                "Range1": settings_1.sms_sent_received_1,"Greater than   sec and  Less than equal to  sec": settings_1.sms_sent_received_2
            },
            "Download Test / HTTP Speed Download Test / iPerf DownloadTest": {
                "Greater than equal to   mbps": settings_1.download_http_iperf_1,"Range1": settings_1.download_http_iperf_2
            },
            "Upload Test / HTTP Speed Upload Test / iPerf UploadTest": {
                "Greater than equal to   mbps": settings_1.upload_http_iperf_1,"Range1": settings_1.upload_http_iperf_2
            },
            "Stream Test(Graph view)": {
                "Greater than equal to   mbps": settings_1.stream_1,"Range1": settings_1.stream_2
            }
        }
        # Iterate over each row in the Excel data
        for index, row in excel_data.iterrows():
            test_type = row['Test_Type'].strip()
            parameter = row['Parameter'].strip()
            value = row['Value']
            if test_type in test_type_locators:
                locator_map = test_type_locators[test_type]
                if parameter in locator_map:
                    locator = locator_map[parameter]
                    # data_df[test_type] = [].append([value])
                    inputtext(driver=driver, locators=locator, value=value)
                    with allure.step(test_type):
                        allure.attach(driver.get_screenshot_as_png(), name=f"{test_type}",attachment_type=allure.attachment_type.PNG)
                else:
                    print(f"Parameter '{parameter}' not recognized for test type '{test_type}'.")
            else:
                print(f"Test type '{test_type}' not recognized.")
        # print(data_df)
        time.sleep(1)
        clickec(driver, settings_1.save_settings_btn)
        time.sleep(2)
        data_extraction_settings(driver, combine_dict)
def data_extraction_settings(driver,combine_dict):
    headers_list = []
    header_element = driver.find_elements(*settings_1.all_default_settings_headers)
    for k in range(len(header_element)):
        try:
            d = header_element[k].text
            headers_list.append(d)
            data_default_content = header_element[k].find_elements(*settings_1.all_default_settings_content)
            data_default_values = header_element[k].find_elements(*settings_1.all_default_settings_values)
            value_index = 0  # Initialize index for values
            combined_values = []
            # Loop through each data element
            for i in range(len(data_default_content)):
                data = data_default_content[i].text
                if data != "Dropped packets":
                    # Check if the data contains a hyphen
                    if '-' in data:
                        value_1 = data_default_values[value_index].get_attribute('value')  # Extract corresponding value
                        value_index += 1  # Increment index
                        and_part = ''  # Initialize 'and' part

                        # Check if there is an 'and' part
                        if 'and' in data:
                            and_part = ' and'
                            value_2 = data_default_values[value_index].get_attribute('value')
                            value_index += 1  # Increment index
                        # Format combined value based on the presence of 'and' part
                        combined_value = f"{data.split('-')[0].strip()}-{value_1}{and_part} -{value_2}" if and_part else f"{data.strip()}{value_1}"
                        combined_values.append(combined_value)
                        # Check if all values are processed
                        if value_index >= len(data_default_values):
                            break

                    elif "Sms status failure" in data:
                        values = data_default_values[value_index].get_attribute('value')
                        value_index += 1
                        combined_values.append(f"Less than {values} sec or Sms status failure")

                    elif any(keyword in data for keyword in ("ms", "mbps", "sec")) and not re.search(r'Less than \d+ sec', data):
                        # Split the data by "and" if it's present
                        parts = data.split('and')
                        # Get the values corresponding to the first part
                        value_1 = data_default_values[value_index].get_attribute('value')
                        # Extract the unit from the first part
                        unit_1 = ''.join(parts[0].split()[-1:])
                        # Remove the unit from the first part
                        data_1 = ' '.join(parts[0].split()[:-1])
                        # Add the value before the keyword with a space
                        combined_value = f"{data_1.strip()} {value_1} {unit_1}"
                        # If there is a second part (i.e., "and" is present)
                        if len(parts) > 1:
                            # Get the values corresponding to the second part
                            value_2 = data_default_values[value_index + 1].get_attribute('value')
                            # Extract the unit from the second part
                            unit_2 = ''.join(parts[1].split()[-1:])
                            # Remove the unit from the second part
                            data_2 = ' '.join(parts[1].split()[:-1])
                            # Add the value before the keyword with a space
                            and_part = ' and' if 'and' in data else ''
                            combined_value += f"{and_part} {data_2.strip()} {value_2} {unit_2}"
                            # Increment the value index
                            value_index += 1
                        combined_values.append(combined_value)
                        value_index += 1
                    else:
                        if 'and' in data:
                            parts = data.split('and')
                            if value_index + 1 < len(data_default_values):
                                value_1 = data_default_values[value_index].get_attribute('value')
                                value_2 = data_default_values[value_index + 1].get_attribute('value')
                                combined_value = f"{parts[0].strip()} {value_1} and {value_2}"
                                combined_values.append(combined_value)
                                value_index += 2
                        else:
                            if value_index < len(data_default_values):
                                value = data_default_values[value_index].get_attribute('value')
                                combined_values.append(f"{data.strip()} {value}")
                                value_index += 1
            if len(combined_values) != 0:
                print(combined_values)
                combine_dict[d] = combined_values
        except Exception as e:
            pass
    # Print or return the combined values
    print(combine_dict)

#########################################################################################################################################################################################
