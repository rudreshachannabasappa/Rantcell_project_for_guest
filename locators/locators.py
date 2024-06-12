from selenium.webdriver.common.by import By

class Login_Logout:
    link_login = (By.XPATH, "//a[normalize-space()='LOGIN']", "Login link")
    link_login1 = (By.XPATH, "//a[normalize-space()='LOGIN']")
    email = (By.ID, "email")
    textbox_username = (By.ID, "email", "Username")
    textbox_password = (By.ID, "password", "Password")
    button_login = (By.ID, "loginbutton", "Login button")
    dashboard = (By.XPATH, "//span[normalize-space()='Dashboard']", "Dashboard")
    dashboard_id = (By.ID,"refreshDashboard","Dashboard")
    dropdown_dropdown_toggle = (By.XPATH, "//span[@class='ng-binding']//i[@class='caret']", "Drop Down Toggle")
    link_logout = (By.XPATH, "//a[normalize-space()='Logout']", "Logout")
class side_menu_Components:
    campaignCheckBox = (By.XPATH, '//input[@type="checkbox"][@id="deSelectAll"]', "Campaign CheckBox")
    androidtestdata = (By.XPATH, "//span[text()='Android Test Data']", "Android Test Data")
    protestdata = (By.XPATH, "//span[text()='Pro TestData']", "Pro TestData")
    litetestdata = (By.XPATH, "//span[text()='LITE TestData']", "LITE TestData")
    device_element = (By.XPATH,"//li[@class='treeview active']//ul[@class='treeview-menu style-1 menu-open']//li")
    campaign_element = (By.XPATH,"//ul[@class='treeview-menu style-1 menu-open']//li[@class='treeview ng-scope active']//span")
    element = (By.XPATH,"//ul[@class='treeview-menu style-1 menu-open']//li[@class='treeview ng-scope active']//a[contains(text(),'Show More')]")
class select_Map_View_Components:
    Expand_Map_View = (By.XPATH, '//a[@id="exapandMap"]', "Expand Map View")
    Test_Type_Dropdown = (By.XPATH, "//*[@id='mainContent']/div[1]/div/div[1]/div/a[1]")
    Call_Test_locator = (By.XPATH, "//ul[@class='dropdown-menu']//li[@class='ng-scope']//a[@class='ng-binding'][contains(text(),'Call Test')]", "Call Test")
    nested_locators1 = [{"locator by": By.LINK_TEXT, "locator": "{}"}]
    cluster_blobmap_locator = (By.XPATH, "//div[@class ='cluster']")
    blobmap = (By.XPATH, "//div[@role ='button']//img[@src='https://maps.gstatic.com/mapfiles/transparent.png']")
    map_element = (By.XPATH, "//div[@class='box']//div[@class='box-body relative']")
    Data_Table = (By.XPATH, "//div[@class='info-window ng-scope']//table")
    export_selection_box = (By.XPATH, "//*[@id='exportselectionBox']")
    map_element_presence = (By.XPATH, "//div[@class='gm-style-mtc']/button[text()[contains(.,'Map')]]")
    map_element_verify = (By.XPATH, '//div[@class="gm-style-mtc"]/button[text()[contains(.,"Map")]]')
    satellite_element_presence = (By.XPATH, '//div[@class="gm-style-mtc"]/button[text()[contains(.,"Satellite")]]')
    satellite_element_verify = (By.XPATH, '//div[@class="gm-style-mtc"]/button[text()[contains(.,"Satellite")]]')
    map_menu_dropdown = (By.XPATH, "//div[@class='box-tools pull-right open']//ul[@role='menu']")
    graph_view_header = (By.XPATH, "//*[@id='mainContent']/div[1]/div[2]/section[2]/div/div/div[1]/h3")
    No_test_data_element = (By.XPATH,"// h3[contains(text(), 'No test data found. Please try different date and ')]")
class Map_view_Search_Box_not_visible_do_page_up:
    Search_Element =  (By.XPATH, "//input[@id='address']")

class Map_View_Select_and_ReadData:
    Call_Test_locator2 = (By.XPATH, "//a[@ng-show='tt.type'][normalize-space()='Call Test']", "Call Test")
    Failed_calls_locator = (By.XPATH,"//ul[@class='dropdown-menu']//li[@class='ng-scope']//a[@class='ng-binding'][contains(text(),'Failed Calls')]","Failed Calls")
    Test_Type_Dropdown_for_call_Test = (By.XPATH, "//*[@id='mainContent']/div[1]/div/div[1]/div/a[1]", "Test Type Dropdown for call_Test")

class close_button:
        closeFullTableView = (By.XPATH, "//a[@id='closeFullTableView']//button[@class='btn btn-box-tool btn-sm btn-default']", "close Full Table View")

class hover:
    canvas = (By.XPATH, '(//div[contains(@id, "chart")]/div/canvas[2])[position()=1]')
    Graph_Tootip_element = (By.XPATH, '//*[contains(@id, "chart")]/div/div[contains(@class, "tooltip")]')
    Dropdown_btn = (By.XPATH, '(//button[@data-toggle="dropdown"])[position()=1]')

class get_graph_data:
    graphdataTooltipElement = (By.XPATH, '//*[contains(@id, "chart")]/div/div[contains(@class, "tooltip")]')
    Dropdown_btn = (By.XPATH, '(//button[@data-toggle="dropdown"])[position()=1]')

class hover_over_second_graph:
    canvas = (By.XPATH, '(//div[contains(@id, "chart")]/div/canvas[2])[position()=2]')
    Graph_Tootip_element = (By.XPATH, '(//*[contains(@id, "chart")]/div/div[contains(@class, "tooltip")])[position()=2]')
    Dropdown_btn = (By.XPATH, '(//button[@data-toggle="dropdown"])[position()=2]')

class get_secondGraph_data:
    graphdataTooltipElement = (By.XPATH, '(//*[contains(@id, "chart")]/div/div[contains(@class, "tooltip")])[position()=2]')
    Dropdown_btn = (By.XPATH, '(//button[@data-toggle="dropdown"])[position()=2]')

class Graph_View_Components:
    ExpGraph = (By.XPATH, "//a[@id='fullview']//button[@class='btn btn-box-tool btn-sm btn-default pull-right']", "ExpGraph")
    closeFullView = (By.ID, "closeFullView", "close Full View")
    drop_down_toggle = (By.XPATH, '(//button[@data-toggle="dropdown"])[position()=1]')
    Graph_dropdown_list_txt = (By.XPATH,"//div[@class='box-tools pull-right box-info open']//ul[@role='menu']")
    second_graph_position = (By.XPATH, '(//div[contains(@id, "chart")]/div/canvas[2])[position()=2]')
    By_closeFullView = (By.ID, "closeFullView")

class expand_tableView_verify_popUp:
    DataTable_Expand = (By.ID, 'expandTableview', "Data Table Expand")
    DataTable_Close = (By.ID, "closeFullTableView", "Data Table Close")
    Campaigns_Name = (By.XPATH, "//*[@id='loaderCamp']/abbr/a", "Campaigns Name")
    Close_Pop_Up_Data_point_Btn = (By.XPATH, "//div[@id='myModal']//button[@type='button'][normalize-space()='Close']", "Close Pop Up Data point Btn")
    pop_up_data_element = (By.XPATH, '//*[@id="loaderCamp"]/abbr/a')

class List_Of_Campaigns_components_Search_Box_not_visible_do_page_up:
    Search_Element = (By.XPATH, "//input[@placeholder='search ...']")

class List_Of_Campaigns_Export_Dashboard:
    List_Of_Campaigns_Export_Dropdown = (By.XPATH, "//*[@id='tableView']//select")
    List_Of_Campaigns_Export_Dropdown_Options = [{"locator by": By.XPATH, "locator": "//option[contains(text(),'{}')]"}]

class operator_comparison_table:
    Operator_comparison_data = (By.XPATH, "//div[@id='scrollId']")
    operator_comparison_web = (By.XPATH, "//table[@id='webtestTableFixed']")
    operator_comparison_web_siblingtable = (By.XPATH, "//table[@id='webtestTableFixed']/following-sibling::table")

class pdf_view:
    parent_checkbox_pdf = (By.XPATH, "//div[@id='checkboxes']")
    save_pdf_export = (By.XPATH, "//a[@ng-click='saveAsPDF()']", "save As PDF")
    generate_report_pdf = (By.XPATH, "//a[normalize-space()='Generating report..']")
    Export_pdf_table = [{"locator by": By.XPATH, "locator": "//div[@id ='{}']"}]
class table_summary:
    header_table_summary = (By.XPATH,"//div[@class='div-table-content-wrapper']//table[@class='table table-condensed table-borderedWithoutWordWrap table-hover table-striped']")
    content_table_summary = (By.XPATH, "//div[@class='div-table-content']//table")
class nw_freeze:
    version = (By.XPATH, "//b[normalize-space()='Version']")
class load_locators:
    loading_locators = [
            (By.XPATH, '//*[@id="loadingImgtestdata"]//div//img'),
            (By.XPATH, '//*[@id="loading_img"]//i'),
            (By.XPATH, '//*[@id="loaderFast"]//i')
        ]
class individual_pop_table():
    loaderCamp = (By.XPATH, '//*[@id="loaderCamp"]/abbr/a', "loaderCamp")
    Tableview_header = (By.XPATH, "//div[@id='myModal']//div[@class='modal-body']//div[@id='hideTableview']//div//h4")
    Tableview_siblingtable_header = (By.XPATH, "./following-sibling::div/table")
    Tableview_siblingtable_content = (By.XPATH, "./following-sibling::div/div/table")
    Close_button_of_ipu = (By.XPATH, "//div[@class='modal-content']//button[@id='closemodal']", "Close button of ipu")

class remote_test:
    remotetest = (By.XPATH, "//span[normalize-space()='Remote Test']", "Remote Test")
    android_pro_is_inactive = (By.XPATH, "//li[@ng-click='deviceTypeTabSeperation('AndroidPro')'][not(@class='active')]")
    check_devices = (By.XPATH, "./following-sibling::ul//a[contains(text(),'Check')]")
    check_device_tab = (By.XPATH, "//div[@id='checkdevices']//div[@class='modal-body']")
    timer_xpath = (By.XPATH, "//div[@class='btn btn-default pull-right']//timer[contains(@class, 'ng-binding')]")
    timerspan_xpath = (By.XPATH, "//div[@class='btn btn-default pull-right']//timer[contains(@class, 'ng-binding')]//span")
    online_or_offline_status = (By.XPATH, "//td[normalize-space()='Offline' or normalize-space()='Online']")
    status_of_devices_Offline = (By.XPATH, "//td[@ng-if='!checkdevice.Status'][normalize-space()='Offline']")
    checkdevice = "checkdevice.Status == 'Delivered'"
    status_of_devices_Online = (By.XPATH, f'//td[@ng-if="{checkdevice}"][normalize-space()="Online"]')
    device_name = (By.XPATH,"//div[@id='checkdevices']//i[@class='fa fa-mobile']/parent::td")

    Run_Test = (By.XPATH, "./following-sibling::ul//a[contains(text(),'Run Test')]")
    closebtn_ofcheckdevice = (By.XPATH,"//div[(@id='checkdevicesFailed' and @style='display: block;') or (@id='checkdevices' and @style='display: block;')]//button[@type='button'][normalize-space()='Close']","closebtn_ofcheckdevice")
    closebtn_ofrun_test = (By.XPATH, "//div[@id='scheduledtask']//a[@type='button'][normalize-space()='Close']", "closebtn_ofrun_test")
    run_test_tab = (By.XPATH, "//div[@id='scheduledtask']//h4[@class='modal-title text-info'][normalize-space()='Run Test']")
    test_name = (By.XPATH, "//form[@id='scheduleForm']//input[@id='testname']", "Test name text box")
    iteration_textbox = (By.XPATH,"//div[@class='form-group has-feedback']//div[@class='col-xs-6 col-sm-4 col-md-4 col-lg-4']//input[@id='iterations']","Iteration_textbox")
    delays_bw_tests = (By.XPATH, "//form[@id='scheduleForm']//input[@id='Delay']", "Delays_bw_tests")
    startbtn_ofrun_test = (By.XPATH, "//a[@ng-disabled='scheduleForm.$invalid || !disableIfcheckboxisSelected()']", "startbtn_ofrun_test")

    ping_test_checkbox = (By.ID, "pingtestChk", "Ping_test_checkbox")
    ping_test_form = (By.XPATH, "//form[@id='pingtestModal']//div//label")
    host_textbox = (By.XPATH, "./following-sibling::div//input[@id='host']", "Host_textbox")
    packetsize_textbox = (By.XPATH, "./following-sibling::div//input[@id='packetsize']", "packetsize_textbox")
    pingtest_okbtn = (By.XPATH, "//a[@ng-disabled='pingtestForm.host.$invalid || pingtestForm.packetsize.$invalid ']", "Pingtest_okbtn")
    pingtest_closebtn = (By.XPATH, "//div[@id='pingtest']//a[@type='button'][normalize-space()='close']", "pingtest_closebtn")

    call_test_checkbox = (By.ID, "calltestChk", "Call_test_checkbox")
    call_test_form = (By.XPATH, "//form[@id='calltestModal']//label")
    Call_B_Party_Phone_Number_textbox = (By.XPATH, "./following-sibling::div/input[@id='number']", "B_Party_Phone_Number_textbox")
    Call_Duration_textbox = (By.XPATH, "./following-sibling::div/input[@id='duration']", "Call_Duration_textbox")
    calltest_okbtn = (By.XPATH, "//a[@ng-disabled='calltestForm.$invalid && !disableCallTest()']", "Calltest_okbtn")
    calltest_closebtn = (By.XPATH, "//div[@id='calltest']//a[@type='button'][normalize-space()='close']", "calltest_closebtn")

    sms_test_checkbox = (By.ID, "smstestChk", "Sms_test_checkbox")
    sms_test_form = (By.XPATH, "//div[@id='smstest']//form[@id='smstestModal']//label")
    sms_B_Party_Phone_Number_textbox = (By.XPATH, "./following-sibling::div//input[@id='smsnumber']", "B_Party_Phone_Number_textbox")
    sms_Wait_Duration_textbox = (By.XPATH, "./following-sibling::div//input[@id='smsduration']", "Wait_Duration_textbox")
    smstest_okbtn = (By.XPATH, "//a[@ng-disabled='smstestForm.$invalid && !disableSmsTest()']", "Smstest_okbtn")
    smstest_closebtn = (By.XPATH, "//div[@id='smstest']//a[@type='button'][normalize-space()='close']", "smstest_closebtn")

    speed_test_checkbox = (By.ID, "speedtestChk", "Speed_test_checkbox")
    speed_test_form = (By.XPATH, "//form[@id='speedtestModal']//label")
    speed_Use_Default_Server_checkbox = (By.XPATH, "./following-sibling::div//input[@id='defaultserver']", "Use_Default_Server_checkbox")
    Select_Download_Test_File_Size_dropdown = (By.ID, "downloadfilesize", "Select_Download_Test_File_Size_dropdown")
    FTP_Server_textbox = (By.XPATH, "./following-sibling::div//input[@id='ftp']", "FTP_Server_textbox")
    UserName_textbox = (By.XPATH, "./following-sibling::div//input[@id='ftpusername']", "UserName_textbox")
    Password_textbox = (By.XPATH, "./following-sibling::div//input[@id='ftppassword']", "Password_textbox")
    Download_File_Name_textbox = (By.XPATH, "./following-sibling::div//input[@id='downloadfilename']", "Download_File_Name_textbox")
    Enable_Upload_Test_checkbox = (By.XPATH, "./following-sibling::div//input[@id='uploadtest']", "Enable_Upload_Test_checkbox")
    speed_File_Size_textbox = (By.XPATH, "./following-sibling::div//input[@id='filesize']", "File_Size_textbox")
    Enable_FTP_Stop_Timer_checkbox = (By.XPATH, "./following-sibling::div//input[@id='ftpstoptimer']", "Enable_FTP_Stop_Timer_checkbox")
    speed_Wait_Duration_textbox = (By.XPATH, "./following-sibling::div//input[@id='ftptimetorun']", "Wait_Duration_textbox")
    speedtest_okbtn = (By.XPATH, "//a[@ng-disabled='!speedtestValidate()']", "speedtest_okbtn")
    speedtest_closebtn = (By.XPATH, "//div[@id='speedtest']//a[@type='button'][normalize-space()='close']", "speedtest_closebtn")

    http_speed_test_checkbox = (By.ID, "httpspeedtestChk", "Http_Speed_test_checkbox")
    http_speed_test_form = (By.XPATH, "//div[@id='httpspeedtest']//form[@id='httpspeedtestModal']//label")
    Enter_custom_URL_checkbox = (By.XPATH, "//div[@id='httpspeedtest']//input[@id='defaulturl']", "Enter_custom_URL_checkbox")
    Enter_URL_textbox = (By.XPATH, "//div[@class='form-group']//input[@id='http']", "Enter_URL_textbox")
    HTTP_Speed_Download_Test_File_Size_dropdown = (By.XPATH, "//div[@id='httpspeedtest']//select[@id='httpdownloadfilesize']","HTTP_Speed_Download_Test_File_Size_dropdown")
    Enable_HTTP_Speed_Test_Upload_Test_checkbox = (By.XPATH, "//div[@id='httpspeedtest']//input[@id='httpuploadtest']", "Enable_HTTP_Speed_Test_Upload_Test_checkbox")
    http_speed_File_Size_textbox = (By.XPATH, "//div[@class='form-group']//input[@id='httpuploadfilesize']", "File_Size_textbox")
    Enable_HTTP_Speed_Test_stop_timer_checkbox = (By.XPATH, "//input[@id='httpstoptimer']", "Enable_HTTP_Speed_Test_stop_timer_checkbox")
    HTTP_Speed_Set_Timeout_textbox = (By.ID, "httptimetorun", "Set_Timeout_textbox")
    Enter_Custom_Upload_URL_checkbox = (By.XPATH, "//input[@ng-model='httpCustomUploadUrl.customUploadCheck']", "Enter_Custom_Upload_URL_checkbox")
    Upload_URL_textbox = (By.XPATH, "//input[@ng-model='httpCustomUploadUrl.url']", "Upload_URL_textbox")
    http_speedtest_okbtn = (By.ID, "httpsubmitBtn", "http_speedtest_okbtn")
    http_speedtest_closebtn = (By.XPATH, "//div[@id='httpspeedtest']//a[@type='button'][normalize-space()='close']", "http_speedtest_closebtn")

    iperf_testcheckbox = (By.ID, "iperfChk", "iperf_test_checkbox")
    iperf_test_form = (By.XPATH, "//div[@id='iperftest']//form[@id='Iperftestmodal']//label")
    Iperf_Use_Default_Server_checkbox = (By.XPATH, "//div[@id='iperftest']//input[@id='defaultserver']", "Use_Default_Server_checkbox")
    Host_Name_textbox = (By.XPATH, "//div[@ng-hide='checkboxIperfDefaultServer']//input[@id='ftp']", "Host_Name_textbox")
    Test_Duration_textbox = (By.XPATH, "//div[@id='iperftest']//input[@placeholder='Enter Test duration']", "Test_Duration_textbox")
    Enable_Iperf_Upload_Test_checkbox = (By.XPATH, "//input[@ng-model='checkboxIperfvalueUpload']", "Enable_Iperf_Upload_Test_checkbox")
    TCP_Mode_checkbox = (By.XPATH, "//input[@ng-model='checkboxIperfvalueTCP']", "TCP_Mode_checkbox")
    UDP_Mode_checkbox = (By.XPATH, "//input[@ng-model='checkboxIperfvalueUDP']", "UDP_Mode_checkbox")
    Enter_the_Bandwidth_textbox = (By.XPATH, "//div[@ng-show='checkboxIperfvalueUDP']//input[@name='IperfBandwidth']", "Enter_the_Bandwidth_textbox")
    iperf_test_okbtn = (By.XPATH, "//a[@ng-disabled='Iperftestmodal.$invalid || !iperfTestValidate()']", "iperf_test_okbtn")
    iperf_test_closebtn = (By.XPATH, "//a[@ng-click='cancelIperf()']", "iperf_test_closebtn")

    stream_checkbox = (By.XPATH, "//input[@id='streamtestChk']", "Stream Test")
    txt_box_url = (By.XPATH, "//input[@id='streamtesturl']", "video URL")
    submit_ok_btn = (By.XPATH, "//a[@id='streamsubmitBtn']", "Ok Button")
    enter_url_checkbox = (By.XPATH, "//input[@name='defaultstreamurlchk']","Url")
    run_test_close_btn = (By.XPATH, "//div[@id='scheduledtask']//a[@type='button'][normalize-space()='Close']", "Run Test Close Button")

    # webtest_checkbox = (By.XPATH, "//form[@name='scheduleForm']//div//div//label[contains(text(),'Web Test:')]","webtest checkbox")
    webtest_checkbox = (By.ID, "webtestChk","webtest checkbox")
    web_url = (By.XPATH, "//input[@ng-model='user.WebtestUrl']", "input URL")
    web_test_okbtn = (By.XPATH, "//div[@id='webtest']//a[@type='button'][normalize-space()='OK']", "web_test_okbtn")

    run_test_start_status = (By.XPATH,"//div[(@id='deliveryReportRefreshModal' and @style='display: block;') or (@id='deliveryReportModal' and @style='display: block;')]//td[starts-with(normalize-space(),'Failed') or starts-with(normalize-space(),'Success')]","run_test_start_status")
    run_test_start_statusclose_btn = (By.XPATH,"//div[(@id='deliveryReportRefreshModal' and @style='display: block;') or (@id='deliveryReportModal' and @style='display: block;')]//button[@type='button'][normalize-space()='Close']","run_test_start_statusclose_btn")
    run_test_start_status_popup = (By.XPATH,"//div[@id='deliveryReportRefreshModal' or @id='deliveryReportModal']//div[@class='modal-body']")

    table_view_refresh = (By.XPATH,"//div[@id='tableView']//a[normalize-space()='Refresh']","table_view_refresh")

class add_test_group_1:
    test_group_btn = (By.XPATH, "//a[normalize-space()='Add Test Group']", "Add Test Group")
    group_name = (By.XPATH, "//input[@placeholder='Enter Group Name']", "Test Group Name")
    next_button = (By.XPATH, "//a[@id='nextBtn']", "Next Button")
    add_button = (By.XPATH, "//a[@id='savebutton']", "Add Button")

class edit_group:
    edit_group_btn = (By.XPATH, "//div[@class='btn-group open']//a[@data-toggle='dropdown'][normalize-space()='Edit Group']")
    add_device_btn = (By.XPATH, "//div[@class='btn-group open']//ul[@id='responsivesubmenu']//a[contains(text(),'Add')]", "Add Device")
    add_button = (By.XPATH, "//a[@ng-disabled='addDeviceForm.$invalid']", "Add Button")
    delete_device_btn = (By.XPATH, "//div[@class='btn-group open']//a[contains(text(),'Delete Devices')]", "delete_device_btn")
    delete_btn = (By.XPATH, "//button[@ng-click='deleteSelecteddevices($event)']", "delete button")
    delete_group_btn_in_edit_group = (By.XPATH, "//div[@class='btn-group open']//ul[@id='responsivesubmenu']//li[5]//a[1]", "delete_group_btn")
    delete_group_btn = (By.XPATH, "//button[normalize-space()='Delete Group']", "delete_group_btn")

class continuous:
    continuous_hover = (By.XPATH, "//div[@class='btn-group open']//a[@data-toggle='dropdown'][normalize-space()='Continuous Test']")
    add_continuous_button = (By.XPATH, "./following-sibling::ul/li/a[contains(.,'Add')]", "Add continuous test")
    delete_continuous_button = (By.XPATH,"./following-sibling::ul/li/a[contains(.,'Delete')]","Delete continuous test")
    test_name_field = (By.XPATH, "//form[@id='continuoustestForm']//input[@id='testname']", "Test name text box")
    ping_checkbox = (By.ID, "pingtestChk2", "Ping Test Checkbox")
    call_checkbox = (By.ID, "calltestChk2", "Call Test Checkbox")
    speed_checkbox = (By.ID, "speedtestChk2", "Speed Test Checkbox")
    http_checkbox = (By.ID, "httpspeedtestChk2", "HTTP Test Checkbox")
    sms_checkbox = (By.ID, "smstestChk2", "Sms Test Checkbox")
    stream_checkbox = (By.ID, "streamtestChk2", "Stream Test Checkbox")
    web_checkbox = (By.ID, "webtestChk2", "Web Test Checkbox")
    iperf_checkbox = (By.ID, "iperfTestChk2", "Iperf Test Checkbox")
    multi_bparty_checkbox = (By.ID,"enableCallContTest","Multi Bparty test Checkbox")
    multi_bparty_test_form = (By.XPATH,"//div[@id='callcontinuoustest']//form[@id='callconttestModal']//label")
    multi_bparty_phone_number = (By.XPATH,"./following-sibling::div/input[@id='callcontnumber']","Multi bparty phone number")
    call_duration = (By.XPATH ,"./following-sibling::div/input[@id='callcontduration']","Call duration")
    multi_bparty_ok_btn = (By.XPATH ,"//div[@id='callcontinuoustest']//a[@type='submit'][normalize-space()='OK']","Multi bparty ok btn")
    multi_bparty_closebtn = (By.XPATH,"//div[@id='callcontinuoustest']//a[@type='button'][normalize-space()='close']", "Multi bparty close btn" )
    run_button = (By.XPATH, "//div[@id='continuoustest']//a[@type='submit'][normalize-space()='Run']", "Run Button")
    close_button = (By.XPATH, "//a[@ng-click='cancelContinuoustest()']", "Close Button")
class schedule_test:
    schedule_test_btn=(By.XPATH,"//body/div[@class='wrapper']/div[@class='ng-scope']/div[@class='content-wrapper ng-scope']/section[@id='mainContent']/div[@class='ng-scope']/section[@class='content responsiveRemoteTest ng-scope']/section/div[@class='row']/div[@class='col-lg-12 col-md-12 col-sm-12']/div[@class='tab-content']/div[@id='androidpro']/section/section[@class='content responsiveRemoteTest']/div[@id='tableView']/div[@class='box-body']/div[@class='row']/div[@class='col-lg-3 ng-scope']/div[@class='box box-primary']/div[@class='box-header']/div[@class='box-tools']/div[@class='btn-group open']/ul[@role='menu']/li[9]/a[1]","schedule_test_btn")
    increse_time_up_arrow_btn=(By.XPATH,"//a[@title='Increment Minute']//span[@class='glyphicon glyphicon-chevron-up']","increse_time_up_arrow_btn")
    test_name = (By.XPATH, "//form[@id='scheduletestForm']//input[@id='testname']", "Test name text box")
    iteration_textbox= (By.XPATH,"//form[@id='scheduletestForm']//input[@id='iterations']","iteration_textbox")
    delays_bw_tests = (By.XPATH, "//form[@id='scheduletestForm']//input[@id='Delay']", "Delays_bw_tests")
    ping_test_checkbox = (By.ID, "pingtestChk1", "Ping_test_checkbox")
    call_test_checkbox = (By.ID, "calltestChk1", "Call Test Checkbox")
    speed_test_checkbox = (By.ID, "speedtestChk1", "Speed Test Checkbox")
    http_speed_test_checkbox = (By.ID, "httpspeedtestChk1", "HTTP Test Checkbox")
    sms_test_checkbox = (By.ID, "smstestChk1", "Sms Test Checkbox")
    stream_test_checkbox = (By.ID, "streamtestChk1", "Stream Test Checkbox")
    web_test_checkbox = (By.ID, "webtestChk1", "Web Test Checkbox")
    iperf_test_checkbox = (By.ID, "iperfChk1", "Iperf Test Checkbox")
    run_button = (By.XPATH, "//div[@id='scheduletest']//a[@type='submit'][normalize-space()='Run']", "Run Button")
    close_button = (By.XPATH, "//div[@id='scheduletest']//a[@type='button'][normalize-space()='Close']", "Close Button")

#################################################################################################################################
class settings_1:
    ######################################################### settings section xpaths ###############################################
    default_settings_btn = (By.XPATH, "//a[normalize-space()='DefaultSettings']", "Default Settings Button")
    default_settings1 = (By.XPATH, "//a[normalize-space()='DefaultSettings']")
    save_settings_btn = (By.XPATH, "//button[normalize-space()='Save Settings']", "Save Settings Button")
    btn_setting = (By.XPATH, "//span[normalize-space()='Settings']", "Settings Button")
    all_default_settings_headers = (By.XPATH,"//section[@class='col-lg-7 col-lg-offset-1 connectedSortable']//div[starts-with(@id, 'anchor')]//div//h5[1][not(contains(text(),'Device alarms') or contains(text(),'Threshold Crossing Alarms') or contains(text(),'Continuous Test Failure Alarms') or contains(text(),'Cell Identification Options') or contains(text(),'Mobile network coverage type (Data Type)') or contains(text(),'WebTest') or text()='Stream Test' or contains(text(),'SMS Test') or contains(text(),'Call Test'))]")
    all_default_settings_content = (By.XPATH,"./parent::div/following-sibling::div//div[@class='tableContent']")
    all_default_settings_values = (By.XPATH, "./parent::div/following-sibling::div//div[@class='tableContent']//input")
    ###################################################### operator comparison and map legend xpaths ###############################
    operator_comparison_data = (By.XPATH,"//div[@id='scrollId']//tbody//tr//td[1]//span")
    map_legend_each_elements = (By.XPATH,"//div[@id='mapLegend']//a")
    ##################################################### pdf section xpaths ###################################################################################
    pdf_btn_setting = (By.XPATH, "//option[contains(text(),'Export As PDF')]", "PDF Button")
    ping_pdf = (By.XPATH,"//div[@id ='tableViewPing']//td[1]//span")
    download_test = (By.XPATH,"//div[@id ='tableViewdownload']//td[1]//span")
    upload_test = (By.XPATH, "//div[@id ='tableViewupload']//td[1]//span")
    http_download_test = (By.XPATH, "//div[@id ='pdfHttDownload']//td[1]//span")
    http_upload_test = (By.XPATH, "//div[@id ='httpUploadtableview']//td[1]//span")
    tcpiperf_download_test = (By.XPATH, "//div[@id ='pdfIperfTcpDownloadNQC']//td[1]//span")
    tcpiperf_upload_test = (By.XPATH, "//div[@id ='pdfIperfTcpUploadNQC']//td[1]//span")
    call_test = (By.XPATH, "//div[@id ='callTestTableview']//td[1]//span")
    stream_test = (By.XPATH, "//div[@id ='streamTestTableView']//td[1]//span")
    rssi_rscp_test = (By.XPATH, "//div[@id ='rssitableview']//td[1]//span")
    rsrp_test = (By.XPATH, "//div[@id ='rsrptableview']//td[1]//span")
    rsrq_test = (By.XPATH, "//div[@id ='rsrqtableview']//td[1]//span")
    nrssrsrp_test = (By.XPATH, "//div[@id ='nrSsRsrptableview']//td[1]//span")
    nrssrsrq_test = (By.XPATH, "//div[@id ='nrSsRsrqtableview']//td[1]//span")
    ltesnr_test = (By.XPATH, "//div[@id ='lteSNRtableview']//td[1]//span")
    nrsssinr_test = (By.XPATH, "//div[@id ='nrSsSinrtableview']//td[1]//span")

    ######################################################### Change Settings Locator ###############################################################################
    rssi_rscp_1 = (By.XPATH, "//input[@name='rssigreater']", "rssi_rscp_1")
    rssi_rscp_2 = (By.XPATH, "//input[@id='rssiNormal']", "rssi_rscp_2")
    rssi_rscp_3 = (By.XPATH, "//input[@id='rssiLower']", "rssi_rscp_3")

    wifi_rssi_1 = (By.XPATH, "//input[@name='wifirssigreater']", "wifi_rssi_1")
    wifi_rssi_2 = (By.XPATH, "//input[@id='wifirssiNormal']", "wifi_rssi_2")
    wifi_rssi_3 = (By.XPATH, "//input[@id='wifirssiLower']", "wifi_rssi_3")

    rsrp_1 = (By.XPATH, "//input[@name='rsrpGreater']", "rsrp_1")
    rsrp_2 = (By.XPATH, "//input[@id='rsrpNormal']", "rsrp_2")
    rsrp_3 = (By.XPATH, "//input[@id='rsrpLower']", "rsrp_3")

    rsrq_1 = (By.XPATH, "//input[@name='rsrqGreater']", "rsrq_1")
    rsrq_2 = (By.XPATH, "//input[@id='rsrqNormal']", "rsrq_2")

    ltesnr_1 = (By.XPATH, "(//input[@name='forestgreen'])[1]", "ltesnr_1")
    ltesnr_2 = (By.XPATH, "(//input[@name='applegreen1'])[1]", "ltesnr_2")
    ltesnr_3 = (By.XPATH, "(//input[@name='palegreen1'])[1]", "ltesnr_3")
    ltesnr_4 = (By.XPATH, "(//input[@name='blue1'])[1]", "ltesnr_4")
    ltesnr_5 = (By.XPATH, "(//input[@name='brown1'])[1]", "ltesnr_5")
    ltesnr_6 = (By.XPATH, "(//input[@name='orange1'])[1]", "ltesnr_6")
    ltesnr_7 = (By.XPATH, "(//input[@name='snrAmber1'])[1]", "ltesnr_7")

    cdma_rssi_1 = (By.XPATH, "//input[@name='cdmarssiGreater']", "cdma_rssi_1")
    cdma_rssi_2 = (By.XPATH, "//input[@id='cdmarssiNormal']", "cdma_rssi_2")
    cdma_rssi_3 = (By.XPATH, "//input[@id='cdmarssiLower']", "cdma_rssi_3")

    ecno_1 = (By.XPATH, "//input[@name='cdmaecnoGreater']", "ecno_1")
    ecno_2 = (By.XPATH, "//input[@id='cdmaecnoNormal']", "ecno_2")

    cdma_snr_1 = (By.XPATH, "//input[@name='cdmasnrGreater']", "cdma_snr_1")
    cdma_snr_2 = (By.XPATH, "//input[@id='cdmasnrNormal']", "cdma_snr_2")

    nrSsSINR_1 = (By.XPATH, "(//input[@name='forestgreen'])[2]", "nrSsSINR_1")
    nrSsSINR_2 = (By.XPATH, "(//input[@name='applegreen1'])[2]", "nrSsSINR_2")
    nrSsSINR_3 = (By.XPATH, "(//input[@name='palegreen1'])[2]", "nrSsSINR_3")
    nrSsSINR_4 = (By.XPATH, "(//input[@name='blue1'])[2]", "nrSsSINR_4")
    nrSsSINR_5 = (By.XPATH, "(//input[@name='brown1'])[2]", "nrSsSINR_5")
    nrSsSINR_6 = (By.XPATH, "(//input[@name='orange1'])[2]", "nrSsSINR_6")
    nrSsSINR_7 = (By.XPATH, "(//input[@name='snrAmber1'])[2]", "nrSsSINR_7")

    nrSsRSRP_1 = (By.XPATH, "//input[@name='nrSsrsrpGreater']", "nrSsRSRP_1")
    nrSsRSRP_2 = (By.XPATH, "//input[@id='nrSsrsrpNormal']", "nrSsRSRP_2")
    nrSsRSRP_3 = (By.XPATH, "//input[@id='nrSsrsrpLower']", "nrSsRSRP_3")

    nrSsRSRQ_1 = (By.XPATH, "//input[@name='nrSsrsrqGreater']", "nrSsRSRQ_1")
    nrSsRSRQ_2 = (By.XPATH, "//input[@id='nrSsrsrqNormal']", "nrSsRSRQ_2")

    ping_1 = (By.XPATH, "//input[@name='pinggreater']", "ping_1")

    call_setup_time_1 = (By.XPATH, "//input[@name='callsetupGreater']", "call_setup_time_1")
    call_setup_time_2 = (By.XPATH, "//input[@id='callsetupnormal']", "call_setup_time_2")

    sms_sent_received_1 = (By.XPATH, "//input[@name='smssetupGreater']", "sms_sent_received_1")
    sms_sent_received_2 = (By.XPATH, "//input[@id='smssetupnormal']", "sms_sent_received_2")

    download_http_iperf_1 = (By.XPATH, "//input[@name='speedtestdownloadGreater']", "download_http_iperf_1")
    download_http_iperf_2 = (By.XPATH, "//input[@id='speedtestdownloadNormal']", "download_http_iperf_2")

    upload_http_iperf_1 = (By.XPATH, "//input[@name='speedtestuploadGreater']", "upload_http_iperf_1")
    upload_http_iperf_2 = (By.XPATH, "//input[@id='speedtestuploadNormal']", "upload_http_iperf_2")

    stream_1 = (By.XPATH, "//input[@name='streamtestdownloadGreater']", "stream_1")
    stream_2 = (By.XPATH, "//input[@id='streamtestdownloadNormal']", "stream_2")

class date_time:
    date_and_time_click_button=(By.XPATH,"//span[normalize-space()='Date and Time']","Date_and_Time")
    expand_table_button =(By.XPATH,"//a[@id='expandTableview']//button[@class='btn btn-box-tool btn-sm btn-default']","expand _table_button")
