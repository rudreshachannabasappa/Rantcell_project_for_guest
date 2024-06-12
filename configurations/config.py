class ReadConfig():
    # Update the Root Folder Path where testdata and excelreport subfolders are created to store Test_Data.xlsx and Excel Reports
    ################# This path can be changed as per user need , If user need to change the path, use can change ####################################

    test_data_folder_rootpath = "C:\\RantCell_Automation_Data_and_Reports"

    #################################################### DON'T UPDATE ######################################################################
    # Absolute path of Test_Data.xlsx,Map_View_Component.xlsx,Graph_View_Component.xlsx,Pdf_Export,Table_summary,Parameter_validation,etc excel files
    test_data_path = test_data_folder_rootpath + "\\testdata\\Test_Data.xlsx"
    excel_report_path = test_data_folder_rootpath + "\\excelreport\\"
    map_view_components_excelpath = test_data_folder_rootpath + "\\testdata\\Map_View_Components.xlsx"
    graph_view_components_excelpath = test_data_folder_rootpath + "\\testdata\\Graph_View_Components.xlsx"
    source_dest = ["excelreport", "testdata","downloads"]
    excel_report_path_for_timestamp = test_data_folder_rootpath + "\\excelreport"
    excel_report_path_with_timestamp = test_data_folder_rootpath + "\\excelreport\\test_run_excelreport"
    updating_source_folders = ["testdata"]
    test_case_downloading_files_path=test_data_folder_rootpath + "\\downloads"
    test_case_downloading_files_path_timestamp=test_data_folder_rootpath + "\\downloads\\"
    test_run_excelreportdata_path = test_data_folder_rootpath+"\\excelreport\\TestRunFolderName.txt"
    test_run_download_file_path =test_data_folder_rootpath+"\\downloads\\TestRun_downloadfileFolderName.txt"
    pdf_export_excel_path = test_data_folder_rootpath+"\\testdata\\pdf_export.xlsx"
    parameter_validation_excel_path = test_data_folder_rootpath+"\\testdata\\Parameter_validation.xlsx"
    table_summary_excel_path = test_data_folder_rootpath+"\\testdata\\table_summary.xlsx"
    individual_popup_excel_path = test_data_folder_rootpath+"\\testdata\\Individualpopup.xlsx"
    nqc_testdata_excel_path = test_data_folder_rootpath+"\\testdata\\Nqctabledata.xlsx"
    remote_test_tesdata_path= test_data_folder_rootpath+"\\testdata\\Remote_test.xlsx"
    settings_path = test_data_folder_rootpath + "\\testdata\\Map_components_Settings.xlsx"
    ############################################################################################################################################
