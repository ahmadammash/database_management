Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True 
Set objWorkbook = objExcel.Workbooks.Open("D:\AmmashA.IA\OneDrive - University of Leeds\UOL Data Manager\database_management\data_files\leeds_data\convert_to_tbl_ranges.xlsm")
objExcel.Application.Run "convert_to_tbl_ranges.xlsm!ProcessAllExcelFiles"
objExcel.ActiveWorkbook.Save
objExcel.Application.Quit
WScript.Quit
