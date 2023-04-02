Set xlsxApp = CreateObject("Excel.Application")
	xlsxApp.visible = True

Set xlsxWorkbook = xlsxApp.Workbooks.Open("V1_20220628_CB.xlsm")
xlsxApp.Run("ExportSheetsToCSV")