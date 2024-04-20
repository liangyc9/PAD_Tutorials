Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("%SelectedFile%")
Set objWorksheet = objWorkbook.Worksheets(1)  
Set objRange = objWorksheet.Range("B5:C8")
objRange.Select
'Add a chart
Set colCharts = objExcel.Charts
colCharts.Add()
Set objChart = colCharts(1)
objChart.ChartType = 5
objChart.Activate
objChart.HasLegend = TRUE
objChart.ChartTitle.Text = "金融資產派狀圖分析"