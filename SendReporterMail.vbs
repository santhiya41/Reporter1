Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWBook = objExcel.Workbooks.Open("C:/Users/ssoundaram/Desktop/BatchReport_10-9-2019.xls")
Set ActiveSheet = objWBook.Worksheets("Dashboard")

'Get the Row Count
intRowCount = ActiveSheet.usedrange.rows.count

ActiveSheet.Range("A1:O"&Cstr(intRowCount)).Select

With ActiveSheet.MailEnvelope
	.Item.To = "ssoundaram@deloitte.com"
	.Item.CC = ""
	.Item.Subject = "SARA QA Test Results "&month(now)& "-" &day(now)& "-" & year(now)
	.Item.Send
End With

objWBook.save
objWBook.Close
objExcel.Quit

	