'xlsx转pdf
Option Explicit
Sub xlsxConverter()
On Error Resume Next
Dim sEveryFile As String,sSourcePath As String,sNewSavePath As String
Dim CurXls As Object
sSourcePath = "E:\XLSX文件\"  
'假定待转换的xlsx文件全部在"E:\XLSX文件\"下，你需要按实际情况修改。
sEveryFile = Dir(sSourcePath &"*.xlsx")
Do While sEveryFile <> ""
   Set Curxls = Workbooks.Open(sSourcePath & sEveryFile, , msoTrue )
   sNewSavePath = VBA.Strings.Replace(sSourcePath & sEveryFile, ".xlsx", ".pdf")
   '转化后的文件也在"E:\xlsx文件\"下，当然你可以按需修改。
   CurXls.ExportAsFixedFormat xlTypePDF,sNewSavePath
   '更多格式可参见文末的截图ExportAsFixedFormat
   CurXls.Close SaveChanges:=False
   sEveryFile= Dir
Loop
Set CurXls = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''
'xlsx转xls、csv
Option Explicit
Sub xlsxConverter()
On Error Resume Next
Dim sEveryFile As String,sSourcePath As String,sNewSavePath As String
Dim CurXls As Object
sSourcePath = "E:\XLSX文件\"  
'假定待转换的xlsx文件全部在"E:\XLSX文件\"下，你需要按实际情况修改。
sEveryFile = Dir(sSourcePath &"*.xlsx")
Do While sEveryFile <> ""
   Set Curxls = Workbooks.Open(sSourcePath & sEveryFile, ,msoTrue)
   sNewSavePath = VBA.Strings.Replace(sSourcePath & sEveryFile, ".xlsx", ".xls")
   '如果想导出csv，就把第12行行尾的xls换成csv
   '如果想把xls转为xlsx，把第9行的xlsx改为xls，把第12行行尾的".xlsx", ".xls"改为".xls", ".xlsx"
   '转化后的文件也在"E:\xlsx文件\"下，当然你可以按需修改。
   CurXls.SaveAs sNewSavePath, xlExcel8
   'xls对应xlExcel8,csv对应xlCSV,xlsx对应xlWorkbookDefault
   '更多格式可参见文末的截图XlFileFormat Enumeration (Excel)
   CurXls.Close SaveChanges:=False
   sEveryFile= Dir
Loop
Set CurXls = Nothing
End Sub