'docx转pdf、doc、rtf、txt
Option Explicit
Sub docx2other()
On Error Resume Next
Dim sEveryFile As String,sSourcePath As String,sNewSavePath As String
Dim CurDoc As Object
sSourcePath = "E:\DOCX文件\"  
'假定待转换的docx文件全部在"E:\DOCX文件\"下，你需要按实际情况修改。
sEveryFile = Dir(sSourcePath &"*.docx")
Do While sEveryFile <> ""
   Set CurDoc = Documents.Open(sSourcePath & sEveryFile, , , , , , , , , , , msoFalse)
   sNewSavePath = VBA.Strings.Replace(sSourcePath & sEveryFile, ".docx", ".pdf")
   '如果想导出doc/rtf/txt等，就把上一行行尾的pdf换成doc/rtf/txt
   '转化后的文件也在"E:\DOCX文件\"下，当然你可以按需修改。
   CurDoc.SaveAs2 sNewSavePath, wdFormatPDF
   'pdf对应wdFormatPDF,doc对应wdFormatDocument,rtf对应wdFormatRTF,txt对应wdFormatText
   '更多格式可参见文末的截图WdSaveFormat Enumeration
   CurDoc.Close SaveChanges:=False
   sEveryFile= Dir
Loop
Set CurDoc = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''
'pdf、doc、rtf、txt转docx
Option Explicit
Sub other2docx()
On Error Resume Next
Dim sEveryFile As String,sSourcePath As String,sNewSavePath As String
Dim CurDoc As Object
sSourcePath = "E:\PDF文件\"
'假定待转换的pdf文件全部在"E:\PDF文件\"下，你需要按实际情况修改。
sEveryFile = Dir(sSourcePath &"*.pdf")
Do While sEveryFile <> ""
   Set CurDoc = Documents.Open(sSourcePath & sEveryFile, , , , , , , , , , , msoFalse)
   CurDoc.Convert 
   sNewSavePath = VBA.Strings.Replace(sSourcePath & sEveryFile, ".pdf", ".docx")
   '要把doc/rtf/txt转为docx，则把上面第9行和第13行两处".pdf"改为".doc"/".rtf"/".txt"
   '转化后的文件也在"E:\PDF文件\"下，当然你可以按需修改。
   CurDoc.SaveAs2 sNewSavePath, wdFormatDocumentDefault
   CurDoc.Close SaveChanges:=False
   sEveryFile = Dir
Loop
Set CurDoc = Nothing
End Sub