'pptx转pdf、ppt
Option Explicit
Sub pptxConverter()
On Error Resume Next
Dim sEveryFile As String,sSourcePath As String,sNewSavePath As String
Dim CurPpt As Object
sSourcePath = "E:\PPTX文件\"  
'假定待转换的pptx文件全部在"E:\PPTX文件\"下，你需要按实际情况修改。
sEveryFile = Dir(sSourcePath &"*.pptx")
Do While sEveryFile <> ""
   Set CurPpt = Presentations.Open(sSourcePath & sEveryFile, msoTrue , , msoFalse)
   sNewSavePath = VBA.Strings.Replace(sSourcePath & sEveryFile, ".pptx", ".pdf")
   '如果想导出ppt，就把第12行行尾的pdf换成ppt
   '如果想把ppt转为pptx，把第9行的pptx改为ppt，把第12行行尾的 ".pptx", ".pdf"改为 ".ppt", ".pptx"
   '转化后的文件也在"E:\PPTX文件\"下，当然你可以按需修改。
   CurPpt.SaveAs sNewSavePath, ppSaveAsPDF
   'pdf对应ppSaveAsPDF,ppt对应ppSaveAsPresentation,pptx对应ppSaveAsDefault
   '更多格式可参见文末的截图PpSaveAsFileType
   CurPpt.Close SaveChanges:=False
   sEveryFile= Dir
Loop
Set CurPpt = Nothing
End Sub