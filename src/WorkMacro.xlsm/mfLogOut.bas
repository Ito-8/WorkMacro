Attribute VB_Name = "mfLogOut"
'参照設定 : Microsoft scripting runtime
'参照設定の自動設定：http://kouten0430.hatenablog.com/entry/2017/10/21/152301


'Dim wsLog As Worksheet



Sub mfLogOutInitialize()
    
    ' Set wsLog = wb.Worksheets.Add(after:=Sheets(wb.Worksheets.Count))
    
    Dim LogFileName As String
    
    LogFileName = "Log_" & Now
    LogFileName = Replace(LogFileName, "/", "")
    LogFileName = Replace(LogFileName, " ", "")
    LogFileName = Replace(LogFileName, ":", "")
    
    ' wsLog.Name = LogSheetName
    
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    Dim tso As TextStream
    
    Set tso = FSO.CreateTextFile(ActiveWorkbook.Path & "\" & LogFileName & ".txt")
    
    tso.Write "hogehoge"
    
    
    
End Sub


