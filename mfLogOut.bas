Attribute VB_Name = "mfLogOut"
Dim wsLog As Worksheet



Function mfLogOutInitialize(wb As Workbook) As Boolean
    
    Set wsLog = wb.Worksheets.Add(after:=Sheets(wb.Worksheets.Count))
    
    Dim LogSheetName As String
    
    LogSheetName = "Log_" & Now
    LogSheetName = Replace(LogSheetName, "/", "")
    LogSheetName = Replace(LogSheetName, " ", "")
    LogSheetName = Replace(LogSheetName, ":", "")
    
    wsLog.Name = LogSheetName
    
    mfLogOutinitilalize = True
    
End Function


