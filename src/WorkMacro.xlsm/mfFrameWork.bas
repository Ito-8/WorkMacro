Attribute VB_Name = "mfFrameWork"
'éQè∆ê›íË:Microsoft scripting runtime

Option Explicit

Sub oreoreFW()

    Dim targetFile As File
    Dim targetBook As Workbook
    
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    For Each targetFile In FSO.GetFolder(ActiveWorkbook.Path & "\input\").Files
        If targetFile.Name Like "*.xlsx" Then
            targetFile.Copy Destination:=ActiveWorkbook.Path & "\output\", overwritefiles:=True
            Set targetFile = FSO.GetFile(ActiveWorkbook.Path & "\output\" & targetFile.Name)
            mfKoushinNitiji filePath:=targetFile.Path, Nitiji:="2021/04/14 20:30:00"
            
'            Set targetBook = Workbooks.Open(targetFile.Path, 3, False, , , , False)
'
'            Dim targetSheet As Worksheet
'            For Each targetSheet In targetBook.Worksheets
'                mfWriteLog msg:=targetBook.Name & "/" & targetSheet.Name
'                targetSheet.Cells(1, 1) = Now
'
'            Next targetSheet
'
'            targetBook.Close savechanges:=True

            End If
    Next targetFile

End Sub



