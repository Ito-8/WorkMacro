Attribute VB_Name = "mfLogOut"
'参照設定 : Microsoft scripting runtime
'参照設定の自動設定：http://kouten0430.hatenablog.com/entry/2017/10/21/152301

Dim FSO As FileSystemObject
Dim LogFileName As String


Sub mfLogOutInitialize()
    
    LogFileName = "Log_" & Now
    LogFileName = Replace(LogFileName, "/", "")
    LogFileName = Replace(LogFileName, " ", "")
    LogFileName = Replace(LogFileName, ":", "")
    LogFileName = LogFileName & ".csv"

    Set FSO = New FileSystemObject
    
    Dim tso As TextStream
    Set tso = FSO.CreateTextFile(ActiveWorkbook.Path & "\" & LogFileName)

    tso.Close
    
    mfWriteLog ("マクロ起動")
    
End Sub


Sub mfWriteLog(msg As String)
    
    msg = Now & "," & msg & vbCrLf

    Dim tso As TextStream
    Set tso = FSO.OpenTextFile(ActiveWorkbook.Path & "\" & LogFileName, ForAppending)
    
    tso.Write msg
    
    tso.Close

End Sub
