Attribute VB_Name = "mfLogOut"
'�Q�Ɛݒ� : Microsoft scripting runtime
'�Q�Ɛݒ�̎����ݒ�Fhttp://kouten0430.hatenablog.com/entry/2017/10/21/152301

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
    
    mfWriteLog ("�}�N���N��")
    
End Sub


Sub mfWriteLog(msg As String)
    
    msg = Now & "," & msg & vbCrLf

    Dim tso As TextStream
    Set tso = FSO.OpenTextFile(ActiveWorkbook.Path & "\" & LogFileName, ForAppending)
    
    tso.Write msg
    
    tso.Close

End Sub
