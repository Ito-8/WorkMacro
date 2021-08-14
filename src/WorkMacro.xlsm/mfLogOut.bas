Attribute VB_Name = "mfLogOut"
'�Q�Ɛݒ� : Microsoft scripting runtime
'�Q�Ɛݒ�̎����ݒ�Fhttp://kouten0430.hatenablog.com/entry/2017/10/21/152301

Option Explicit

Dim LogFilePath As String


Sub mfLogOutInitialize()
    
    LogFilePath = "Log_" & Now
    LogFilePath = Replace(LogFilePath, "/", "")
    LogFilePath = Replace(LogFilePath, " ", "")
    LogFilePath = Replace(LogFilePath, ":", "")
    LogFilePath = LogFilePath & ".csv"
    LogFilePath = ActiveWorkbook.Path & "\Log\" & LogFilePath

    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    Dim tso As TextStream
    Set tso = FSO.CreateTextFile(LogFilePath)

    tso.Close
    
    mfWriteLog ("�}�N���N��")
    
End Sub


Sub mfWriteLog(msg As String)
    
    msg = Now & "," & msg & vbCrLf
    
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    Dim tso As TextStream
    Set tso = FSO.OpenTextFile(LogFilePath, ForAppending)
    
    tso.Write msg
    
    tso.Close

End Sub
