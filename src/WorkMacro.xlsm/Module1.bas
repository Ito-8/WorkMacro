Attribute VB_Name = "Module1"
Sub startmacro_Click()

    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    
    mfLogOutInitialize
    mfWriteLog ("hogehoge1")
    mfWriteLog ("hogehoge2")
    mfWriteLog ("hogehoge3")
    
    MsgBox "Finish"
    Exit Sub
    


End Sub
