Attribute VB_Name = "Module1"
Sub startmacro_Click()

    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    
    If mfLogOutInitialize(wb) Then
        GoTo Err:
    End If
    
    MsgBox "Finish"
    Exit Sub
    
Err:
    MsgBox "error"


End Sub
