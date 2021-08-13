Attribute VB_Name = "Module1"
Option Explicit

Sub startmacro_Click()

    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    
    mfLogOutInitialize
    oreoreFW
    
    MsgBox "Finish"
    Exit Sub
    


End Sub
