Attribute VB_Name = "mfTools"
'�Q�Ɛݒ�FMicrosoft Shell Controls And Automation

'mfKoushinNitiji:�t�@�C���̍X�V������ύX����
'�����F
'   filepath:�Ώۃt�@�C���̃t���p�X(�t�H���_�p�X+�t�@�C����)
'   Nitiji:�ύX��̃t�@�C���X�V����(�`���̗�F2021/04/14 20:30:00)

Sub mfKoushinNitiji(filePath As String, Nitiji As String)
    
    Dim fileName As String
    Dim PathName As String
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    fileName = FSO.GetFileName(filePath)
    PathName = FSO.GetParentFolderName(filePath)
    
    Dim shell As Shell32.shell
    Dim tFolder As Shell32.folder
    Dim tfile As FolderItem
    
    Set shell = New Shell32.shell
    Set tFolder = shell.Namespace(PathName)
    Set tfile = tFolder.ParseName(fileName)
    tfile.ModifyDate = Nitiji

End Sub


Function mfReplace(replaceTarget As String) As String

    Dim wsReplaceMaster As Worksheet
    
    Set wsReplaceMaster = ActiveWorkbook.Worksheets("ReplaceMaster")
    
    Dim i As Integer
    i = 2
    
    Do While wsReplaceMaster.Cells(i, 1).Value <> ""
        replaceTarget = Replace(replaceTarget, wsReplaceMaster.Cells(i, 1), wsReplaceMaster.Cells(i, 2).Value)
        i = i + 1
    Loop
    
    mfReplace = replaceTarget
    
End Function
