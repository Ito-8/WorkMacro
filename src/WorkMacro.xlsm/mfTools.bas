Attribute VB_Name = "mfTools"
'参照設定：Microsoft Shell Controls And Automation

'mfKoushinNitiji:ファイルの更新日時を変更する
'引数：
'   filepath:対象ファイルのフルパス(フォルダパス+ファイル名)
'   Nitiji:変更後のファイル更新日時(形式の例：2021/04/14 20:30:00)

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


