Attribute VB_Name = "mfMail"
'�Q�Ɛݒ� : Microsoft Outlook 16.0 Object Library

Sub mailsend()
    
    Dim objOutlook As Outlook.Application
    Dim objMail As Outlook.MailItem
    
    Set objOutlook = New Outlook.Application
    Set objMail = objOutlook.CreateItem(olMailItem)

    With objMail
        .To = "hi.t4ro@gmail.com"       '���[������
        .Subject = "macro mail test"    '���[������
        .Body = mfReplace(ActiveWorkbook.Worksheets("mailTemplate").Cells(1, 1).Value)      '���[���{��
        .BodyFormat = olFormatPlain     '���[���̌`��
        .Display                        '���������[���̕\��
    End With
    
    MsgBox "���[���쐬����"

End Sub
