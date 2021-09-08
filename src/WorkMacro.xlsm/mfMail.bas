Attribute VB_Name = "mfMail"
'参照設定 : Microsoft Outlook 16.0 Object Library

Sub mailsend()
    
    Dim objOutlook As Outlook.Application
    Dim objMail As Outlook.MailItem
    
    Set objOutlook = New Outlook.Application
    Set objMail = objOutlook.CreateItem(olMailItem)

    With objMail
        .To = "hi.t4ro@gmail.com"       'メール宛先
        .Subject = "macro mail test"    'メール件名
        .Body = mfReplace(ActiveWorkbook.Worksheets("mailTemplate").Cells(1, 1).Value)      'メール本文
        .BodyFormat = olFormatPlain     'メールの形式
        .Display                        '下書きメールの表示
    End With
    
    MsgBox "メール作成完了"

End Sub
