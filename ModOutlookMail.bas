Attribute VB_Name = "ModOutlookMail"
Option Explicit

Sub TestSendOutlookMail()
    
    Dim ToAddress$, InputTitle$, Bunsyo
    Dim HyoCell As Range, HyoInsertIti%
    ToAddress = "AAAAA@gmail.com;BBBB@gmail.com"
    InputTitle = "AAAA"
    ReDim Bunsyo(1 To 6)
    Bunsyo(1) = "A"
    Bunsyo(2) = "B"
    Bunsyo(3) = "C"
    Bunsyo(4) = "D"
    Bunsyo(5) = "E"
    Bunsyo(6) = "F"
    
    Set HyoCell = Range("F4:I9") '←←←←←←←←←←←←←←←←←←←←←←←
    HyoInsertIti = 3
    
    Call SendOutlookMail(ToAddress, InputTitle, Bunsyo, , , , , , True)
'    Call SendOutlookMail(ToAddress, InputTitle, Bunsyo, HyoCell, HyoInsertIti, , , , True)

End Sub

Sub SendOutlookMail(ToAddress As String, _
                        InputTitle As String, Bunsyo, _
                        Optional HyoCell As Range, Optional HyoInsertIti%, _
                        Optional CCAddress As String = "", Optional BCCAddress As String = "", _
                        Optional AttachPathList = Empty, _
                        Optional SendenAruNaraTrue = True, _
                        Optional AutoSendingNaraTrue = False)
'Outlookメールを自動送信する
'使用には「Microsoft Outlook 16.0 Object Library」のライブラリを参照すること
'20210721
    
'ToAddress・・・宛先アドレス、複数なら「;」を間に入れること
'InputTitle・・・件名
'Bunsyo・・・メール本文
'HyoCell・・・挿入する表のセル範囲
'HyoInsertIti・・・メール本文に表を挿入する位置
'CCAddress・・・CCのアドレス
'BCCAddress・・・BCCのアドレス
'AttachPathList・・・添付ファイルのファイルパス　リストで入力すること
'SendenAruNaraTrue・・・宣伝文を文末で入力するか
'AutoSendingNaraTrue・・・自動送信する場合はTrue デフォルトはFalse

    
    Dim objOutlook As Outlook.Application
    Dim objMail As Outlook.MailItem

    Set objOutlook = New Outlook.Application
    Set objMail = objOutlook.CreateItem(olMailItem)

    Dim attachObj As Outlook.Attachments
    Set attachObj = objMail.Attachments
    
    Dim HyoSheet As Worksheet
    
    Dim strBunsyo$
    If HyoCell Is Nothing Then
        '表がない
        strBunsyo = 本文を連結する(Bunsyo)
    Else
        strBunsyo = 本文を表挿入用に修正(Bunsyo, HyoInsertIti)
        Set HyoSheet = HyoCell.Parent   '←←←←←←←←←←←←←←←←←←←←←←←
    End If

    Dim SendenBun As String
    SendenBun = "<<本メールはExcelのマクロ機能を用いて自動で送信されています。>>" '←←←←←←←←←←←←←←←←←←←←←←←

    If SendenAruNaraTrue Then
        strBunsyo = strBunsyo & vbLf & vbLf & SendenBun
    End If

    Dim Rs&, Re&, Cs&, Ce& '始端行,列番号および終端行,列番号(Long型)
    
    Dim r
    
    Dim TmpAttathPath
    
    '添付ファイル添付
    If IsEmpty(AttachPathList) = False Then
        
        If AttachPathList(1) <> "" Then
            For Each TmpAttathPath In AttachPathList
                attachObj.Add TmpAttathPath
            Next
        End If
    End If
    
    With objMail
        .To = ToAddress
        
        If CCAddress <> "" Then
            .CC = CCAddress
        End If
        
        If BCCAddress <> "" Then
            .BCC = BCCAddress
        End If
        
        .Subject = InputTitle
        
        .Display    ' メール作成画面で表示する
        
        .Body = strBunsyo
        .BodyFormat = 2 'HTML
     
        '表挿入
        If Not HyoCell Is Nothing Then
            HyoSheet.Activate
            HyoCell.Select
            Selection.Copy
            
            r = InStr(.Body, "【図】") - HyoInsertIti
            .Body = Replace(.Body, "【図】", "")
            
            On Error Resume Next
            objOutlook.ActiveInspector.WordEditor.Range(r, r).Paste
            On Error GoTo 0

        End If
                       
        If AutoSendingNaraTrue Then
            .Send   ' メール送信
        End If
    
    End With
    
End Sub

Function 本文を表挿入用に修正(Bunsyo, HyoInsertIti%)
'Outlookメール送信用のライブラリ
'配列に入った文章を改行を追加して連結して文字列にする
'表の挿入位置に"【図】"を追加する。
'20210721
    
    '空白位置にスペースを入れる（表の挿入位置がうまくいく）
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Output1, Output2$
    Output1 = Bunsyo
    For I = 1 To UBound(Bunsyo)
        If Output1(I) = "" Then
            Output1(I) = " "
        End If
    Next I
    
    Output2 = ""
    For I = 1 To UBound(Output1)
        If I = HyoInsertIti Then
            Output2 = Output2 & "【図】"
        End If
        
        Output2 = Output2 & Output1(I) & vbLf
    
    Next I
    
    '最終位置に表がつく場合
    If HyoInsertIti > UBound(Output1, 1) Then
        Output2 = Output2 & "【図】"
    End If
          
    本文を表挿入用に修正 = Output2
    
End Function

Function 本文を連結する(Bunsyo)
'Outlookメール送信用のライブラリ
'配列に入った文章を改行を追加して連結して文字列にする
'20210721

    Dim Output
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    N = UBound(Bunsyo, 1)
    Output = ""
    For I = 1 To N
        If I = 1 Then
            Output = Bunsyo(I)
        Else
            Output = Output & vbLf & Bunsyo(I)
        End If
    Next I
    
    本文を連結する = Output
    
End Function
