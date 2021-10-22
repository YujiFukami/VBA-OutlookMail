Attribute VB_Name = "ModOutlookMail"
Option Explicit

'TestSendOutlookMail �E�E�E���ꏊ�FFukamiAddins3.ModOutlookMail
'SendOutlookMail     �E�E�E���ꏊ�FFukamiAddins3.ModOutlookMail
'�{����\�}���p�ɏC���E�E�E���ꏊ�FFukamiAddins3.ModOutlookMail
'�{����A������      �E�E�E���ꏊ�FFukamiAddins3.ModOutlookMail



Sub TestSendOutlookMail()
    
    Dim ToAddress    As String
    Dim InputTitle   As String
    Dim Bunsyo
    Dim HyoCell      As Range
    Dim HyoInsertIti As Integer
    ToAddress = "AAAAA@gmail.com;BBBB@gmail.com"
    InputTitle = "AAAA"
    ReDim Bunsyo(1 To 6)
    Bunsyo(1) = "A"
    Bunsyo(2) = "B"
    Bunsyo(3) = "C"
    Bunsyo(4) = "D"
    Bunsyo(5) = "E"
    Bunsyo(6) = "F"
    
    Set HyoCell = Range("F4:I9") '����������������������������������������������
    HyoInsertIti = 3
    
    Call SendOutlookMail(ToAddress, InputTitle, Bunsyo, , , , , , True)

End Sub

Sub SendOutlookMail(ToAddress As String, _
                        InputTitle As String, Bunsyo, _
                        Optional HyoCell As Range, Optional HyoInsertIti As Integer, _
                        Optional CCAddress As String = "", Optional BCCAddress As String = "", _
                        Optional AttachPathList = Empty, _
                        Optional SendenAruNaraTrue = True, _
                        Optional AutoSendingNaraTrue = False)
'Outlook���[�����������M����
'�g�p�ɂ́uMicrosoft Outlook 16.0 Object Library�v�̃��C�u�������Q�Ƃ��邱��
'20210721
    
'ToAddress          �E�E�E����A�h���X�A�����Ȃ�u;�v���Ԃɓ���邱��
'InputTitle         �E�E�E����
'Bunsyo             �E�E�E���[���{��
'HyoCell            �E�E�E�}������\�̃Z���͈�
'HyoInsertIti       �E�E�E���[���{���ɕ\��}������ʒu
'CCAddress          �E�E�ECC�̃A�h���X
'BCCAddress         �E�E�EBCC�̃A�h���X
'AttachPathList     �E�E�E�Y�t�t�@�C���̃t�@�C���p�X�@���X�g�œ��͂��邱��
'SendenAruNaraTrue  �E�E�E��`���𕶖��œ��͂��邩
'AutoSendingNaraTrue�E�E�E�������M����ꍇ��True �f�t�H���g��False

    
    Dim objOutlook As Outlook.Application
    Dim objMail    As Outlook.MailItem
    Dim attachObj  As Outlook.Attachments
    Dim HyoSheet   As Worksheet
    Dim strBunsyo  As String
    Dim SendenBun  As String

    Set objOutlook = New Outlook.Application
    Set objMail = objOutlook.CreateItem(olMailItem)
    Set attachObj = objMail.Attachments
    
    If HyoCell Is Nothing Then
        '�\���Ȃ�
        strBunsyo = �{����A������(Bunsyo)
    Else
        strBunsyo = �{����\�}���p�ɏC��(Bunsyo, HyoInsertIti)
        Set HyoSheet = HyoCell.Parent   '����������������������������������������������
    End If

    SendenBun = "<<�{���[����Excel�̃}�N���@�\��p���Ď����ő��M����Ă��܂��B>>" '����������������������������������������������

    If SendenAruNaraTrue Then
        strBunsyo = strBunsyo & vbLf & vbLf & SendenBun
    End If

    Dim Rs           As Long
    Dim Re           As Long
    Dim Cs           As Long
    Dim Ce           As Long
    Dim R
    Dim TmpAttathPath
    
    '�Y�t�t�@�C���Y�t
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
        
        .Display    ' ���[���쐬��ʂŕ\������
        
        .Body = strBunsyo
        .BodyFormat = 2 'HTML
     
        '�\�}��
        If Not HyoCell Is Nothing Then
            HyoSheet.Activate
            HyoCell.Select
            Selection.Copy
            
            R = InStr(.Body, "�y�}�z") - HyoInsertIti
            .Body = Replace(.Body, "�y�}�z", "")
            
            On Error Resume Next
            objOutlook.ActiveInspector.WordEditor.Range(R, R).Paste
            On Error GoTo 0

        End If
                       
        If AutoSendingNaraTrue Then
            .Send   ' ���[�����M
        End If
    
    End With
    
End Sub

Function �{����\�}���p�ɏC��(Bunsyo, HyoInsertIti As Integer)
'Outlook���[�����M�p�̃��C�u����
'�z��ɓ��������͂����s��ǉ����ĘA�����ĕ�����ɂ���
'�\�̑}���ʒu��"�y�}�z"��ǉ�����B
'20210721
    
    '�󔒈ʒu�ɃX�y�[�X������i�\�̑}���ʒu�����܂������j
    Dim I       As Integer
    Dim J       As Integer
    Dim K       As Integer
    Dim M       As Integer
    Dim N       As Integer
    Dim Output1
    Dim Output2 As String
    Output1 = Bunsyo
    For I = 1 To UBound(Bunsyo)
        If Output1(I) = "" Then
            Output1(I) = " "
        End If
    Next I
    
    Output2 = ""
    For I = 1 To UBound(Output1)
        If I = HyoInsertIti Then
            Output2 = Output2 & "�y�}�z"
        End If
        
        Output2 = Output2 & Output1(I) & vbLf
    
    Next I
    
    '�ŏI�ʒu�ɕ\�����ꍇ
    If HyoInsertIti > UBound(Output1, 1) Then
        Output2 = Output2 & "�y�}�z"
    End If
          
    �{����\�}���p�ɏC�� = Output2
    
End Function

Function �{����A������(Bunsyo)
'Outlook���[�����M�p�̃��C�u����
'�z��ɓ��������͂����s��ǉ����ĘA�����ĕ�����ɂ���
'20210721

    Dim Output
    Dim I     As Integer
    Dim J     As Integer
    Dim K     As Integer
    Dim M     As Integer
    Dim N     As Integer
    N = UBound(Bunsyo, 1)
    Output = ""
    For I = 1 To N
        If I = 1 Then
            Output = Bunsyo(I)
        Else
            Output = Output & vbLf & Bunsyo(I)
        End If
    Next I
    
    �{����A������ = Output
    
End Function


