VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJindu 
   ClientHeight    =   1395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6570
   OleObjectBlob   =   "frmJindu.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "frmJindu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = 1
End Sub

Private Sub UserForm_Initialize()
    Dim w, h

    With Me
        .StartUpPosition = 0  'λ����ʾ�����ֶ�
        .Caption = "������ʾ"
        w = Application.Width 'excel�Ŀ�
        h = Application.Height 'excel�ĸ߶�
        .Left = Application.Left + (w - .Width) \ 2 ' ���㴰����߾�
        .Top = Application.Top + (h - .Height) \ 2  ' ���㴰��ĸ߶�
        '���涼��Ϊ���ô�����˫�����������ʾ��ͬһ����Ļ��
        .text.Caption = "0%"
        .bar.Width = 0
     End With
End Sub

Sub Init()
    Me.Show 0
End Sub
Sub Quit()
    Unload Me
End Sub

Sub jindu(i, total, Optional xian_shi = "���")
    Dim baifen
    baifen = i / total
    Me.text.Caption = xian_shi & "��" & Int(baifen * 100) & "%"
    DoEvents
    Me.bar.Width = Me.dibu.Width * baifen
End Sub

