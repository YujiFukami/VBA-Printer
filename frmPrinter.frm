VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrinter 
   Caption         =   "����@�ݒ�"
   ClientHeight    =   3648
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3468
   OleObjectBlob   =   "frmPrinter.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Cmd����@�ɐݒ�_Click()
    
    Dim PrinterName$
    If Me.list����@.Value = Null Then
        Exit Sub
    End If
    
    On Error Resume Next
    PrinterName = Me.list����@.Value
    On Error GoTo 0
    
    If PrinterName <> "" Then
        Call ����@�ݒ�(PrinterName, False)
        Call ���ݐݒ����@�\��
    End If
    

End Sub

Private Sub Cmd����_Click()
    
    Unload Me
    
End Sub

'20210708�ǉ�
Private Sub list����@_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim PrinterName$
    PrinterName = Me.list����@.Value
    Call ����@�ݒ�(PrinterName, False)
    Call ���ݐݒ����@�\��
    
End Sub

Private Sub tgl����@�ɐݒ�_Click()
    
    If Me.tgl����@�ɐݒ�.Value = True Then
        Me.tgl����@�ɐݒ�.Value = False
        Exit Sub
    End If
    
    Dim PrinterName$
    If Me.list����@.Value = Null Then
        Exit Sub
    End If
    
    On Error Resume Next
    PrinterName = Me.list����@.Value
    On Error GoTo 0
    
    If PrinterName <> "" Then
        Call ����@�ݒ�(PrinterName, False)
        Call ���ݐݒ����@�\��
    End If
    
End Sub

Private Sub tgl����_Click()
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim PrinterList
    PrinterList = ����@�ꗗ�擾
    Me.list����@.List = PrinterList
    
    Call ���ݐݒ����@�\��
    
End Sub
Sub ���ݐݒ����@�\��()

    Dim NowPrinerName$
    NowPrinerName = Application.ActivePrinter
    NowPrinerName = Split(NowPrinerName, " on Ne")(0)
    Me.txt�ݒ蒆����@.Text = NowPrinerName

End Sub
