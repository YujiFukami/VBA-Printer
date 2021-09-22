Attribute VB_Name = "ModPrinter"
Option Explicit

Function GetSettingPrinter()
'20210719

    GetSettingPrinter = Application.ActivePrinter
    
End Function

Function GetPrinterList()
'�ݒ�\�ȃv�����^�[�ꗗ�擾
'20210719
    
    Dim myShell As Object
    Dim myItem As Object
    Set myShell = CreateObject("Shell.Application")
    
    Dim PrinterList
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    ReDim PrinterList(1 To 1)
    
    K = 0
    For Each myItem In myShell.Namespace(&H4).Items
        K = K + 1
        ReDim Preserve PrinterList(1 To K)
        PrinterList(K) = myItem.Name
    Next
    
    GetPrinterList = PrinterList
    
End Function

Sub SetPrinter(PrinterName$, Optional MessageIrunaraTrue = True)
'�v�����^�[���������Ώۂ̃v�����^�[�ݒ�
'20210719

'����
'PrinterName         �E�E�E�v�����^�[���iString�^�j
'[MessageIrunaraTrue]�E�E�E�m�F���b�Z�[�W�����邩�ǂ����B�f�t�H���g��True
                                                                         

    Dim I% '�����グ�p(Integer�^)
    Dim SetteiName$
    Dim SetteiKanryoNaraTrue As Boolean
    SetteiKanryoNaraTrue = False
    
    '�u�v�����^�[�� on Ne**�v�́u**�v�̔ԍ���1�������Ă��܂��������T��
    On Error Resume Next
    For I = 1 To 99
        SetteiName = PrinterName & " on Ne" & Format(I, "00:")
        
        Application.ActivePrinter = SetteiName
        If Application.ActivePrinter = SetteiName Then
            SetteiKanryoNaraTrue = True
            Exit For
        End If

    Next I
    On Error GoTo 0
    
    '�m�F���b�Z�[�W
    If SetteiKanryoNaraTrue Then
        '�ݒ�ɐ��������ꍇ
        If MessageIrunaraTrue Then
            MsgBox (SetteiName & "������@�ɐݒ肵�܂���")
        End If
                
    Else
        '�ݒ�Ɏ��s�����ꍇ
        MsgBox (PrinterName & "�͈���ݒ�ł��܂���")
    End If
    
End Sub
