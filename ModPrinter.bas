Attribute VB_Name = "ModPrinter"
Option Explicit
Sub ����@�ݒ�t�H�[���N��()
'20210719

    frmPrinter.Show

End Sub
Function ����@�ꗗ�擾()
'20210719�ǉ�
    
    Dim myShell As Object
    Dim myItem As Object
    Set myShell = CreateObject("Shell.Application")
    
    Dim PrinterList
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    ReDim PrinterList(1 To 1)
    
    K = 0
    For Each myItem In myShell.Namespace(&H4).ITEMS
        K = K + 1
        ReDim Preserve PrinterList(1 To K)
        PrinterList(K) = myItem.Name
    Next
    
    ����@�ꗗ�擾 = PrinterList
    
End Function
Sub ����@�ݒ�(PrinterName$, Optional MessageIruNaraTrue = True)
'20210719�ǉ�
    
    Dim I% '�����グ�p(Integer�^)
    Dim SetteiName$
    Dim SetteiKanryoNaraTrue As Boolean
    SetteiKanryoNaraTrue = False
    
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
    
    If SetteiKanryoNaraTrue Then
        If MessageIruNaraTrue Then
            MsgBox (SetteiName & "������@�ɐݒ肵�܂���")
        End If
                
    Else
        MsgBox (PrinterName & "�͈���ݒ�ł��܂���")
    End If
    
End Sub
Function �ݒ�ς݃v�����^�[���擾()
'20210719

    �ݒ�ς݃v�����^�[���擾 = Application.ActivePrinter
    
End Function
