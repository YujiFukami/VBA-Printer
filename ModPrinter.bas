Attribute VB_Name = "ModPrinter"
Option Explicit
Sub 印刷機設定フォーム起動()
'20210719

    frmPrinter.Show

End Sub
Function 印刷機一覧取得()
'20210719追加
    
    Dim myShell As Object
    Dim myItem As Object
    Set myShell = CreateObject("Shell.Application")
    
    Dim PrinterList
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    ReDim PrinterList(1 To 1)
    
    K = 0
    For Each myItem In myShell.Namespace(&H4).ITEMS
        K = K + 1
        ReDim Preserve PrinterList(1 To K)
        PrinterList(K) = myItem.Name
    Next
    
    印刷機一覧取得 = PrinterList
    
End Function
Sub 印刷機設定(PrinterName$, Optional MessageIruNaraTrue = True)
'20210719追加
    
    Dim I% '数え上げ用(Integer型)
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
            MsgBox (SetteiName & "を印刷機に設定しました")
        End If
                
    Else
        MsgBox (PrinterName & "は印刷設定できません")
    End If
    
End Sub
Function 設定済みプリンター名取得()
'20210719

    設定済みプリンター名取得 = Application.ActivePrinter
    
End Function
