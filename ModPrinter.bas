Attribute VB_Name = "ModPrinter"
Option Explicit

Function GetSettingPrinter()
'20210719

    GetSettingPrinter = Application.ActivePrinter
    
End Function

Function GetPrinterList()
'設定可能なプリンター一覧取得
'20210719
    
    Dim myShell As Object
    Dim myItem As Object
    Set myShell = CreateObject("Shell.Application")
    
    Dim PrinterList
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
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
'プリンター名から印刷対象のプリンター設定
'20210719

'引数
'PrinterName         ・・・プリンター名（String型）
'[MessageIrunaraTrue]・・・確認メッセージがいるかどうか。デフォルトはTrue
                                                                         

    Dim I% '数え上げ用(Integer型)
    Dim SetteiName$
    Dim SetteiKanryoNaraTrue As Boolean
    SetteiKanryoNaraTrue = False
    
    '「プリンター名 on Ne**」の「**」の番号を1つずつ試してうまくいくやつを探索
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
    
    '確認メッセージ
    If SetteiKanryoNaraTrue Then
        '設定に成功した場合
        If MessageIrunaraTrue Then
            MsgBox (SetteiName & "を印刷機に設定しました")
        End If
                
    Else
        '設定に失敗した場合
        MsgBox (PrinterName & "は印刷設定できません")
    End If
    
End Sub
