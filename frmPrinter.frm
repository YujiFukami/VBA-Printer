VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrinter 
   Caption         =   "印刷機設定"
   ClientHeight    =   3648
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3468
   OleObjectBlob   =   "frmPrinter.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Cmd印刷機に設定_Click()
    
    Dim PrinterName$
    If Me.list印刷機.Value = Null Then
        Exit Sub
    End If
    
    On Error Resume Next
    PrinterName = Me.list印刷機.Value
    On Error GoTo 0
    
    If PrinterName <> "" Then
        Call 印刷機設定(PrinterName, False)
        Call 現在設定印刷機表示
    End If
    

End Sub

Private Sub Cmd閉じる_Click()
    
    Unload Me
    
End Sub

'20210708追加
Private Sub list印刷機_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim PrinterName$
    PrinterName = Me.list印刷機.Value
    Call 印刷機設定(PrinterName, False)
    Call 現在設定印刷機表示
    
End Sub

Private Sub tgl印刷機に設定_Click()
    
    If Me.tgl印刷機に設定.Value = True Then
        Me.tgl印刷機に設定.Value = False
        Exit Sub
    End If
    
    Dim PrinterName$
    If Me.list印刷機.Value = Null Then
        Exit Sub
    End If
    
    On Error Resume Next
    PrinterName = Me.list印刷機.Value
    On Error GoTo 0
    
    If PrinterName <> "" Then
        Call 印刷機設定(PrinterName, False)
        Call 現在設定印刷機表示
    End If
    
End Sub

Private Sub tgl閉じる_Click()
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim PrinterList
    PrinterList = 印刷機一覧取得
    Me.list印刷機.List = PrinterList
    
    Call 現在設定印刷機表示
    
End Sub
Sub 現在設定印刷機表示()

    Dim NowPrinerName$
    NowPrinerName = Application.ActivePrinter
    NowPrinerName = Split(NowPrinerName, " on Ne")(0)
    Me.txt設定中印刷機.Text = NowPrinerName

End Sub
