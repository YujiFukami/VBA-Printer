VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrinter 
   Caption         =   "σό@έθ"
   ClientHeight    =   3648
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3468
   OleObjectBlob   =   "frmPrinter.frx":0000
   StartUpPosition =   1  'I[i[ tH[Μ
End
Attribute VB_Name = "frmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Cmdσό@Ιέθ_Click()
    
    Dim PrinterName$
    If Me.listσό@.Value = Null Then
        Exit Sub
    End If
    
    On Error Resume Next
    PrinterName = Me.listσό@.Value
    On Error GoTo 0
    
    If PrinterName <> "" Then
        Call σό@έθ(PrinterName, False)
        Call »έέθσό@\¦
    End If
    

End Sub

Private Sub CmdΒΆι_Click()
    
    Unload Me
    
End Sub

'20210708ΗΑ
Private Sub listσό@_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim PrinterName$
    PrinterName = Me.listσό@.Value
    Call σό@έθ(PrinterName, False)
    Call »έέθσό@\¦
    
End Sub

Private Sub tglσό@Ιέθ_Click()
    
    If Me.tglσό@Ιέθ.Value = True Then
        Me.tglσό@Ιέθ.Value = False
        Exit Sub
    End If
    
    Dim PrinterName$
    If Me.listσό@.Value = Null Then
        Exit Sub
    End If
    
    On Error Resume Next
    PrinterName = Me.listσό@.Value
    On Error GoTo 0
    
    If PrinterName <> "" Then
        Call σό@έθ(PrinterName, False)
        Call »έέθσό@\¦
    End If
    
End Sub

Private Sub tglΒΆι_Click()
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim PrinterList
    PrinterList = σό@κζΎ
    Me.listσό@.List = PrinterList
    
    Call »έέθσό@\¦
    
End Sub

Sub »έέθσό@\¦()

    Dim NowPrinerName$
    NowPrinerName = Application.ActivePrinter
    NowPrinerName = Split(NowPrinerName, " on Ne")(0)
    Me.txtέθσό@.Text = NowPrinerName

End Sub
