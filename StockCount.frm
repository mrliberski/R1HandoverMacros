VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StockCount 
   Caption         =   "Stock Count"
   ClientHeight    =   4440
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4992
   OleObjectBlob   =   "StockCount.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StockCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton3_Click()

    Dim ctl ' Clear all textboxes
        For Each ctl In Me.Controls
            If TypeOf ctl Is MSForms.TextBox Then
                ctl.text = vbNullString
            End If
        Next ctl
        
        StockCount.Repaint
    
    
End Sub

Private Sub TextBox1_Change()
    OnlyNumbers
End Sub
Private Sub TextBox2_Change()
    OnlyNumbers
End Sub

Private Sub TextBox25_Change()
    OnlyNumbers
End Sub

Private Sub TextBox26_Change()
    OnlyNumbers
End Sub

Private Sub TextBox27_Change()
    OnlyNumbers
End Sub

Private Sub TextBox3_Change()
    OnlyNumbers
End Sub
Private Sub TextBox4_Change()
    OnlyNumbers
End Sub
 Private Sub TextBox5_Change()
    OnlyNumbers
End Sub
Private Sub TextBox6_Change()
    OnlyNumbers
End Sub
Private Sub TextBox7_Change()
    OnlyNumbers
End Sub
Private Sub TextBox8_Change()
    OnlyNumbers
End Sub
Private Sub TextBox9_Change()
    OnlyNumbers
End Sub
Private Sub TextBox10_Change()
    OnlyNumbers
End Sub
Private Sub TextBox11_Change()
    OnlyNumbers
End Sub
Private Sub TextBox12_Change()
    OnlyNumbers
End Sub
Private Sub TextBox13_Change()
    OnlyNumbers
End Sub
Private Sub TextBox14_Change()
    OnlyNumbers
End Sub
Private Sub TextBox15_Change()
    OnlyNumbers
End Sub
Private Sub TextBox16_Change()
    OnlyNumbers
End Sub
Private Sub TextBox17_Change()
    OnlyNumbers
End Sub
Private Sub TextBox18_Change()
    OnlyNumbers
End Sub
Private Sub TextBox19_Change()
    OnlyNumbers
End Sub
Private Sub TextBox20_Change()
    OnlyNumbers
End Sub
Private Sub TextBox21_Change()
    OnlyNumbers
End Sub
Private Sub TextBox22_Change()
    OnlyNumbers
End Sub
Private Sub TextBox23_Change()
    OnlyNumbers
End Sub
Private Sub TextBox24_Change()
    OnlyNumbers
End Sub

Private Sub OnlyNumbers()
    If TypeName(Me.ActiveControl) = "TextBox" Then
        With Me.ActiveControl
            If Not IsNumeric(.value) And .value <> vbNullString Then
                MsgBox "Sorry, only numbers allowed"
                .value = vbNullString
            End If
        End With
    End If
End Sub
Private Sub CommandButton1_Click()
    StockCount.Hide
End Sub

Private Sub CommandButton2_Click()

On Error Resume Next

    Dim xlh56 As Integer
    Dim xrh56 As Integer
    Dim xwad56 As Integer
    Dim xlh54 As Integer
    Dim xrh54 As Integer
    Dim xwad54 As Integer
    Dim xlhkab As Integer
    Dim xrhkab As Integer
    Dim f56demist As Integer
    Dim f54demist As Integer
    Dim ftray As Integer
    Dim evo As Integer
    
    xlh56 = StockCount.TextBox1.value
    xrh56 = StockCount.TextBox2.value
    xwad56 = StockCount.TextBox3.value
    xlh54 = StockCount.TextBox4.value
    xrh54 = StockCount.TextBox5.value
    xwad54 = StockCount.TextBox6.value
    
    xlhkab = StockCount.TextBox12.value
    xrhkab = StockCount.TextBox13.value
    
    f56demist = StockCount.TextBox26.value
    f54demist = StockCount.TextBox27.value

    ftray = StockCount.TextBox24.value
    evo = StockCount.TextBox25.value
    
    StockCount.Hide
    
    Dim ctl ' Clear all textboxes
        For Each ctl In Me.Controls
            If TypeOf ctl Is MSForms.TextBox Then
                ctl.text = vbNullString
            End If
        Next ctl
    
    
    Call stock_count_mail(xlh56, xrh56, xwad56, xlh54, xrh54, xwad54, xlhkab, xrhkab, _
    ftray, evo, f56demist, f54demist)
    
End Sub

