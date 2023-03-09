VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Shifts 
   Caption         =   "Select Shift"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   2436
   OleObjectBlob   =   "Shifts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Shifts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Booking.ComboBox1.text = Shifts.CommandButton1.Caption
    Shifts.Hide
End Sub

Private Sub CommandButton2_Click()
    Booking.ComboBox1.text = Shifts.CommandButton2.Caption
    Shifts.Hide
End Sub

Private Sub CommandButton3_Click()
    Booking.ComboBox1.text = Shifts.CommandButton3.Caption
    Shifts.Hide
End Sub

Private Sub CommandButton4_Click()
    Booking.ComboBox1.text = Shifts.CommandButton4.Caption
    Shifts.Hide
End Sub

Private Sub CommandButton5_Click()
    Booking.ComboBox1.text = Shifts.CommandButton5.Caption
    Shifts.Hide
End Sub

Private Sub CommandButton6_Click()
    Booking.ComboBox1.text = Shifts.CommandButton6.Caption
    Shifts.Hide
End Sub
