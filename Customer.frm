VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Customer 
   Caption         =   "Select Customer"
   ClientHeight    =   3315
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   2412
   OleObjectBlob   =   "Customer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Booking.ComboBox2.text = Customer.CommandButton1.Caption
    Customer.Hide
End Sub

Private Sub CommandButton2_Click()
    Booking.ComboBox2.text = Customer.CommandButton2.Caption
    Customer.Hide
End Sub

Private Sub CommandButton3_Click()
    Booking.ComboBox2.text = Customer.CommandButton3.Caption
    Customer.Hide
End Sub

Private Sub CommandButton4_Click()
    Booking.ComboBox2.text = Customer.CommandButton4.Caption
    Customer.Hide
End Sub

Private Sub CommandButton5_Click()
    Booking.ComboBox2.text = Customer.CommandButton5.Caption
    Customer.Hide
End Sub
