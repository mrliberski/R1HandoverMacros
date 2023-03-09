VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} deliveryreq 
   Caption         =   "Delivery Check In Request"
   ClientHeight    =   2685
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5568
   OleObjectBlob   =   "deliveryreq.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "deliveryreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    deliveryreq.Hide
End Sub

Private Sub CommandButton2_Click()
    Dim delivery_name As String
    delivery_name = deliveryreq.TextBox1.value
    deliveryreq.Hide
    delivery_check (delivery_name)
End Sub

