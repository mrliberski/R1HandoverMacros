VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} L551_Report 
   Caption         =   "L551 Despatch Report"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4728
   OleObjectBlob   =   "L551_Report.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "L551_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

ans = MsgBox("Report will be sent immediately. Do you wish to continue?", vbOKCancel, "Confirmation required")
If ans = vbCancel Then Exit Sub

On Error GoTo 0

    Dim arrivalTime As String
    Dim leaveTime As String
    Dim lhFronts As Integer
    Dim rhFronts As Integer
    Dim lhRears As Integer
    Dim rhRears As Integer
    Dim empties As Integer
    
    arrivalTime = Format(TextBox1.text, "hh:mm")
    leaveTime = Format(TextBox2.text, "hh:mm")
    
    If TextBox3 = NullString And TextBox4 = NullString And TextBox5 = NullString And TextBox6 = NullString Then
        answer = MsgBox("No quantities were typed in, Are you sure you want to send report?", _
            vbOKCancel, "No input")
            
            If answer = vbCancel Then
                Debug.Print "exiting, nothing was typed in"
                Exit Sub
            End If
    End If
    
    
    If TextBox3 = NullString Then
        lhFronts = 0
    Else
        lhFronts = Int(TextBox3.text)
    End If
    
    If TextBox4 = NullString Then
        rhFronts = 0
    Else
        rhFronts = Int(TextBox4.text)
    End If
    
    If TextBox5 = NullString Then
        lhRears = 0
    Else
        lhRears = Int(TextBox5.text)
    End If

    If TextBox6 = NullString Then
        rhRears = 0
    Else
        rhRears = Int(TextBox6.text)
    End If
    
    If TextBox7 = NullString Then
        empties = 0
    Else
        empties = Int(TextBox7.text)
    End If
    
    Debug.Print arrivalTime
    Debug.Print leaveTime
    Debug.Print lhFronts
    Debug.Print rhFronts
    Debug.Print lhRears
    Debug.Print rhRears
    Debug.Print empties
  
    Call L551_despatchReport(arrivalTime, leaveTime, lhFronts, rhFronts, lhRears, rhRears, empties)
    
    
    
    Unload Me

End Sub

Private Sub CommandButton2_Click()
    Unload Me
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

