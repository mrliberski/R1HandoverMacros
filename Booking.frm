VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Booking 
   Caption         =   "Empties Booking Form"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4500
   OleObjectBlob   =   "Booking.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Booking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ComboBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Shifts.Show
End Sub
Private Sub ComboBox2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Customer.Show
End Sub
Private Sub ComboBox1_selectionChange()
    MsgBox "fff"
End Sub



Private Sub CommandButton5_Click()
'reset
            Booking.TextBox1.value = ""            'clearing up old data remains from the form
            Booking.TextBox2.value = ""            'from the booking form
            Booking.ComboBox1.value = ""
            Booking.ComboBox2.value = ""
            Booking.TextBox3.value = ""
            Booking.TextBox4.value = ""
            Booking.TextBox5.value = ""
            Booking.TextBox6.value = ""
            Booking.TextBox7.value = ""
            Booking.TextBox8.value = ""
            
            Booking.TextBox1.Enabled = True
            Booking.TextBox2.Enabled = True
            Booking.ComboBox1.Enabled = True
            Booking.ComboBox2.Enabled = True
            Booking.TextBox3.Enabled = True
            Booking.TextBox4.Enabled = True

            TextBox1.SetFocus
    
             Booking.Repaint


End Sub

Private Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox1.text = dateVariable
End Sub

Private Sub CommandButton1_Click()

'=============================== CHECK IF DATABASE IS ACCESSIBLE
Dim FPath As String
FPath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
If Dir(FPath) = "" Then
    MsgBox "Database is not accessible. Please try again later.", vbOKOnly, "Could not find database."
    Exit Sub
End If
'======================================================================

'ADD
    On Error Resume Next ' in case of type mismatch
    
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim MyConn
        MyConn = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb" 'adres do bazy
    Dim enviro
        enviro = Environ("Username")
    Dim Datum
        Datum = Date

    Dim j As Long           'element number in data stream
    Dim k As Integer        'column counter
    Dim sSQL As String      'SQL command string
    
    
            
    sSQL = "SELECT * FROM Packaging_Log"     'sql query - select * from selected database
    
    Set cnn = New ADODB.Connection
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn
    End With
    
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseServer
    rst.Open source:=sSQL, ActiveConnection:=cnn, _
        CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, _
        Options:=adCmdText
            
    If rst.Supports(adAddNew) Then
        rst.AddNew
                
                rst(1).value = Datum
                rst(2).value = enviro
                rst(3).value = "RED1"
                rst(4).value = DateValue(TextBox1.value)             'now we are adding data to it
                rst(5).value = TextBox2.value             'from the booking form
                rst(6).value = UCase(ComboBox1.value)
                rst(7).value = UCase(ComboBox2.value)
                rst(8).value = UCase(TextBox3.value)
                rst(9).value = UCase(TextBox4.value)
                rst(10).value = UCase(TextBox5.value)
                rst(11).value = TextBox6.value
                rst(12).value = UCase(TextBox7.value)
                rst(13).value = UCase(TextBox8.value)

                rst.Update ' update database
                
                rst.Close           'close it up
                cnn.Close           'and clear up before next throw
                Set rst = Nothing   'and so on
                Set cnn = Nothing   'and on
    Else
                rst.Close           'close it up
                cnn.Close           'and clear up before next throw
                Set rst = Nothing   'and so on
                Set cnn = Nothing   'and on
    End If
    '==============================================
    
            Booking.TextBox1.value = ""            'clearing up old data remains from the form
            Booking.TextBox2.value = ""            'from the booking form
            Booking.ComboBox1.value = ""
            Booking.ComboBox2.value = ""
            Booking.TextBox3.value = ""
            Booking.TextBox4.value = ""
            Booking.TextBox5.value = ""
            Booking.TextBox6.value = ""
            Booking.TextBox7.value = ""
            Booking.TextBox8.value = ""
            
            Booking.TextBox1.Enabled = True
            Booking.TextBox2.Enabled = True
            Booking.ComboBox1.Enabled = True
            Booking.ComboBox2.Enabled = True
            Booking.TextBox3.Enabled = True
            Booking.TextBox4.Enabled = True

            TextBox1.SetFocus
    
    Booking.Repaint

'MsgBox "Entry Added to db", vbInformation, "OK"


End Sub

Private Sub CommandButton2_Click()
'Cancel button
'must reset form please
    Booking.Hide

End Sub

Private Sub CommandButton3_Click()
'=============================== CHECK IF DATABASE IS ACCESSIBLE
Dim FPath As String
FPath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
If Dir(FPath) = "" Then
    MsgBox "Database is not accessible. Please try again later.", vbOKOnly, "Could not find database."
    Exit Sub
End If
'======================================================================

'Add and RESET FORM FOR THE NEXT DELIVERY
'ADD
    On Error Resume Next ' in case of type mismatch
    
  
    
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim MyConn
        MyConn = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb" 'adres do bazy
    Dim enviro
        enviro = Environ("Username")
    Dim Datum
        Datum = Date
    
    '==============================================

    Dim j As Long           'element number in data stream
    Dim k As Integer        'column counter
    Dim sSQL As String      'SQL command string
            
    sSQL = "SELECT * FROM Packaging_Log"     'sql query - select * from selected database
    
    Set cnn = New ADODB.Connection
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn
    End With
    
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseServer
    rst.Open source:=sSQL, ActiveConnection:=cnn, _
        CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, _
        Options:=adCmdText
            
    If rst.Supports(adAddNew) Then
        rst.AddNew
                
                rst(1).value = Datum
                rst(2).value = enviro
                rst(3).value = "RED1"
                rst(4).value = DateValue(TextBox1.value)             'now we are adding data to it
                rst(5).value = TextBox2.value             'from the booking form
                rst(6).value = UCase(ComboBox1.value)
                rst(7).value = UCase(ComboBox2.value)
                rst(8).value = UCase(TextBox3.value)
                rst(9).value = UCase(TextBox4.value)
                rst(10).value = UCase(TextBox5.value)
                rst(11).value = UCase(TextBox6.value)
                rst(12).value = UCase(TextBox7.value)
                rst(13).value = UCase(TextBox8.value)


                rst.Update ' update database
                
                rst.Close           'close it up
                cnn.Close           'and clear up before next throw
                Set rst = Nothing   'and so on
                Set cnn = Nothing   'and on
    Else
                rst.Close           'close it up
                cnn.Close           'and clear up before next throw
                Set rst = Nothing   'and so on
                Set cnn = Nothing   'and on
    End If
    '==============================================

    Booking.Hide
    
    Booking.TextBox1.Enabled = True
    Booking.TextBox2.Enabled = True
    Booking.ComboBox1.Enabled = True
    Booking.ComboBox2.Enabled = True
    Booking.TextBox3.Enabled = True
    Booking.TextBox4.Enabled = True
    
    TextBox2.value = ""
    ComboBox2.value = ""
    TextBox3.value = ""
    TextBox4.value = ""
    TextBox5.value = ""         'resetting part of form we can clear
    TextBox6.value = ""
    TextBox7.value = ""
    TextBox8.value = ""
    TextBox2.SetFocus
    
    Booking.Show

End Sub


Private Sub CommandButton4_Click()

'=============================== CHECK IF DATABASE IS ACCESSIBLE
Dim FPath As String
FPath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
If Dir(FPath) = "" Then
    MsgBox "Database is not accessible. Please try again later.", vbOKOnly, "Could not find database."
    Exit Sub
End If
'======================================================================

    'NEXT ITEM BUTTON
    On Error Resume Next
    

    
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim MyConn
        MyConn = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb" 'adres do bazy
    Dim enviro
        enviro = Environ("Username")
    Dim Datum
        Datum = Date
        
    
    Dim NextRow                                                'need to find out where the end of the table is
    NextRow = Sheet2.ListObjects("Table6").ListRows.Count + 7  'we count the size of table and would add 7 to allow offset
    Sheet2.ListObjects("Table6").ListRows.Add                  'we'll addan extra row to add data to it
    
    Sheet2.Cells(NextRow, 1).value = DateValue(TextBox1.value)             'now we are adding data to it
    Sheet2.Cells(NextRow, 2).value = TextBox2.value             'from the booking form
    Sheet2.Cells(NextRow, 3).value = UCase(ComboBox1.value)
    Sheet2.Cells(NextRow, 4).value = UCase(ComboBox2.value)
    Sheet2.Cells(NextRow, 5).value = UCase(TextBox3.value)
    Sheet2.Cells(NextRow, 6).value = UCase(TextBox4.value)
    Sheet2.Cells(NextRow, 7).value = UCase(TextBox5.value)
    Sheet2.Cells(NextRow, 8).value = UCase(TextBox6.value)
    Sheet2.Cells(NextRow, 9).value = UCase(TextBox7.value)
    Sheet2.Cells(NextRow, 10).value = UCase(TextBox8.value)
    
    
    '==============================================

    Dim j As Long           'element number in data stream
    Dim k As Integer        'column counter
    Dim sSQL As String      'SQL command string
            
    sSQL = "SELECT * FROM Packaging_Log"     'sql query - select * from selected database
    
    Set cnn = New ADODB.Connection
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn
    End With
    
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseServer
    rst.Open source:=sSQL, ActiveConnection:=cnn, _
        CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, _
        Options:=adCmdText
            
    If rst.Supports(adAddNew) Then
        rst.AddNew
                
                rst(1).value = Datum
                rst(2).value = enviro
                rst(3).value = "RED1"
                rst(4).value = DateValue(TextBox1.value)             'now we are adding data to it
                rst(5).value = TextBox2.value             'from the booking form
                rst(6).value = UCase(ComboBox1.value)
                rst(7).value = UCase(ComboBox2.value)
                rst(8).value = UCase(TextBox3.value)
                rst(9).value = UCase(TextBox4.value)
                rst(10).value = UCase(TextBox5.value)
                rst(11).value = UCase(TextBox6.value)
                rst(12).value = UCase(TextBox7.value)
                rst(13).value = UCase(TextBox8.value)
                

                rst.Update ' update database
                
                rst.Close           'close it up
                cnn.Close           'and clear up before next throw
                Set rst = Nothing   'and so on
                Set cnn = Nothing   'and on
    Else
                rst.Close           'close it up
                cnn.Close           'and clear up before next throw
                Set rst = Nothing   'and so on
                Set cnn = Nothing   'and on
    End If
    '==============================================

    Booking.Hide                'we have to hide it here in order to refresh it
    
    TextBox1.Enabled = False
    TextBox2.Enabled = False
    ComboBox1.Enabled = False
    ComboBox2.Enabled = False
    TextBox3.Enabled = False
    TextBox4.Enabled = False
    
    TextBox5.value = ""         'resetting part of form we can clear
    TextBox6.value = ""
    TextBox7.value = ""
    TextBox8.value = ""
    TextBox5.SetFocus
    
    Booking.Show
 
    
End Sub

Private Sub UserForm_Initialize()

    ComboBox1.AddItem "RED"
    ComboBox1.AddItem "YELLOW"
    ComboBox1.AddItem "BLUE"
    ComboBox1.AddItem "GREEN"
    ComboBox1.AddItem "ORANGE"
    ComboBox2.AddItem "OXFORD"
    ComboBox2.AddItem "NED"
    ComboBox2.AddItem "HUYTON"
    
    'ComboBox2.AddItem "DROITWICH"
    
End Sub

