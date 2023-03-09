VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AssAdd 
   Caption         =   "Add New Record"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7776
   OleObjectBlob   =   "AssAdd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AssAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public rownums
Public Function resetme(rownums)
    
    If rownums < 1 Then rownums = 1

    AssesmentAdd.Label22.Caption = Sheet20.Cells(rownums, 1).value
    AssesmentAdd.TextBox1.text = Sheet20.Cells(rownums, 2).value
    AssesmentAdd.TextBox2.text = Format(Sheet20.Cells(rownums, 3).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox3.text = Format(Sheet20.Cells(rownums, 4).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox4.text = Format(Sheet20.Cells(rownums, 5).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox5.text = Format(Sheet20.Cells(rownums, 6).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox6.text = Format(Sheet20.Cells(rownums, 7).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox7.text = Format(Sheet20.Cells(rownums, 8).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox8.text = Format(Sheet20.Cells(rownums, 9).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox9.text = Format(Sheet20.Cells(rownums, 10).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox10.text = Format(Sheet20.Cells(rownums, 11).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox11.text = Format(Sheet20.Cells(rownums, 12).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox12.text = Format(Sheet20.Cells(rownums, 13).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox13.text = Format(Sheet20.Cells(rownums, 14).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox14.text = Format(Sheet20.Cells(rownums, 15).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox15.text = Format(Sheet20.Cells(rownums, 16).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox16.text = Sheet20.Cells(rownums, 17).value
    ComboBox1.text = Sheet20.Cells(rownums, 18).value
    ComboBox2.text = Sheet20.Cells(rownums, 19).value
    
    
End Function
Private Sub isAvailable(ByVal adres As String)
    If Dir(adres) <> NullString Then
        AssAdd.Label20.Caption = "Database availability Status: Available"
    Else
        AssAdd.Label20.Caption = "Database availability Status: Unavailable"
    End If
End Sub
Private Sub refreshFeed()

    '============================================== 'firstly let's check if we can access a database
    If Dir("J:\Pub-LOGISTICS\Packaging\Packaging.accdb") = NullString Then
        MsgBox "Could not connect to database. Try again later.", vbCritical, "Resource off limits."
        Exit Sub 'file dosn't exist or is not accessible
    End If
    '=============================================
    '============================================== ' clear spreadheet
    'CLEAR CONTENT
    Dim lastRowNum As Long
    lastRowNum = Sheet20.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Sheet20.Select
    Sheet20.Range("A1:S" & lastRowNum).ClearContents
    'Selection.ClearContents
    '==============================================
    '============================================== ' SQL QUERY BUILD
    Dim SqlString As String
        SqlString = "SELECT * FROM Assesments "
        
    'If CheckBox2.value = True Then
        'SqlString = SqlString & "WHERE ([ReceiveQty] < [AdvisedQty] OR [ReceiveQty] > [AdvisedQty]) "
    'End If

    SqlString = SqlString & "ORDER BY ID"
    '==============================================

    Call connectExecute(SqlString)
    'Call paintList
    'AssesmentAdd.Repaint




End Sub
Private Sub connectExecute(ByVal sqlRequest As String)

    '============================================== ' define connections
    Dim cnn As ADODB.Connection
        Set cnn = New ADODB.Connection
    Dim rst As ADODB.Recordset
        Set rst = New ADODB.Recordset
    Dim fld As ADODB.Field
    Dim MyConn
        MyConn = dataBaseAddress
        MyConn = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"  'adres do bazy
    '==============================================
    
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn                                      'open connection
    End With
    
    rst.CursorLocation = adUseServer
    rst.Open source:=sqlRequest, ActiveConnection:=cnn, _
        CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, _
        Options:=adCmdText

    '==============================================
    
    Sheet20.Range("A1").CopyFromRecordset rst            'we paste all retrieved data into cell A1

    '==============================================
    'Sheet20.Columns("F:F").NumberFormat = "hh:mm:ss;@" 'format time column as time as it will keep resetting
    Sheet20.Columns("C:P").NumberFormat = "dd/mm/yyyy;@"
    '==============================================
    Set rst = Nothing
    cnn.Close
    Set cnn = Nothing
    '==============================================

End Sub

Private Sub CommandButton2_Click()
On Error Resume Next
Unload CalendarForm
Unload Me
Unload AssesmentAdd

End Sub



Private Sub countRows() 'find last row and column

    Dim lastRowNum As Long
    Dim lastColumnNum As Long
    
    lastRowNum = Sheet20.Cells(Rows.Count, 1).End(xlUp).Row
    lactColumnNum = Sheet20.Cells(1, Columns.Count).End(xlToLeft).Column
    
End Sub

Private Sub CommandButton5_Click()
    Call countRows
    
    Dim ctlSource As Control
    Set ctlSource = AssesmentAdd.ListBox1
    Dim strItems As String
    Dim intCurrentRow As Integer
    Dim numerRow As Integer
    numerRow = 1
    
    For intCurrentRow = 0 To ctlSource.ListCount - 1
        If ctlSource.Selected(intCurrentRow) Then
            strItems = strItems & ctlSource.Column(0, intCurrentRow)
            'MsgBox "Selected ID: " & strItems & vbNewLine & "Editing is not yet implemented in this project."
            Exit For
        End If
        
        numerRow = numerRow + 1
    
    Next intCurrentRow
     
    Set ctlSource = Nothing

    '==============================
    rownums = numerRow
    resetme (rownums)
End Sub

Private Sub CommandButton6_Click()

If Label22.Caption = NullString Then Exit Sub

answer = MsgBox("Selected record will now be updated, do you wish to continue?", vbOKCancel, "Confirmation required")
If answer = vbCancel Then Exit Sub



'On Error Resume Next
'============================================== 'firstly let's check if we can access a database
    Dim BasePath As String
        BasePath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
    Dim FilesPath As String
        FilesPath = ""              'clear file path
        FilesPath = Dir(BasePath)   'reassign

    If FilesPath = "" Then
        MsgBox "Could not connect to database. Try again later.", vbCritical, "Could not connect"
        Exit Sub 'file dosn't exist
    End If
'=============================================

Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim fld As ADODB.Field
Dim MyConn

Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
MyConn = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"


cnn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & MyConn & ";Persist Security Info=False;"
cnn.Open

With rst
    .Open "Assesments", cnn, adOpenKeyset, adLockPessimistic, adCmdTable
End With

rst.Filter = "ID = '" & AssesmentAdd.Label22.Caption & "'"

rst!Names = TextBox1.text
rst!B1 = TextBox2.text
rst!B2 = DateValue(TextBox3.text)
rst!A1 = DateValue(TextBox4.text)
rst!A2 = DateValue(TextBox5.text)
rst!H1 = DateValue(TextBox6.text)
rst!F1 = DateValue(TextBox7.text)
rst!P1 = DateValue(TextBox8.text)
rst!M3A = DateValue(TextBox9.text)
rst!M3B = DateValue(TextBox10.text)
rst!A4 = DateValue(TextBox11.text)
rst!A5 = DateValue(TextBox12.text)
rst!D1 = DateValue(TextBox13.text)
rst!Remote = DateValue(TextBox14.text)
rst!Assessment = DateValue(TextBox15.text)
rst!Comments = TextBox16.text
rst!Site = ComboBox1.text
rst!Shift = ComboBox2.text

rst.Update
rst.Close

Set rst = Nothing
cnn.Close
Set cnn = Nothing

'=====================================
Sheet20.Cells(rownums, 2).value = TextBox1.text
Sheet20.Cells(rownums, 3).value = DateValue(TextBox2.text)
Sheet20.Cells(rownums, 4).value = DateValue(TextBox3.text)
Sheet20.Cells(rownums, 5).value = DateValue(TextBox4.text)
Sheet20.Cells(rownums, 6).value = DateValue(TextBox5.text)
Sheet20.Cells(rownums, 7).value = DateValue(TextBox6.text)
Sheet20.Cells(rownums, 8).value = DateValue(TextBox7.text)
Sheet20.Cells(rownums, 9).value = DateValue(TextBox8.text)
Sheet20.Cells(rownums, 10).value = DateValue(TextBox9.text)
Sheet20.Cells(rownums, 11).value = DateValue(TextBox10.text)
Sheet20.Cells(rownums, 12).value = DateValue(TextBox11.text)
Sheet20.Cells(rownums, 13).value = DateValue(TextBox12.text)
Sheet20.Cells(rownums, 14).value = DateValue(TextBox13.text)
Sheet20.Cells(rownums, 15).value = DateValue(TextBox14.text)
Sheet20.Cells(rownums, 16).value = DateValue(TextBox15.text)
Sheet20.Cells(rownums, 17).value = TextBox16.text
Sheet20.Cells(rownums, 18).value = ComboBox1.text
Sheet20.Cells(rownums, 19).value = ComboBox2.text
'=====================================
 
'MsgBox "Database was updated.", vbInformation, "Procedure completed."

'=====================================

Call paintList
resetme (rownums)
Me.Repaint

End Sub

Private Sub CommandButton1_Click()

If TextBox1.text = NullString Then Exit Sub

If ComboBox1.text = NullString Or ComboBox2 = NullString Then
    MsgBox "You must enter site and shift name first."
    Exit Sub
End If

    
 '=============================== CHECK IF DATABASE IS ACCESSIBLE
Dim FPath As String
FPath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
If Dir(FPath) = "" Then
    MsgBox "Database is not accessible. Please try again later.", vbOKOnly, "Could not find database."
    Exit Sub
End If
'======================================================================

'ADD

    'On Error Resume Next ' in case of type mismatch
    
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim MyConn
        MyConn = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb" 'adres do bazy


    Dim j As Long           'element number in data stream
    Dim k As Integer        'column counter
    Dim sSQL As String      'SQL command string
    
    
            
    sSQL = "SELECT * FROM Assesments"     'sql query - select * from selected database
    
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
                
                rst(1).value = TextBox1.value
                If TextBox2.value <> NullString Then rst(2) = DateValue(TextBox2.value) Else rst(2) = NullString
                If TextBox3.value <> NullString Then rst(3) = DateValue(TextBox3.value) Else rst(3) = NullString
                If TextBox4.value <> NullString Then rst(4) = DateValue(TextBox4.value) Else rst(4) = NullString            'now we are adding data to it
                If TextBox5.value <> NullString Then rst(5) = DateValue(TextBox5.value) Else rst(5) = NullString          'from the booking form
                If TextBox6.value <> NullString Then rst(6) = DateValue(TextBox6.value) Else rst(6) = NullString
                If TextBox7.value <> NullString Then rst(7) = DateValue(TextBox7.value) Else rst(7) = NullString
                If TextBox8.value <> NullString Then rst(8) = DateValue(TextBox8.value) Else rst(8) = NullString
                If TextBox9.value <> NullString Then rst(9) = DateValue(TextBox9.value) Else rst(9) = NullString
                If TextBox10.value <> NullString Then rst(10) = DateValue(TextBox10.value) Else rst(10) = NullString
                If TextBox11.value <> NullString Then rst(11) = DateValue(TextBox11.value) Else rst(11) = NullString
                If TextBox12.value <> NullString Then rst(12) = DateValue(TextBox12.value) Else rst(12) = NullString
                If TextBox13.value <> NullString Then rst(13) = DateValue(TextBox13.value) Else rst(13) = NullString
                If TextBox14.value <> NullString Then rst(14) = DateValue(TextBox14.value) Else rst(14) = NullString
                If TextBox15.value <> NullString Then rst(15) = DateValue(TextBox15.value) Else rst(15) = NullString
                rst(16).value = TextBox16.value
                rst(17).value = UCase(ComboBox1.text)
                rst(18).value = UCase(ComboBox2.text)
                

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
    
            TextBox1.value = ""            'clearing up old data remains from the form
            TextBox2.value = "" 'from the booking form
            
            ComboBox1.value = ""
            ComboBox2.value = ""
            
            TextBox3.value = ""
            TextBox4.value = ""
            TextBox5.value = ""
            TextBox6.value = ""
            TextBox7.value = ""
            TextBox8.value = ""
            TextBox9.value = ""
            TextBox10.value = ""
            TextBox11.value = ""
            TextBox12.value = ""
            TextBox13.value = ""
            TextBox14.value = ""
            TextBox15.value = ""
            TextBox16.value = ""

            TextBox1.SetFocus
            Unload CalendarForm
            MsgBox "Entry Added to db.", vbInformation, "OK"

End Sub

Private Sub Label21_Click()

End Sub

Private Sub TextBox2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Debug.Print "clicking textbox 2 on add page"
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox2.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox3.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox4.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox5.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox6.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox7.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox8.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox9_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox9.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox10_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox10.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox11_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox11.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox12_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox12.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox13_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox13.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox14_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox14.text = dateVariable
    Unload CalendarForm
End Sub
Private Sub TextBox15_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then AssAdd.TextBox15.text = dateVariable
    Unload CalendarForm
End Sub

Private Sub UserForm_Activate()

    On Error GoTo 0
    Unload CalendarForm
    Debug.Print "page loaded"
    Call isAvailable("J:\Pub-LOGISTICS\Packaging\Packaging.accdb")
    
End Sub



Private Sub UserForm_Initialize()

    ComboBox1.AddItem "RED1"
    ComboBox1.AddItem "RED2"
    ComboBox1.AddItem "DRO"
    ComboBox1.AddItem "ALL"
    ComboBox1.AddItem "OTHER"
    ComboBox1.AddItem "LEFT"
    
    ComboBox2.AddItem "ORANGE"
    ComboBox2.AddItem "GREEN"
    ComboBox2.AddItem "YELLOW"
    ComboBox2.AddItem "BLUE"
    ComboBox2.AddItem "RED"
    ComboBox2.AddItem "ALL"
    ComboBox2.AddItem "OTHER"
    ComboBox2.AddItem "LEFT"
    
End Sub

Private Sub UserForm_Terminate()
On Error Resume Next
Unload CalendarForm
Unload Me
Unload AssesmentAdd

End Sub
