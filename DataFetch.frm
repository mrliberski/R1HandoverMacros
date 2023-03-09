VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataFetch 
   Caption         =   "Fetch Data"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3288
   OleObjectBlob   =   "DataFetch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataFetch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click() 'fetch all button

'Fetch all data from database and paste into B2

    On Error Resume Next ' in case of type mismatch
    
    '==============================================
    Dim BasePath As String
        BasePath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'BasePath = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"
        
    Dim FilesPath As String
        FilesPath = ""              'clear file path
        FilesPath = Dir(BasePath)   'reassign


    If FilesPath = "" Then
        MsgBox "Could not connect to database. It may not exist, or you may have no permission to access it. Try again later.", vbCritical, "Could not connect"
        Exit Sub 'file dosn't exist
    End If
    '==============================================
    '==============================================
    '==============================================
    '==============================================
    '==============================================
    '==============================================
    'CLEAR CONTENT
    Sheet18.Select
    Range("A2:O9999").Select
    Selection.ClearContents
    '==============================================
    
    Dim cnn As ADODB.Connection
        Set cnn = New ADODB.Connection
    Dim rst As ADODB.Recordset
        Set rst = New ADODB.Recordset
    Dim fld As ADODB.Field
    Dim MyConn
        MyConn = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'MyConn = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"   'adres do bazy
    Dim sSQL As String                                                'SQL command string
        sSQL = "SELECT * FROM Packaging_Log"                          'sql query - select * from selected database
    
    '==============================================
    
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn                                      'open connection
    End With
    
    rst.CursorLocation = adUseServer
    rst.Open source:=sSQL, ActiveConnection:=cnn, _
        CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, _
        Options:=adCmdText

    '==============================================
    
    Sheet18.Range("A2").CopyFromRecordset rst            'we paste all retrieved data into cell A2

    '==============================================
    Sheet18.Columns("F:F").Select                               'format time column as time as it will keep resetting
    Selection.NumberFormat = "hh:mm:ss;@"
    '==============================================
    
    Set rst = Nothing
    cnn.Close
    Set cnn = Nothing
    '==============================================
    DataFetch.Hide
    '==============================================
    Sheet18.Select
    Range("A2").Select



End Sub
Sub getDataFromAccess() ' SAmple bit

Dim DBFullName As String
Dim connect As String, source As String
Dim Connection As ADODB.Connection
Dim Recordset As ADODB.Recordset
Dim Col As Integer
Dim startdt As String
Dim stopdt As String
Dim refresh

refresh = MsgBox("Start New Query?", vbYesNo)
If refresh = vbYes Then
    Sheet18.Cells.Clear
    startdt = Application.InputBox("Please Input Start Date for Query (MM/DD/YYYY): ", "Start Date")
    stopdt = Application.InputBox("Please Input Stop Date for Query (MM/DD/YYYY): ", "Stop Date")

    DBFullName = "X:\MyDocuments\CMS\CMS Database.mdb"
    ' Open the connection
    Set Connection = New ADODB.Connection
    connect = "Provider=Microsoft.ACE.OLEDB.12.0;"
    connect = connect & "Data Source=" & DBFullName & ";"
    Connection.Open connectionString:=connect

    Set Recordset = New ADODB.Recordset
    With Recordset
        source = "SELECT * FROM Tracking WHERE Date_Logged BETWEEN " & startdt & " AND " & stopdt & " ORDER BY Date_Logged"
        .Open source:=source, ActiveConnection:=Connection

        For Col = 0 To Recordset.Fields.Count - 1
            Range(“A1”).Offset(0, Col).value = Recordset.Fields(Col).name
        Next

        Range(“A1”).Offset(1, 0).CopyFromRecordset Recordset
    End With
    ActiveSheet.Columns.AutoFit
   Set Recordset = Nothing
    Connection.Close
    Set Connection = Nothing

End Sub

Private Sub CommandButton2_Click() 'fetch data button

    '============================================== 'firstly let's check if we can access a database
    Dim BasePath As String
        BasePath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'BasePath = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"
    Dim FilesPath As String
        FilesPath = ""              'clear file path
        FilesPath = Dir(BasePath)   'reassign

    If FilesPath = "" Then
        MsgBox "Could not connect to database. It may not exist, or you may have no permission to access it. Try again later.", vbCritical, "Could not connect"
        Exit Sub 'file dosn't exist
    End If

    '============================================== ' SQL QUERY BUILD

    Dim startDate As String         'these two define date range we want to look up
    Dim StopDate As String          'stop date is optional
    
    If TextBox1.text = "" Then
        MsgBox "You must select date first", vbInformation, "Date not selected."
        Exit Sub
    End If
    
    startDate = Format(DateValue(TextBox1.text), "dd\/mmm\/yyyy")               'start date was selected

    If TextBox2.text = "" Then              'if stop date is not selected we'll just assume we want it to be the same as start date
        StopDate = Format(DateValue(TextBox1.text), "dd\/mmm\/yyyy")
    Else: StopDate = Format(DateValue(TextBox2.text), "dd\/mmm\/yyyy")
    End If
    
    Dim SqlString As String
        SqlString = "SELECT * FROM Packaging_Log WHERE "
        SqlString = SqlString & "DelDate BETWEEN #" & startDate & "# "
        SqlString = SqlString & "AND #" & StopDate & "# "
        
    If OptionButton2.value = True Then
        SqlString = SqlString & "AND (ReceiveQty < AdvisedQty OR ReceiveQty > AdvisedQty) "
    End If
'==============================================
    If CheckBox1.value = True And CheckBox2.value = True Then
        SqlString = SqlString
    ElseIf CheckBox1.value = False And CheckBox2.value = False Then
        SqlString = SqlString
    ElseIf CheckBox1.value = True Then
        SqlString = SqlString & "AND [ComplaintNo] IS NOT NULL "
    ElseIf CheckBox2.value = True Then
        SqlString = SqlString & "AND [ComplaintNo] IS NULL "
    End If
'==============================================
    SqlString = SqlString & "ORDER BY DelDate"
'==============================================
    
    '============================================== ' clear form
        'CLEAR CONTENT
    Sheet18.Select
    Range("A2:O9999").Select
    Selection.ClearContents
    '==============================================
    Dim cnn As ADODB.Connection
        Set cnn = New ADODB.Connection
    Dim rst As ADODB.Recordset
        Set rst = New ADODB.Recordset
    Dim fld As ADODB.Field
    Dim MyConn
        MyConn = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'MyConn = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"   'adres do bazy
    '==============================================
    
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn                                      'open connection
    End With
    
    rst.CursorLocation = adUseServer
    rst.Open source:=SqlString, ActiveConnection:=cnn, _
        CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, _
        Options:=adCmdText

    '==============================================
    
    Sheet18.Range("A2").CopyFromRecordset rst            'we paste all retrieved data into cell A2

    '==============================================
    Sheet18.Columns("F:F").Select                               'format time column as time as it will keep resetting
    Selection.NumberFormat = "hh:mm:ss;@"
    '==============================================
    
    Set rst = Nothing
    cnn.Close
    Set cnn = Nothing
    '==============================================
    DataFetch.Hide
    '==============================================
    Sheet18.Select
    Range("A2").Select
    '==============================================
    
    '==============================================
    
    '==============================================
        
    
End Sub

Private Sub CommandButton3_Click() ' close button
    DataFetch.Hide
End Sub

Private Sub CommandButton4_Click()
'find delivery note
    '============================================== 'firstly let's check if we can access a database
    Dim BasePath As String
        BasePath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'BasePath = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"
    Dim FilesPath As String
        FilesPath = ""              'clear file path
        FilesPath = Dir(BasePath)   'reassign

    If FilesPath = "" Then
        MsgBox "Could not connect to database. It may not exist, or you may have no permission to access it. Try again later.", vbCritical, "Could not connect"
        Exit Sub 'file dosn't exist
    End If

    '============================================== ' SQL QUERY BUILD

    Dim startDate As String         'these two define date range we want to look up
    Dim StopDate As String          'stop date is optional
    
    'If TextBox1.Text = "" Then
        'MsgBox "You must select date first", vbInformation, "Date not selected."
        'Exit Sub
   ' End If
    
    'StartDate = Format(DateValue(TextBox1.Text), "dd\/mm\/yyyy")               'start date was selected
    
    'If TextBox2.Text = "" Then              'if stop date is not selected we'll just assume we want it to be the same as start date
    '    StopDate = Format(DateValue(TextBox1.Text), "dd\/mm\/yyyy")
    'Else: StopDate = Format(DateValue(TextBox2.Text), "dd\/mm\/yyyy")
    'End If

    Dim SqlString As String
        SqlString = "SELECT * FROM Packaging_Log WHERE "
        SqlString = SqlString & "DelNo = '" & TextBox3.text & "'"

        
    
    '============================================== ' clear form
        'CLEAR CONTENT
    Sheet18.Select
    Range("A2:O9999").Select
    Selection.ClearContents
    '==============================================
    Dim cnn As ADODB.Connection
        Set cnn = New ADODB.Connection
    Dim rst As ADODB.Recordset
        Set rst = New ADODB.Recordset
    Dim fld As ADODB.Field
    Dim MyConn
        MyConn = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'MyConn = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"   'adres do bazy
    '==============================================
    
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn                                      'open connection
    End With
    
    rst.CursorLocation = adUseServer
    rst.Open source:=SqlString, ActiveConnection:=cnn, _
        CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, _
        Options:=adCmdText

    '==============================================
    
    Sheet18.Range("A2").CopyFromRecordset rst            'we paste all retrieved data into cell A2

    '==============================================
    Sheet18.Columns("F:F").Select                               'format time column as time as it will keep resetting
    Selection.NumberFormat = "hh:mm:ss;@"
    '==============================================
    
    Set rst = Nothing
    cnn.Close
    Set cnn = Nothing
    '==============================================
    DataFetch.Hide
    '==============================================
    Sheet18.Select
    Range("A2").Select
    '==============================================
    
    '==============================================
    
    '==============================================
        

End Sub

Private Sub CommandButton5_Click() ' add complaint number

    '============================================== 'firstly let's check if we can access a database
    Dim BasePath As String
        BasePath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'BasePath = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"
    Dim FilesPath As String
        FilesPath = ""              'clear file path
        FilesPath = Dir(BasePath)   'reassign

    If FilesPath = "" Then
        MsgBox "Could not connect to database. It may not exist, or you may have no permission to access it. Try again later.", vbCritical, "Could not connect"
        Exit Sub 'file dosn't exist
    End If
    '=============================================

    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim MyConn
        'MyConn = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"   'adres do bazy
    '==============================================
    '=============================================
    'see if ID number is empty or not
Dim SqlString As String
For i = 2 To 1000
    If Not Sheets("Data Fetch").Cells(i, 1).value = "" Then
        If Not Sheets("Data Fetch").Cells(i, 15).value = "" Then
            Set cnn = New ADODB.Connection
            Set rst = New ADODB.Recordset
            MyConn = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
            With cnn
                .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Open MyConn                                      'open connection
            End With
            '=============================================
            SqlString = "UPDATE Packaging_Log SET [ComplaintNo] = " & Sheet18.Cells(i, 15).value & " WHERE [ID] = " & Sheet18.Cells(i, 1).value
                rst.CursorLocation = adUseServer
                rst.Open source:=SqlString, ActiveConnection:=cnn, _
                    CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, _
                    Options:=adCmdText
            Set rst = Nothing
            cnn.Close
            Set cnn = Nothing
                    
        End If
    End If
Next i
    '=============================================
    '==============================================
    '==============================================
    '==============================================
    '==============================================
'============================================== ' clear form
    'CLEAR CONTENT
    Sheet18.Select
    Range("A2:O9999").Select
    Selection.ClearContents
'==============================================
    DataFetch.Hide
    '==============================================
    Sheet18.Select
    Range("A2").Select
    '==============================================
MsgBox "Records were amended succesfully", vbInformation, "Procedure completed."

End Sub

Private Sub CommandButton6_Click() 'LAST WEEK
    '============================================== 'firstly let's check if we can access a database
    Dim BasePath As String
        BasePath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'BasePath = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"
    Dim FilesPath As String
        FilesPath = ""              'clear file path
        FilesPath = Dir(BasePath)   'reassign

    If FilesPath = "" Then
        MsgBox "Could not connect to database. It may not exist, or you may have no permission to access it. Try again later.", vbCritical, "Could not connect"
        Exit Sub 'file dosn't exist
    End If
    '=============================================
    '============================================== ' clear form
        'CLEAR CONTENT
    Sheet18.Select
    Range("A2:O9999").Select
    Selection.ClearContents
    '==============================================
    '============================================== ' SQL QUERY BUILD

    Dim startDate As String         'these two define date range we want to look up
    Dim StopDate As String          'stop date is optional
    
    startDate = Format(DateValue(Date) - 7, "dd\/mmm\/yyyy")    'start date was set to 7 days before today
    StopDate = Format(DateValue(Date), "dd\/mmm\/yyyy")         'end date is today
    
    
    Dim SqlString As String
        SqlString = "SELECT * FROM Packaging_Log WHERE "
        SqlString = SqlString & "DelDate BETWEEN #" & startDate & "# "
        SqlString = SqlString & "AND #" & StopDate & "# "
        
    If OptionButton2.value = True Then
        SqlString = SqlString & "AND (ReceiveQty < AdvisedQty OR ReceiveQty > AdvisedQty) "
    End If

    If CheckBox1.value = True And CheckBox2.value = True Then
        SqlString = SqlString
    ElseIf CheckBox1.value = False And CheckBox2.value = False Then
        SqlString = SqlString
    ElseIf CheckBox1.value = True Then
        SqlString = SqlString & "AND [ComplaintNo] IS NOT NULL "
    ElseIf CheckBox2.value = True Then
        SqlString = SqlString & "AND [ComplaintNo] IS NULL "
    End If

    SqlString = SqlString & "ORDER BY DelDate"
    '==============================================
    '============================================== ' define connections
    Dim cnn As ADODB.Connection
        Set cnn = New ADODB.Connection
    Dim rst As ADODB.Recordset
        Set rst = New ADODB.Recordset
    Dim fld As ADODB.Field
    Dim MyConn
        MyConn = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'MyConn = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"   'adres do bazy
    '==============================================
    
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn                                      'open connection
    End With
    
    rst.CursorLocation = adUseServer
    rst.Open source:=SqlString, ActiveConnection:=cnn, _
        CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, _
        Options:=adCmdText

    '==============================================
    
    Sheet18.Range("A2").CopyFromRecordset rst            'we paste all retrieved data into cell A2

    '==============================================
    Sheet18.Columns("F:F").Select                               'format time column as time as it will keep resetting
    Selection.NumberFormat = "hh:mm:ss;@"
    '==============================================
    
    Set rst = Nothing
    cnn.Close
    Set cnn = Nothing
    '==============================================
    DataFetch.Hide
    '==============================================
    Sheet18.Select
    Range("A2").Select
    '==============================================
End Sub

Private Sub CommandButton7_Click() ' LAST 30 days
    '============================================== 'firstly let's check if we can access a database
    Dim BasePath As String
        BasePath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'BasePath = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"
    Dim FilesPath As String
        FilesPath = ""              'clear file path
        FilesPath = Dir(BasePath)   'reassign

    If FilesPath = "" Then
        MsgBox "Could not connect to database. It may not exist, or you may have no permission to access it. Try again later.", vbCritical, "Could not connect"
        Exit Sub 'file dosn't exist
    End If
    '=============================================
    '============================================== ' clear form
        'CLEAR CONTENT
    Sheet18.Select
    Range("A2:O9999").Select
    Selection.ClearContents
    '==============================================
    '============================================== ' SQL QUERY BUILD

    Dim startDate As String         'these two define date range we want to look up
    Dim StopDate As String          'stop date is optional
    
    startDate = Format(DateValue(Date) - 30, "dd\/mmm\/yyyy")    'start date was set to 7 days before today
    StopDate = Format(DateValue(Date), "dd\/mmm\/yyyy")         'end date is today
    
    
    Dim SqlString As String
        SqlString = "SELECT * FROM Packaging_Log WHERE "
        SqlString = SqlString & "DelDate BETWEEN #" & startDate & "# "
        SqlString = SqlString & "AND #" & StopDate & "# "
        
    If OptionButton2.value = True Then
        SqlString = SqlString & "AND (ReceiveQty < AdvisedQty OR ReceiveQty > AdvisedQty) "
    End If

    If CheckBox1.value = True And CheckBox2.value = True Then
        SqlString = SqlString
    ElseIf CheckBox1.value = False And CheckBox2.value = False Then
        SqlString = SqlString
    ElseIf CheckBox1.value = True Then
        SqlString = SqlString & "AND [ComplaintNo] IS NOT NULL "
    ElseIf CheckBox2.value = True Then
        SqlString = SqlString & "AND [ComplaintNo] IS NULL "
    End If

    SqlString = SqlString & "ORDER BY DelDate"
    '==============================================
    '============================================== ' define connections
    Dim cnn As ADODB.Connection
        Set cnn = New ADODB.Connection
    Dim rst As ADODB.Recordset
        Set rst = New ADODB.Recordset
    Dim fld As ADODB.Field
    Dim MyConn
        MyConn = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'MyConn = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"   'adres do bazy
    '==============================================
    
    With cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn                                      'open connection
    End With
    
    rst.CursorLocation = adUseServer
    rst.Open source:=SqlString, ActiveConnection:=cnn, _
        CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, _
        Options:=adCmdText

    '==============================================
    
    Sheet18.Range("A2").CopyFromRecordset rst            'we paste all retrieved data into cell A2

    '==============================================
    Sheet18.Columns("F:F").Select                               'format time column as time as it will keep resetting
    Selection.NumberFormat = "hh:mm:ss;@"
    '==============================================
    
    Set rst = Nothing
    cnn.Close
    Set cnn = Nothing
    '==============================================
    DataFetch.Hide
    '==============================================
    Sheet18.Select
    Range("A2").Select
    '==============================================
End Sub

Private Sub OptionButton1_Click()
    OptionButton2.value = False
End Sub
Private Sub OptionButton2_Click()
    OptionButton1.value = False
End Sub

Private Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox1.text = dateVariable
End Sub
Private Sub TextBox2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox2.text = dateVariable
End Sub


Sub updateAccess() 'sample update code

Dim cn As ADODB.Connection

Dim rstProducts As ADODB.Recordset

Dim sProduct As String

Dim cPrice As String

Dim counter As Integer

Application.DisplayAlerts = False

Set cn = New ADODB.Connection

cn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Products.accdb;Persist Security Info=False;"

cn.Open

Set rstProducts = New ADODB.Recordset

With rstProducts

.Open "ProductTable", cn, adOpenKeyset, adLockPessimistic, adCmdTable

End With

sProduct = Sheet1.Cells(2, 1).value ' row 1 contains column headings

counter = 0

Do While Not sProduct = ""

sProduct = Sheet1.Cells(2 + counter, 1).value

cPrice = Sheet1.Cells(2 + counter, 2).value

rstProducts.Filter = "ProductName = '" & sProduct & "'"

If rstProducts.EOF Then

rstProducts.AddNew

rstProducts("rstProducts!ProductName").value = sProduct

rstProducts("rstProducts!Price").value = cPrice

Else

rstProducts!Price = cPrice

End If

rstProducts.Update

counter = counter + 1

sProduct = Sheet1.Cells(2 + counter, 1).value

Loop

rstProducts.Close

Set rstProducts = Nothing

cn.Close

Set cn = Nothing

Application.DisplayAlerts = True
 
anotherSample = "UPDATE Packaging_Log SET ComplaintNo = 'None' WHERE [Last Name] = 'Smith'"

End Sub

