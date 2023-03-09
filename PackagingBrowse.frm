VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PackagingBrowse 
   Caption         =   "Packaging Log View"
   ClientHeight    =   8400.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17988
   OleObjectBlob   =   "PackagingBrowse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PackagingBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public rownum
Public Function resetme(rownum)
    
    If rownum < 1 Then rownum = 1

    PackagingBrowse.TextBox5.text = Sheet18.Cells(rownum, 1).value
    PackagingBrowse.TextBox6.text = Sheet18.Cells(rownum, 2).value
    PackagingBrowse.TextBox7.text = Sheet18.Cells(rownum, 3).value
    PackagingBrowse.TextBox8.text = Sheet18.Cells(rownum, 4).value
    PackagingBrowse.TextBox9.text = Sheet18.Cells(rownum, 5).value
    PackagingBrowse.TextBox10.text = Format(Sheet18.Cells(rownum, 6).value, "hh:mm")
    PackagingBrowse.TextBox11.text = Sheet18.Cells(rownum, 7).value
    PackagingBrowse.TextBox12.text = Sheet18.Cells(rownum, 8).value
    PackagingBrowse.TextBox13.text = Sheet18.Cells(rownum, 9).value
    PackagingBrowse.TextBox14.text = Sheet18.Cells(rownum, 10).value
    PackagingBrowse.TextBox15.text = Sheet18.Cells(rownum, 11).value
    PackagingBrowse.TextBox16.text = Sheet18.Cells(rownum, 12).value
    PackagingBrowse.TextBox17.text = Sheet18.Cells(rownum, 13).value
    PackagingBrowse.TextBox19.text = Sheet18.Cells(rownum, 14).value
    PackagingBrowse.TextBox18.text = Sheet18.Cells(rownum, 15).value
    
    PackagingBrowse.Repaint
    
End Function



Private Sub isAvailable(ByVal adres As String)
    If Dir(adres) <> NullString Then
        Label2.Caption = "Database availability Status: Available"
    Else
        Label2.Caption = "Database availability Status: Unavailable"
    End If
End Sub

Private Sub CommandButton19_Click()
'updaterrrrr

On Error Resume Next
'============================================== 'firstly let's check if we can access a database
    Dim BasePath As String
        BasePath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'BasePath = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"
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
    .Open "Packaging_Log", cnn, adOpenKeyset, adLockPessimistic, adCmdTable
End With

rst.Filter = "ID = '" & TextBox5.text & "'"

rst!DelDate = TextBox9.text
rst!DelTime = TextBox10.text
rst!Shift = TextBox11.text
rst!Customer = TextBox12.text
rst!RegNo = TextBox13.text
rst!DelNo = TextBox14.text
rst!PackCode = TextBox15.text
rst!ReceiveQty = TextBox16.text
rst!AdvisedQty = TextBox17.text
rst!Comments = TextBox19.text
rst!ComplaintNo = TextBox18.text
rst!UserName = Environ("UserName")

rst.Update
rst.Close

Set rst = Nothing
cnn.Close
Set cnn = Nothing

'=====================================
Sheet18.Cells(rownum, 2).value = TextBox6.text
Sheet18.Cells(rownum, 3).value = Environ("UserName")
Sheet18.Cells(rownum, 4).value = TextBox8.text
Sheet18.Cells(rownum, 5).value = TextBox9.text
Sheet18.Cells(rownum, 6).value = TextBox10.text
Sheet18.Cells(rownum, 7).value = TextBox11.text
Sheet18.Cells(rownum, 8).value = TextBox12.text
Sheet18.Cells(rownum, 9).value = TextBox13.text
Sheet18.Cells(rownum, 10).value = TextBox14.text
Sheet18.Cells(rownum, 11).value = TextBox15.text
Sheet18.Cells(rownum, 12).value = TextBox16.text
Sheet18.Cells(rownum, 13).value = TextBox17.text
Sheet18.Cells(rownum, 14).value = TextBox19.text
Sheet18.Cells(rownum, 15).value = TextBox18.text
'=====================================
 
'MsgBox "Database was updated.", vbInformation, "Procedure completed."

'=====================================

Call paintList
resetme (rownum)

End Sub

Private Sub CommandButton20_Click()
'deletor

On Error Resume Next
'============================================== 'firstly let's check if we can access a database
    Dim BasePath As String
        BasePath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
        'BasePath = "C:\Users\pawel.liberski\Desktop\Packaging.accdb"
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
    .Open "Packaging_Log", cnn, adOpenKeyset, adLockPessimistic, adCmdTable
End With

rst.Filter = "ID = '" & TextBox5.text & "'"
rst.Delete
rst.Close

Set rst = Nothing
cnn.Close
Set cnn = Nothing
                    
                    
'=====================================
Sheet18.Rows(rownum).EntireRow.Delete
'=====================================
                
 
MsgBox "Record was deleted.", vbInformation, "Procedure completed."

'=====================================
Call paintList
rownum = rownum - 1
resetme (rownum)
End Sub

Private Sub CommandButton21_Click()
'next
    If Sheet18.Cells(rownum + 1, 1).value <> "" Then
        rownum = rownum + 1
        resetme (rownum)
        Debug.Print "going up " & rownum
    End If

End Sub

Private Sub CommandButton22_Click()
    If rownum > 1 Then
        rownum = rownum - 1
        resetme (rownum)
        Debug.Print "going down " & rownum
    End If
End Sub

Private Sub CommandButton3_Click() ' search

    Call isAvailable("J:\Pub-LOGISTICS\Packaging\Packaging.accdb")
    Call checkDatabase
    Call clearContent

'======= SQL QUERY BUILD ===============================================
Dim startDate As String
Dim endDate As String

If TextBox1.text = "" Or TextBox1.text = "Select Start Date" Then
    startDate = Format(DateValue(Date) - 30, "dd\/mmm\/yyyy")
Else
    startDate = Format(DateValue(TextBox1.text), "dd\/mmm\/yyyy")
End If

If TextBox2.text = "" Or TextBox2.text = "Select End Date" Then
    endDate = Format(DateValue(Date), "dd\/mmm\/yyyy")
Else
    endDate = Format(DateValue(TextBox2.text), "dd\/mmm\/yyyy")
End If

    Dim SqlString As String
        SqlString = "SELECT * FROM Packaging_Log WHERE "
        SqlString = SqlString & "DelDate BETWEEN #" & startDate & "# "
        SqlString = SqlString & "AND #" & endDate & "# "

If TextBox3.text <> "Enter Packaging No" Then
    SqlString = SqlString & "AND [PackCode] = '" & Trim(TextBox3.text) & "' "
End If

If TextBox4.text <> "Enter DN Number" Then
    SqlString = SqlString & "AND [DelNo] = '" & Trim(TextBox4.text) & "' "
End If

        
    If CheckBox2.value = True Then
        SqlString = SqlString & "AND (ReceiveQty < AdvisedQty OR ReceiveQty > AdvisedQty) "
    ElseIf CheckBox3.value = True Then
        SqlString = SqlString & "AND [ComplaintNo] IS NOT NULL "
    ElseIf CheckBox4.value = True Then
        SqlString = SqlString & "AND [ComplaintNo] IS NULL "
    End If

    SqlString = SqlString & "ORDER BY DelDate"
    
    Debug.Print SqlString

'=======================================================================

    Call connectExecute(SqlString)
    Call paintList
    PackagingBrowse.Repaint

End Sub
Private Sub countRows() 'find last row and column

    Dim lastRowNum As Long
    Dim lastColumnNum As Long
    
    lastRowNum = Sheet18.Cells(Rows.Count, 1).End(xlUp).Row
    lactColumnNum = Cells(1, Columns.Count).End(xlToLeft).Column
    
End Sub
Private Sub checkDatabase()
    '============================================== 'firstly let's check if we can access a database
    If Dir("J:\Pub-LOGISTICS\Packaging\Packaging.accdb") = NullString Then
        MsgBox "Could not connect to database. Try again later.", vbCritical, "Resource off limits."
        Exit Sub 'file dosn't exist or is not accessible
    End If
    '=============================================
End Sub
Private Sub clearContent()
    '============================================== ' clear form
    'CLEAR CONTENT
    Dim lastRowNum As Long
    lastRowNum = Sheet18.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Sheet18.Select
    Sheet18.Range("A1:O" & lastRowNum).ClearContents
    'Selection.ClearContents
    '==============================================
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
    
    Sheet18.Range("A1").CopyFromRecordset rst            'we paste all retrieved data into cell A1

    '==============================================
    Sheet18.Columns("F:F").NumberFormat = "hh:mm:ss;@" 'format time column as time as it will keep resetting
    Sheet18.Columns("B:B").NumberFormat = "dd/mm/yyyy;@"
    'Selection.NumberFormat = "hh:mm:ss;@"
    '==============================================
    Set rst = Nothing
    cnn.Close
    Set cnn = Nothing
    '==============================================

End Sub



Private Sub CommandButton4_Click() ' LOAD LAST 30 DAYS
    
    Call isAvailable("J:\Pub-LOGISTICS\Packaging\Packaging.accdb")
    Call checkDatabase
    Call clearContent
    
    '============================================== ' SQL QUERY BUILD
    Dim startDate As String         'these two define date range we want to look up
    Dim StopDate As String          'stop date is optional
    
    startDate = Format(DateValue(Date) - 30, "dd\/mmm\/yyyy")    'start date was set to 7 days before today
    StopDate = Format(DateValue(Date), "dd\/mmm\/yyyy")         'end date is today
    
    
    Dim SqlString As String
        SqlString = "SELECT * FROM Packaging_Log WHERE "
        SqlString = SqlString & "DelDate BETWEEN #" & startDate & "# "
        SqlString = SqlString & "AND #" & StopDate & "# "
        
    If CheckBox2.value = True Then
        SqlString = SqlString & "AND (ReceiveQty < AdvisedQty OR ReceiveQty > AdvisedQty) "
    ElseIf CheckBox3.value = True Then
        SqlString = SqlString & "AND [ComplaintNo] IS NOT NULL "
    ElseIf CheckBox4.value = True Then
        SqlString = SqlString & "AND [ComplaintNo] IS NULL "
    End If

    SqlString = SqlString & "ORDER BY DelDate"
    '==============================================

    Call connectExecute(SqlString)
    Call paintList
    PackagingBrowse.Repaint

End Sub

Private Sub CommandButton5_Click() ' load last week

    Call isAvailable("J:\Pub-LOGISTICS\Packaging\Packaging.accdb")
    Call checkDatabase
    Call clearContent
    
    '============================================== ' SQL QUERY BUILD
    Dim startDate As String         'these two define date range we want to look up
    Dim StopDate As String          'stop date is optional
    
    startDate = Format(DateValue(Date) - 7, "dd\/mmm\/yyyy")    'start date was set to 7 days before today
    StopDate = Format(DateValue(Date), "dd\/mmm\/yyyy")         'end date is today
    
    
    Dim SqlString As String
        SqlString = "SELECT * FROM Packaging_Log WHERE "
        SqlString = SqlString & "DelDate BETWEEN #" & startDate & "# "
        SqlString = SqlString & "AND #" & StopDate & "# "
        
    If CheckBox2.value = True Then
        SqlString = SqlString & "AND (ReceiveQty < AdvisedQty OR ReceiveQty > AdvisedQty) "
    ElseIf CheckBox3.value = True Then
        SqlString = SqlString & "AND [ComplaintNo] IS NOT NULL "
    ElseIf CheckBox4.value = True Then
        SqlString = SqlString & "AND [ComplaintNo] IS NULL "
    End If

    SqlString = SqlString & "ORDER BY DelDate"
    '==============================================

    Call connectExecute(SqlString)
    Call paintList
    PackagingBrowse.Repaint

End Sub

Private Sub CommandButton6_Click() ' load all data

    Call isAvailable("J:\Pub-LOGISTICS\Packaging\Packaging.accdb")
    Call checkDatabase
    Call clearContent
    
    '============================================== ' SQL QUERY BUILD
    Dim SqlString As String
        SqlString = "SELECT * FROM Packaging_Log "
        
    If CheckBox2.value = True Then
        SqlString = SqlString & "WHERE ([ReceiveQty] < [AdvisedQty] OR [ReceiveQty] > [AdvisedQty]) "
    End If

    SqlString = SqlString & "ORDER BY DelDate"
    '==============================================

    Call connectExecute(SqlString)
    Call paintList
    PackagingBrowse.Repaint


End Sub

Private Sub CommandButton7_Click() 'editor

    Call countRows
    
    Dim ctlSource As Control
    Set ctlSource = PackagingBrowse.ListBox1
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
    rownum = numerRow
    resetme (rownum)

End Sub

Private Sub ListBox1_Click()
    Call countRows
    
    Dim ctlSource As Control
    Set ctlSource = PackagingBrowse.ListBox1
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
    rownum = numerRow
    resetme (rownum)
End Sub

Private Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox1.text = dateVariable
End Sub
Private Sub TextBox2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox2.text = dateVariable
End Sub


Private Sub TextBox3_Change()
    If TextBox3.text = vbNullString Then
        TextBox3.text = "Enter Packaging No"
    End If
    If TextBox4.text = vbNullString Or TextBox4.text = " " Then
        TextBox4.text = "Enter DN Number"
    End If
    
    
    
End Sub
Private Sub TextBox3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBox3.text = " "
End Sub
Private Sub TextBox4_Change()
    If TextBox4.text = vbNullString Then
        TextBox4.text = "Enter DN Number"
    End If
    If TextBox3.text = vbNullString Or TextBox3.text = " " Then
        TextBox3.text = "Enter Packaging No"
    End If
    
    
    
End Sub
Private Sub TextBox4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBox4.text = " "
End Sub








Private Sub UserForm_Activate()
    On Error Resume Next
    Call isAvailable("J:\Pub-LOGISTICS\Packaging\Packaging.accdb")
    Call paintList
    resetme (1)
End Sub
Private Sub paintList()

' Purpose:  fill listbox with range values after clicking on CommandButton1
'           (code could be applied to UserForm_Initialize(), too)
' Note:     based on @Siddharth-Rout 's proposal at https://stackoverflow.com/questions/10763310/how-to-populate-data-from-a-range-multiple-rows-and-columns-to-listbox-with-vb
'           but creating a variant data field array directly from range in a one liner
'           (instead of filling a redimensioned array with range values in a loop)
Dim ws      As Worksheet
Dim rng     As Range
Dim myArray                 ' variant, receives one based 2-dim data field array
'~~> Change your sheetname here
Set ws = Sheets("Data Fetch")

'~~> Set you relevant range here
Set rng = ws.Range("A1:O" & ws.Range("A" & ws.Rows.Count).End(xlUp).Row)

With Me.ListBox1
    .Clear
    .ColumnHeads = False
    .ColumnCount = rng.Columns.Count

    '~~> create a one based 2-dim datafield array
     myArray = rng

    '~~> fill listbox with array values
    .List = myArray

    '~~> Set the widths of the column here. Ex: For 5 Columns
    '~~> Change as Applicable
    .ColumnWidths = "30;0;0;0;65;0;0;0;65;65;65;65;65;150;50"
    .TopIndex = 0
End With


End Sub

Sub OKButtonInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make OK Button appear Green when hovered on

  CancelButtonInactive.Visible = True
  OKButtonInactive.Visible = False

End Sub
Sub CancelButtonInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button appear Green when hovered on

CancelButtonInactive.Visible = False
OKButtonInactive.Visible = True

End Sub
Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

CancelButtonInactive.Visible = True
OKButtonInactive.Visible = True

End Sub
Sub OKButton_Click()

    Call countRows
    
    Dim ctlSource As Control
    Set ctlSource = PackagingBrowse.ListBox1
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
    rownum = numerRow
    resetme (rownum)
    
End Sub
Sub CancelButton_Click()
    'TextBox1.text = NullString
    Unload Me
    
End Sub
