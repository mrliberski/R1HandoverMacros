VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AssesmentAdd 
   Caption         =   "Assessments"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   14976
   OleObjectBlob   =   "AssesmentAdd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AssesmentAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public rownum


Public Function resetme(rownum)

    Debug.Print "Reset function called with rownum number: " & rownum
    If rownum < 1 Then rownum = 1

    AssesmentAdd.Label22.Caption = Sheet20.Cells(rownum, 1).value
    AssesmentAdd.TextBox1.text = Sheet20.Cells(rownum, 2).value
    AssesmentAdd.TextBox2.text = Format(Sheet20.Cells(rownum, 3).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox3.text = Format(Sheet20.Cells(rownum, 4).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox4.text = Format(Sheet20.Cells(rownum, 5).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox5.text = Format(Sheet20.Cells(rownum, 6).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox6.text = Format(Sheet20.Cells(rownum, 7).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox7.text = Format(Sheet20.Cells(rownum, 8).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox8.text = Format(Sheet20.Cells(rownum, 9).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox9.text = Format(Sheet20.Cells(rownum, 10).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox10.text = Format(Sheet20.Cells(rownum, 11).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox11.text = Format(Sheet20.Cells(rownum, 12).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox12.text = Format(Sheet20.Cells(rownum, 13).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox13.text = Format(Sheet20.Cells(rownum, 14).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox14.text = Format(Sheet20.Cells(rownum, 15).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox15.text = Format(Sheet20.Cells(rownum, 16).value, "dd/mm/yyyy")
    AssesmentAdd.TextBox16.text = Sheet20.Cells(rownum, 17).value
    ComboBox1.text = Sheet20.Cells(rownum, 18).value
    ComboBox2.text = Sheet20.Cells(rownum, 19).value
    
    
    'PackagingBrowse.TextBox10.text = Format(Sheet18.Cells(rownum, 6).value, "hh:mm"
    
    
    AssesmentAdd.Repaint
    
End Function



Private Sub isAvailable(ByVal adres As String)
    If Dir(adres) <> NullString Then
        Label20.Caption = "Database availability Status: Available"
    Else
        Label20.Caption = "Database availability Status: Unavailable"
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
        
           
    If CheckBox1.value = True Then
        SqlString = SqlString & "WHERE ([Site] = 'RED1' OR [Site] = 'RED2' OR [Site] = 'DRO' OR [Site] = 'ALL') "
    End If
    If CheckBox2.value = True Then
        SqlString = SqlString & "WHERE ([Site] = 'RED1' OR [Site] = 'ALL') "
    End If
    If CheckBox3.value = True Then
        SqlString = SqlString & "WHERE ([Site] = 'RED2' OR [Site] = 'ALL') "
    End If
    If CheckBox4.value = True Then
        SqlString = SqlString & "WHERE ([Site] = 'DRO' OR [Site] = 'ALL') "
    End If
    If CheckBox5.value = True Then
        SqlString = SqlString & "WHERE ([Site] = 'LEFT') "
    End If
        
        
        
    'If CheckBox2.value = True Then
        'SqlString = SqlString & "WHERE ([Site] = [AdvisedQty] OR [Site] = [AdvisedQty]) "
    'End If

    SqlString = SqlString & "ORDER BY [Names]"
    '==============================================

    Call connectExecute(SqlString)
    Call paintList
    AssesmentAdd.Repaint




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
Set ws = Sheets("Assesments")

'~~> Set you relevant range here
Set rng = ws.Range("A1:S" & ws.Range("A" & ws.Rows.Count).End(xlUp).Row)

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
    .ColumnWidths = "0;100;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;30;30"
    .TopIndex = 0
End With

End Sub

Private Sub CheckBox1_Change()
    'All Sites
    If CheckBox1.value = True Then
        CheckBox2.value = False
        CheckBox3.value = False
        CheckBox4.value = False
        CheckBox5.value = False
    End If
    
End Sub
Private Sub CheckBox2_Change()
    'R1
    If CheckBox2.value = True Then
        CheckBox1.value = False
        CheckBox3.value = False
        CheckBox4.value = False
        CheckBox5.value = False
    End If
End Sub
Private Sub CheckBox3_Change()
    'R2
    If CheckBox3.value = True Then
        CheckBox2.value = False
        CheckBox1.value = False
        CheckBox4.value = False
        CheckBox5.value = False
    End If
End Sub
Private Sub CheckBox4_Change()
    'DRO
    If CheckBox4.value = True Then
        CheckBox2.value = False
        CheckBox3.value = False
        CheckBox1.value = False
        CheckBox5.value = False
    End If
End Sub
Private Sub CheckBox5_Change()
    'Leavers
    If CheckBox5.value = True Then
        CheckBox2.value = False
        CheckBox3.value = False
        CheckBox4.value = False
        CheckBox1.value = False
    End If
End Sub

Private Sub CommandButton1_Click()

    On Error Resume Next
    Debug.Print "loading adding page"
    Unload CalendarForm
    Unload Me
    
    AssAdd.Show

End Sub

Private Sub CommandButton2_Click()
    If rownum > 1 Then
        rownum = rownum - 1
        resetme (rownum)
        Debug.Print "going down " & rownum
    End If
End Sub

Private Sub CommandButton3_Click()
    If Sheet20.Cells(rownum + 1, 1).value <> "" Then
        rownum = rownum + 1
        resetme (rownum)
        Debug.Print "going up " & rownum
    End If
End Sub

Private Sub countRows() 'find last row and column

    Dim lastRowNum As Long
    Dim lastColumnNum As Long
    
    lastRowNum = Sheet20.Cells(Rows.Count, 1).End(xlUp).Row
    lastColumnNum = Sheet20.Cells(1, Columns.Count).End(xlToLeft).Column
    
End Sub

Private Sub CommandButton4_Click()
    Call isAvailable("J:\Pub-LOGISTICS\Packaging\Packaging.accdb")
        If Label20.Caption = "Database availability Status: Available" Then Call refreshFeed
    Call paintList
    rownum = 1
    resetme (rownum)
End Sub

Private Sub CommandButton5_Click()
'////edit button click
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
    rownum = numerRow
    resetme (rownum)
End Sub

Private Sub CommandButton6_Click()

If Label22.Caption = NullString Then Exit Sub

answer = MsgBox("Selected record will now be updated, do you wish to continue?", vbOKCancel, "Confirmation required")
If answer = vbCancel Then Exit Sub



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
Sheet20.Cells(rownum, 2).value = TextBox1.text
Sheet20.Cells(rownum, 3).value = DateValue(TextBox2.text)
Sheet20.Cells(rownum, 4).value = DateValue(TextBox3.text)
Sheet20.Cells(rownum, 5).value = DateValue(TextBox4.text)
Sheet20.Cells(rownum, 6).value = DateValue(TextBox5.text)
Sheet20.Cells(rownum, 7).value = DateValue(TextBox6.text)
Sheet20.Cells(rownum, 8).value = DateValue(TextBox7.text)
Sheet20.Cells(rownum, 9).value = DateValue(TextBox8.text)
Sheet20.Cells(rownum, 10).value = DateValue(TextBox9.text)
Sheet20.Cells(rownum, 11).value = DateValue(TextBox10.text)
Sheet20.Cells(rownum, 12).value = DateValue(TextBox11.text)
Sheet20.Cells(rownum, 13).value = DateValue(TextBox12.text)
Sheet20.Cells(rownum, 14).value = DateValue(TextBox13.text)
Sheet20.Cells(rownum, 15).value = DateValue(TextBox14.text)
Sheet20.Cells(rownum, 16).value = DateValue(TextBox15.text)
Sheet20.Cells(rownum, 17).value = TextBox16.text
Sheet20.Cells(rownum, 18).value = ComboBox1.text
Sheet20.Cells(rownum, 19).value = ComboBox2.text
'=====================================
 
'MsgBox "Database was updated.", vbInformation, "Procedure completed."

'=====================================

Call paintList
resetme (rownum)
Me.Repaint

End Sub

Private Sub CommandButton7_Click()
Unload AssesmentAdd


'HTML matrix
'Define your variables.
   Dim iRow As Long
   Dim iStage As Integer
   Dim iCounter As Integer
   Dim iPage As Integer
   
   'Create an .htm file in the same directory as your active workbook.
   Dim sFile As String
   sFile = ActiveWorkbook.Path & "\matrix.html"
   'sFile = ActiveWorkbook.Path & "\test.pdf"
   Close
   
Dim lastRowNum As Long
Dim lastColumnNum As Long
    
lastRowNum = Sheet20.Cells(Rows.Count, 1).End(xlUp).Row
lastColumnNum = Sheet20.Cells(1, Columns.Count).End(xlToLeft).Column
   
   'Open up the temp HTML file and format the header.
   Open sFile For Output As #1
   Print #1, "<html>"
   Print #1, "<head><title>Lifting Equipment Training Matrix</title><style>table, th, td {border: 1px solid #3d3d40;border-collapse: collapse;text-align: center;}</style>"
   Print #1, "<style type=""text/css""  .blue {   background: blue; } table {border-color: #3d3d40;}>"
   Print #1, "  body { color: #3d3d40; font-size:12px;font-family:calibri } "
   Print #1, "</style>"
   Print #1, "</head>"
   Print #1, "<body>"
   
'<thead>BORDERCOLOR="#0000FF" BORDERCOLORLIGHT="#33CCFF" BORDERCOLORDARK="#0000CC"
'<tr><th>Movie</th><th>Downloads</th><th>Grosses</th></tr>
'</thead>


Print #1, "<h1>Lifting Equipment Assessment Training Matrix</h1>"
Print #1, "<h3>Generated on " & Now() & "</h3><hr>"
Print #1, "<table border='1' BORDERCOLOR='#3d3d40'><tr><thead><th>Name & Surname</th><th>C/Balance B1</th>"
Print #1, "<th>C/Balance B2</th>"
Print #1, "<th>&nbsp;PPT A1&nbsp;</th>"
Print #1, "<th>&nbsp;PPT A2&nbsp;</th>"
Print #1, "<th>&nbsp;Tow Train H1&nbsp;</th>"
Print #1, "<th>&nbsp;VNA F1&nbsp;</th>"
Print #1, "<th>&nbsp;Bendi P1&nbsp;</th>"
Print #1, "<th>&nbsp;MEWPS 3A&nbsp;</th>"
Print #1, "<th>&nbsp;MEWPS 3B&nbsp;</th>"
Print #1, "<th>&nbsp;Stacker A4&nbsp;</th>"
Print #1, "<th>&nbsp;Stacker A5&nbsp;</th>"
Print #1, "<th>&nbsp;Reach D1&nbsp;</th>"
Print #1, "<th>&nbsp;Remote&nbsp;</th>"
Print #1, "<th>&nbsp;Assessment&nbsp;</th>"
Print #1, "<th>&nbsp;&nbsp;SITE&nbsp;&nbsp;&nbsp;</th>"
Print #1, "<th>&nbsp;&nbsp;SHIFT&nbsp;&nbsp;&nbsp;</th>"
Print #1, "</thead></tr>"




For i = 1 To lastRowNum
    Print #1, "<tr><td>" & Sheet20.Cells(i, 2).value & "</td>"
    
        'B1
        If Sheet20.Cells(i, 3).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 3).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 3).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 3).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 3).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 3).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
          
        
        'B2
        If Sheet20.Cells(i, 4).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 4).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 4).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 4).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 4).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 4).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
        
        
        
        
        
        'A1
        If Sheet20.Cells(i, 5).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 5).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 5).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 5).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 5).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 5).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
        
        
        
        
        'A2
        If Sheet20.Cells(i, 6).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 6).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 6).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 6).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 6).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 6).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
        
 
        'H1
        If Sheet20.Cells(i, 7).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 7).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 7).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 7).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 7).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 7).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
           
           
        'F1
        If Sheet20.Cells(i, 8).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 8).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 8).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 8).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 8).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 8).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
           
        'P1
        If Sheet20.Cells(i, 9).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 9).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 9).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 9).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 9).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 9).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
           
           
        '3A
        If Sheet20.Cells(i, 10).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 10).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 10).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 10).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 10).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 10).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
           
           
        '3B
        If Sheet20.Cells(i, 11).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 11).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 11).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 11).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 11).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 11).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
           
        'A4
        If Sheet20.Cells(i, 12).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 12).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 12).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 12).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 12).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 12).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
           
        'A5
        If Sheet20.Cells(i, 13).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 13).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 13).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 13).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 13).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 13).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
        
        'D1
        If Sheet20.Cells(i, 14).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 14).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 14).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 14).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 14).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 14).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
        
        'Remote
        If Sheet20.Cells(i, 15).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 15).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 15).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 15).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 15).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 15).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
        
        'Assesment
        If Sheet20.Cells(i, 16).value <> NullString Then
            If 1065 - Int(Now - DateValue(Sheet20.Cells(i, 16).value)) < 31 Then
                Print #1, "<td style='background-color:#ff5117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 16).value)) & "</td>"
            ElseIf 1065 - Int(Now - DateValue(Sheet20.Cells(i, 16).value)) > 350 Then
                Print #1, "<td style='background-color:#6aff00'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 16).value)) & "</td>"
            Else
                Print #1, "<td style='background-color:#ffc117'>" & 1065 - Int(Now - DateValue(Sheet20.Cells(i, 16).value)) & "</td>"
            End If
        Else
            Print #1, "<td></td>"
        End If
        
        'Site
        If Sheet20.Cells(i, 18).value <> NullString Then
            Print #1, "<td>" & Sheet20.Cells(i, 18).value & "</td>"
        Else
            Print #1, "<td></td>"
        End If
        
        'Shift
        If Sheet20.Cells(i, 19).value <> NullString Then
            Print #1, "<td>" & Sheet20.Cells(i, 19).value & "</td>"
        Else
            Print #1, "<td></td>"
        End If
        
        
           
    Print #1, "</tr>"
Next i

 
 
 
 
 
 
 
   Print #1, "</td></table><hr>Made by Pawel Liberski © 2020</body></html>"
   Close
   'Shell "hh " & vbLf & sFile, vbMaximizedFocus
ThisWorkbook.FollowHyperlink (sFile)

End Sub

Private Sub CommandButton8_Click()
    On Error Resume Next
    Unload CalendarForm
    Unload Me
End Sub

Private Sub Label21_Click()

End Sub

Private Sub ListBox1_Click()
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
        
        Debug.Print "Numerrow is " & numerRow
        numerRow = numerRow + 1
    
    Next intCurrentRow
     
    Set ctlSource = Nothing

    '==============================
    rownum = numerRow
    resetme (rownum)
    
End Sub



Private Sub TextBox2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox2.text = dateVariable
    'Unload CalendarForm
    
End Sub
Private Sub TextBox3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox3.text = dateVariable
    'Unload CalendarForm
End Sub
Private Sub TextBox4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox4.text = dateVariable
    'Unload CalendarForm
End Sub
Private Sub TextBox5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox5.text = dateVariable
    'Unload CalendarForm
End Sub
Private Sub TextBox6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox6.text = dateVariable
    'Unload CalendarForm
End Sub
Private Sub TextBox7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox7.text = dateVariable
    'Unload CalendarForm
End Sub
Private Sub TextBox8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox8.text = dateVariable
    'Unload CalendarForm
End Sub
Private Sub TextBox9_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox9.text = dateVariable
    'Unload CalendarForm
End Sub
Private Sub TextBox10_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox10.text = dateVariable
    'Unload CalendarForm
End Sub
Private Sub TextBox11_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox11.text = dateVariable
    'Unload CalendarForm
End Sub
Private Sub TextBox12_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox12.text = dateVariable
    'Unload CalendarForm
End Sub
Private Sub TextBox13_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox13.text = dateVariable
    'Unload CalendarForm
End Sub
Private Sub TextBox14_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox14.text = dateVariable
    'Unload CalendarForm
End Sub
Private Sub TextBox15_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then TextBox15.text = dateVariable
    'Unload CalendarForm
End Sub

Private Sub UserForm_Activate()
    
    Call isAvailable("J:\Pub-LOGISTICS\Packaging\Packaging.accdb")
        If Label20.Caption = "Database availability Status: Available" Then Call refreshFeed
    Call paintList
    Debug.Print "Form activated"
    
    
    
    If rownum < 1 Then rownum = 1
    resetme (rownum)
    
End Sub



Private Sub UserForm_Initialize()

    ComboBox1.AddItem "RED1"
    ComboBox1.AddItem "RED2"
    ComboBox1.AddItem "DRO"
    ComboBox1.AddItem "ALL"
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

