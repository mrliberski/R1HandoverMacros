VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DSReport 
   Caption         =   "BMW Despatch Report"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6588
   OleObjectBlob   =   "DSReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DSReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Function GetSignature(FPath As String) As String
    Dim FSO As Object
    Dim TSet As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TSet = FSO.GetFile(FPath).OpenAsTextStream(1, -2)
    GetSignature = TSet.ReadAll
    TSet.Close
End Function
Private Sub CommandButton1_Click()
    '# Confirm Report sending
    answer = MsgBox("Report will be send out to all recipients now. Are you sure you want to proceed?", vbOKCancel, "Please confirm")
        If answer = vbOK Then
            Call SendReport
            'MsgBox "Sending Report"
        Else: Exit Sub
        End If

End Sub

Private Sub SendReport()

 Call AddToDB
    '# Email report
    'sort out blank fields and set them to zero, otherwise it crashes
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.name Like "Text*" Then
            If ctrl = NullString Then ctrl = "Not stated"
        End If
    Next
                       
 '////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Dim OlApp As Object
    Dim NewMail As Object
    Dim EmailBody As String
    Dim StrSignature As String
    Dim sPath As String
    Dim strString
    Dim strResult
    Dim seqLead As Integer
    Dim Address As String

    Address = Application.ActiveWorkbook.FullName
    
    Set OlApp = CreateObject("Outlook.Application")
    Set NewMail = OlApp.CreateItem(0)

    With NewMail
        .display
    End With
    
Signature = NewMail.htmlbody
    
EmailBody = "<BODY style=font-size:10pt;font-family:Calibri; font-color :#3d3d40;><h3>" & ComboBox1.value & " Despatch Report:</h3>"

Dim newBody

newBody = " <html><head><style>body {color: #3d3d40;font-size:10pt;font-family:Calibri;}; <style>table, th, td {border: 1px solid #3d3d40;border-collapse: collapse;text-align: center;}</style></style></head><body>"
newBody = newBody + "<h3>" & ComboBox1.value & " Despatch Report:</h3>"


'newBody = newBody & "<hr>"


newBody = newBody & "<table border=""1"" cellspacing=""0"" cellpadding=""0"" style=font-size:10pt;font-family:Calibri; border-collapse: collapse; text-align:center;>"
newBody = newBody & "<tr><b>"

newBody = newBody & "<td align=""center"">&nbsp;Data&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Details&nbsp;</td></b></tr>"

'/////////
newBody = newBody & "<tr>"
newBody = newBody & "<td>&nbsp;Vehicle Registration:&nbsp;</td>"
newBody = newBody & "<td>&nbsp;" & TextBox1.value & "&nbsp;</td>"
newBody = newBody & "</tr>"

newBody = newBody & "<tr>"
newBody = newBody & "<td>&nbsp;Trailer Registration:&nbsp;</td>"
newBody = newBody & "<td>&nbsp;" & TextBox2.value & "&nbsp;</td>"
newBody = newBody & "</tr>"

newBody = newBody & "<tr>"
newBody = newBody & "<td>&nbsp;Planned Arrival Time:&nbsp;</td>"
newBody = newBody & "<td>&nbsp;" & TextBox3.value & "&nbsp;</td>"
newBody = newBody & "</tr>"

newBody = newBody & "<tr>"
newBody = newBody & "<td>&nbsp;Actual Arrival Time:&nbsp;</td>"
newBody = newBody & "<td>&nbsp;" & TextBox4.value & "&nbsp;</td>"
newBody = newBody & "</tr>"

newBody = newBody & "<tr>"
newBody = newBody & "<td>&nbsp;Empties delivered:&nbsp;</td>"
newBody = newBody & "<td>&nbsp;" & TextBox5.value & "&nbsp;</td>"
newBody = newBody & "</tr>"

newBody = newBody & "<tr>"
newBody = newBody & "<td>&nbsp;Pallets Mixed:&nbsp;</td>"
newBody = newBody & "<td>&nbsp;" & TextBox6.value & "&nbsp;</td>"
newBody = newBody & "</tr>"

newBody = newBody & "<tr>"
newBody = newBody & "<td>&nbsp;Pallets Shipped:&nbsp;</td>"
newBody = newBody & "<td>&nbsp;" & TextBox7.value & "&nbsp;</td>"
newBody = newBody & "</tr>"

newBody = newBody & "<tr>"
newBody = newBody & "<td>&nbsp;Loading Finish Time:&nbsp;</td>"
newBody = newBody & "<td>&nbsp;" & TextBox8.value & "&nbsp;</td>"
newBody = newBody & "</tr>"

newBody = newBody & "<tr>"
newBody = newBody & "<td>&nbsp;Vehicle Departure Time:&nbsp;</td>"
newBody = newBody & "<td>&nbsp;" & TextBox9.value & "&nbsp;</td>"
newBody = newBody & "</tr>"

newBody = newBody & "<tr>"
newBody = newBody & "<td>&nbsp;Comments and Observations:&nbsp;</td>"
newBody = newBody & "<td>&nbsp;" & TextBox10.value & "&nbsp;</td>"
newBody = newBody & "</tr></table>"





newBody = newBody & "<br>Despatch report version 01/06/2021 - generated on " & Now() & " by " & Environ("username")
newBody = newBody & "</html>"

sPath = "Y:\Application Data\Microsoft\Signatures\Main.htm"
'sPath = "C:\Users\Liberski, Pawel\AppData\Roaming\Microsoft\Signatures\main.htm"
If Dir(sPath) <> "" Then
    StrSignature = GetSignature(sPath)
Else
    StrSignature = ""
End If

'On Error Resume Next

Application.ScreenUpdating = False

    With NewMail
        '.To = "liberski, pawel"
        .To = "pugh, richard; celec, marek; blake, yuell; hussain, ishtiak; kaczynska, daria;"
        .cc = "bennett, chris; wood, craig; Liberski, Pawel; partridge, jamie; sim, alina; kaczynska, daria; masson, darren; oliver, yvonne; wootton, louise; leighton, rebecca; cendrowska, alexandra; southall, paula; "
        .BCC = ""
        .Subject = ComboBox1.value & " Despatch Report"
        '.htmlbody = EmailBody & "</BODY>" & vbNewLine & Signature
        .htmlbody = newBody & "</html></BODY>" '& vbNewLine & Signature
        '.Attachments.Add (Address)
        .send
    End With
    
    'On Error GoTo 0
    Set NewMail = Nothing
    Set OlApp = Nothing
    
    'Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.name Like "Text*" Then ctrl = NullString
    Next
 Application.ScreenUpdating = True
 

 
DSReport.Hide
    
End Sub

Private Sub CommandButton2_Click()
    '# Close Form
    DSReport.Hide
    
End Sub

Private Sub CommandButton3_Click()
    '#Reset All Fields
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.name Like "Text*" Then ctrl = NullString
    Next
    
    Me.Repaint
    
End Sub
Private Sub UserForm_Initialize()
    '#Populate combo box
    ComboBox1.AddItem "RQP"
    ComboBox1.AddItem "STP"
    ComboBox1.AddItem "IP"
    ComboBox1.AddItem "Doors"
    ComboBox1.AddItem "DSR1"
    ComboBox1.AddItem "DSR2"
    
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.name Like "Text*" Then ctrl = NullString
    Next
    
End Sub
Private Sub AddToDB()

If TextBox1.text = NullString Then Exit Sub


'========================== VALIDATE DATA IN TEXTBOX or COMBO
If ComboBox1.text = NullString Then
    MsgBox "You must specify shipment type."
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
    
    
            
    sSQL = "SELECT * FROM Deliveries"     'sql query - select * from selected database
    
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
                
                rst(1).value = DateValue(Date)
                rst(2).value = TimeValue(Now)
                rst(3).value = UCase(Environ("username"))
                rst(4).value = ComboBox1.value
                rst(5).value = TextBox1.value 'vehicle
                rst(6).value = TextBox2.value 'trailer
                rst(7).value = TextBox3.value 'planned
                rst(8).value = TextBox4.value 'actual
                rst(9).value = TextBox8.value 'finish
                rst(10).value = TextBox9.value 'departure
                rst(11).value = TextBox5.value 'delivered
                rst(12).value = TextBox7.value 'shipped
                rst(13).value = TextBox6.value 'mixed
                rst(14).value = TextBox10.value 'observation

                'If TextBox2.value <> NullString Then rst(2) = DateValue(TextBox2.value) Else rst(2) = NullString
                'If TextBox3.value <> NullString Then rst(3) = DateValue(TextBox3.value) Else rst(3) = NullString
                'If TextBox4.value <> NullString Then rst(4) = DateValue(TextBox4.value) Else rst(4) = NullString            'now we are adding data to it
                'If TextBox5.value <> NullString Then rst(5) = DateValue(TextBox5.value) Else rst(5) = NullString          'from the booking form
                'If TextBox6.value <> NullString Then rst(6) = DateValue(TextBox6.value) Else rst(6) = NullString
                'If TextBox7.value <> NullString Then rst(7) = DateValue(TextBox7.value) Else rst(7) = NullString
                'If TextBox8.value <> NullString Then rst(8) = DateValue(TextBox8.value) Else rst(8) = NullString
                'If TextBox9.value <> NullString Then rst(9) = DateValue(TextBox9.value) Else rst(9) = NullString
                'If TextBox10.value <> NullString Then rst(10) = DateValue(TextBox10.value) Else rst(10) = NullString
                'If TextBox11.value <> NullString Then rst(11) = DateValue(TextBox11.value) Else rst(11) = NullString
                'If TextBox12.value <> NullString Then rst(12) = DateValue(TextBox12.value) Else rst(12) = NullString
                'If TextBox13.value <> NullString Then rst(13) = DateValue(TextBox13.value) Else rst(13) = NullString
                'If TextBox14.value <> NullString Then rst(14) = DateValue(TextBox14.value) Else rst(14) = NullString
                'If TextBox15.value <> NullString Then rst(15) = DateValue(TextBox15.value) Else rst(15) = NullString
                
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


End Sub
