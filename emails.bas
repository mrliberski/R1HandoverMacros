Attribute VB_Name = "emails"

'//===============================================================================================================
'// This block triggers autosave of the file every 30 mins. 
Sub Command()
    ControlPage.Show
End Sub
Sub switch()
    Application.OnTime Now + TimeValue("00:30:00"), "autoupdate"
End Sub
Sub autoupdate()
    On Error Resume Next
    'Sheets("Timesheet").Visible = xlSheetVeryHidden
    ThisWorkbook.Activate
    ThisWorkbook.Save
    Sheets("BRIEF").Unprotect
    Sheets("BRIEF").Cells(56, 2).value = "Last autosaved on: " & Date & " at " & Time
    Sheets("BRIEF").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True
    Call Silent_handover
    Call switch

End Sub
'//===============================================================================================================

Sub send_handover()
'// It will create outlook email with shift report and will attach itself to it

Sheets("Timesheet").Visible = xlSheetVeryHidden

ThisWorkbook.Save

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
    
EmailBody = "<BODY style=font-size:10pt;font-family:Calibri; font-color :#3d3d40;>Hi All<br><br>Please find attached shift's report"
Dim newBody

newBody = " <html><head><style>body {color: #3d3d40;font-size:10pt;font-family:Calibri;}</style></head><body>"

newBody = newBody + "<h3>R1 Logistics Shift Handover&nbsp;" & Cells(3, 19).value & "</h3>"


If Cells(5, 4).value <> NullString Then newBody = newBody + "<b>H&S Issues -&nbsp;</b> " & Cells(5, 4).value
newBody = newBody + "<hr>"

If Cells(7, 6).value <> NullString Then newBody = newBody + "<b>FLT Reported Issues -&nbsp;</b> " & Cells(7, 6).value & "<br>"
If Cells(8, 6).value <> NullString Then newBody = newBody + "<b>FLT Issues to be reported - &nbsp;</b> " & Cells(8, 6).value
newBody = newBody + "<hr>"

If Cells(7, 17).value <> NullString Then newBody = newBody + "<b>General Issues -&nbsp;</b> " & Cells(7, 17).value & "<br>"
If Cells(51, 4).value <> NullString Then newBody = newBody + "<b>Recycling Info -&nbsp;</b>" & Cells(51, 4).value & "<br>"
If Cells(53, 4).value <> NullString Then newBody = newBody + "<b>Other Issues -&nbsp;</b>" & Cells(53, 4).value

newBody = newBody + "<hr><h4>IMM</h4>"

If Cells(11, 4).value <> NullString Then newBody = newBody + "<b>5S Audit -&nbsp;</b>" & Cells(11, 4).value & "<br>"
If Cells(13, 6).value <> NullString Then newBody = newBody + "<b>Zero Packaging -&nbsp;</b> " & Cells(13, 6).value & "<br>"
If Cells(14, 4).value <> NullString Then newBody = newBody + "<b>Production / Plan -&nbsp;</b> " & Cells(14, 4).value & "<br>"
If Cells(17, 4).value <> NullString Then newBody = newBody + "<b>Part Shortages -&nbsp;</b> " & Cells(17, 4).value & "<br>"
If Cells(19, 4).value <> NullString Then newBody = newBody + "<b>Quality / SQA -&nbsp;</b> " & Cells(19, 4).value & "<br>"
If Cells(21, 4).value <> NullString Then newBody = newBody + "<b>Aftermarket / Reorders -&nbsp;</b> " & Cells(21, 4).value & "<br>"
If Cells(23, 4).value <> NullString Then newBody = newBody + "<b>Other -&nbsp;</b> " & Cells(23, 4).value & "<br>"
If Cells(27, 4).value <> NullString Then newBody = newBody + "<b>Downtime reported -&nbsp;</b> " & Cells(27, 4).value

newBody = newBody + "<hr><h4>F54 / F5X IP</h4>"

If Cells(11, 9).value <> NullString Then newBody = newBody + "<b>5S Audit -&nbsp;</b>" & Cells(11, 9).value & "<br>"
If Cells(13, 12).value <> NullString Then newBody = newBody + "<b>Loads completed -&nbsp;</b> " & Cells(13, 12).value & "<br>"
If Cells(14, 9).value <> NullString Then newBody = newBody + "<b>Production / Plan -&nbsp;</b> " & Cells(14, 9).value & "<br>"
If Cells(17, 9).value <> NullString Then newBody = newBody + "<b>Part Shortages -&nbsp;</b> " & Cells(17, 9).value & "<br>"
If Cells(19, 9).value <> NullString Then newBody = newBody + "<b>Quality / SQA -&nbsp;</b> " & Cells(19, 9).value & "<br>"
If Cells(21, 9).value <> NullString Then newBody = newBody + "<b>Aftermarket / Reo -&nbsp;</b> " & Cells(21, 9).value & "<br>"
If Cells(23, 9).value <> NullString Then newBody = newBody + "<b>Other -&nbsp;</b> " & Cells(23, 9).value & "<br>"
If Cells(27, 9).value <> NullString Then newBody = newBody + "<b>Downtime reported -&nbsp;</b> " & Cells(27, 9).value

newBody = newBody + "<hr><h4>F54 / F5X GLOVEBOX</h4>"

If Cells(11, 13).value <> NullString Then newBody = newBody + "<b>5S Audit -&nbsp;</b>" & Cells(11, 13).value & "<br>"
If Cells(13, 14).value <> NullString Then newBody = newBody + "<b>F5X GB Plan -&nbsp;</b> " & Cells(13, 14).value & "<br>"
If Cells(14, 14).value <> NullString Then newBody = newBody + "<b>F54 GB Plan -&nbsp;</b> " & Cells(14, 14).value & "<br>"
If Cells(15, 13).value <> NullString Then newBody = newBody + "<b>Production / Packaging -&nbsp;</b> " & Cells(15, 13).value & "<br>"
If Cells(17, 13).value <> NullString Then newBody = newBody + "<b>Part Shortages -&nbsp;</b> " & Cells(17, 13).value & "<br>"
If Cells(19, 13).value <> NullString Then newBody = newBody + "<b>Quality / SQA -&nbsp;</b> " & Cells(19, 13).value & "<br>"
If Cells(21, 13).value <> NullString Then newBody = newBody + "<b>Aftermarket / Reo -&nbsp;</b> " & Cells(21, 13).value & "<br>"
If Cells(23, 13).value <> NullString Then newBody = newBody + "<b>Other -&nbsp;</b> " & Cells(23, 13).value & "<br>"
If Cells(27, 13).value <> NullString Then newBody = newBody + "<b>Downtime reported -&nbsp;</b> " & Cells(27, 13).value

newBody = newBody + "<hr><h4>STORES</h4>"

If Cells(11, 17).value <> NullString Then newBody = newBody + "<b>5S Audit -&nbsp;</b>" & Cells(11, 17).value & "<br>"
If Cells(13, 17).value <> NullString Then newBody = newBody + "<b>Packaging / Racking -&nbsp;</b> " & Cells(13, 17).value & "<br>"
If Cells(15, 17).value <> NullString Then newBody = newBody + "<b>General -&nbsp;</b> " & Cells(14, 17).value & " " & Cells(15, 17).value & "<br>"
If Cells(17, 17).value <> NullString Then newBody = newBody + "<b>Part Shortages -&nbsp;</b> " & Cells(17, 17).value & "<br>"
If Cells(19, 17).value <> NullString Then newBody = newBody + "<b>Quality / SQA -&nbsp;</b> " & Cells(19, 17).value & "<br>"
If Cells(21, 17).value <> NullString Then newBody = newBody + "<b>Aftermarket / Special -&nbsp;</b> " & Cells(21, 17).value & "<br>"
If Cells(23, 17).value <> NullString Then newBody = newBody + "<b>Other -&nbsp;</b> " & Cells(23, 17).value & "<br>"
If Cells(27, 17).value <> NullString Then newBody = newBody + "<b>Downtime reported -&nbsp;</b> " & Cells(27, 17).value

newBody = newBody + "<hr><h4>ZONE 3</h4>"

If Cells(39, 18).value <> NullString Then newBody = newBody + "<b>SILO Deliveries -&nbsp;</b> " & Cells(39, 18).value & "<br>"
If Cells(40, 17).value <> NullString Then newBody = newBody + Cells(40, 17).value

newBody = newBody + "<hr><h4>ZONE 4</h4>"

If Cells(31, 18).value <> NullString Then newBody = newBody + "<b>Reorders -&nbsp;</b> " & Cells(31, 18).value & "<br>"
If Cells(31, 12).value <> NullString Then newBody = newBody + Cells(31, 12).value & "<br>"
If Cells(32, 17).value <> NullString Then newBody = newBody + Cells(32, 17).value

newBody = newBody + "<hr><h4>NED CAR</h4>"

If Cells(48, 11).value <> NullString Then newBody = newBody + "<b>Finished loads -&nbsp;</b> " & Cells(48, 11).value & "<br>"
If Cells(47, 6).value <> NullString Then newBody = newBody + "<b>Collections -&nbsp;</b>" & Cells(47, 6).value & "<br>"
If Cells(48, 6).value <> NullString Then newBody = newBody + Cells(48, 6).value

newBody = newBody + "<hr><h4>DIRECT SUPPLY</h4>"

If Cells(49, 4).value <> NullString Then newBody = newBody + "<b>DS Issues -&nbsp;</b> " & Cells(49, 4).value & "<br>"
If Cells(47, 14).value <> NullString Then newBody = newBody + "<b>DS NED Car Collections -&nbsp;</b>" & Cells(47, 14).value & "<hr></body>"


sPath = "Y:\Application Data\Microsoft\Signatures\Main.htm"

If Dir(sPath) <> "" Then
    StrSignature = GetSignature(sPath)
Else
    StrSignature = ""
End If

For X = 71 To 91
    If Cells(X, 2).value <> "" Then
        adressee = adressee & Cells(X, 2).value & "@grupoantolin.com; "
    End If
Next X

On Error Resume Next

    With NewMail
        .To = adressee
        .cc = ""
        .BCC = ""
        .Subject = "Logistic Shift Report " & Cells(3, 19).value
        '.htmlbody = EmailBody & "</BODY>" & vbNewLine & Signature
        .htmlbody = newBody & "</BODY>" & vbNewLine & Signature
        .Attachments.Add (Address)
        .display
    End With
    
    'On Error GoTo 0
    Set NewMail = Nothing
    Set OlApp = Nothing

End Sub
'//===============================================================================================================

Sub polymer_request(harvest As String)
'// will send out email request basing on information from th eform 

ThisWorkbook.Save

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
    
EmailBody = "<BODY style=font-size:10pt;font-family:Calibri>Dear All<br><br>Please send the polymer stated below to Redditch1 :"
EmailBody = EmailBody & "<hr>"

EmailBody = EmailBody & harvest

EmailBody = EmailBody & "<hr>" & "Request generated on " & Now()
sPath = "Y:\Application Data\Microsoft\Signatures\Main.htm"

If Dir(sPath) <> "" Then
    StrSignature = GetSignature(sPath)
Else
    StrSignature = ""
End If

For X = 71 To 91 ' sciaganie adresow
    If Cells(X, 10).value <> "" Then
        adressee = adressee & Cells(X, 10).value & "@grupoantolin.com; "
    End If
Next X

On Error Resume Next

    With NewMail
        .To = adressee
        .To = ""
        .cc = ""
        .BCC = ""
        .Subject = "R1 Polymer request "
        .htmlbody = EmailBody & "</BODY>" & vbNewLine & Signature
        '.display
        .send
    End With
    
    'On Error GoTo 0
    Set NewMail = Nothing
    Set OlApp = Nothing
    MsgBox "Request has been made.", vbOKOnly, "Message sent"
ThisWorkbook.Save
End Sub
'//===============================================================================================================

Function GetSignature(FPath As String) As String
    Dim FSO As Object
    Dim TSet As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TSet = FSO.GetFile(FPath).OpenAsTextStream(1, -2)
    GetSignature = TSet.ReadAll
    TSet.Close
End Function
'//===============================================================================================================

'// Redundant
Sub L551_despatchReport(ByVal arrivalTime As String, ByVal leaveTime As String, _
            ByVal lhFronts As Integer, ByVal rhFronts As Integer, _
            ByVal lhRears As Integer, ByVal rhRears As Integer, _
            ByVal empties As Integer)
            
        On Error GoTo 0
        
        With Application
    .ScreenUpdating = False
End With
            
        Dim OlApp As Object
        Dim NewMail As Object
        Dim EmailBody As String
        Dim StrSignature As String
        Dim sPath As String
        Dim strString
        Dim strResult
        Dim seqLead As Integer
        
        'Dim Address As String
        'Address = Application.ActiveWorkbook.FullName

        Set OlApp = CreateObject("Outlook.Application")
        Set NewMail = OlApp.CreateItem(0)
        
        'With NewMail
        '    .display
        'End With
        
        Signature = NewMail.htmlbody

        EmailBody = "<BODY style=font-size:10pt;font-family:Calibri><b>L551 DESPATCH REPORT</b><hr>"
        EmailBody = EmailBody & "Trailer arrived at: " & arrivalTime
        EmailBody = EmailBody & "<br> Trailer left site at: " & leaveTime
        EmailBody = EmailBody & "<br><br>DESPATCHED <hr>"
        EmailBody = EmailBody & "197431012-100 - LH Fronts: " & lhFronts
        EmailBody = EmailBody & "<br>197431011-100 - RH Fronts: " & rhFronts
        EmailBody = EmailBody & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        EmailBody = EmailBody & "197431022     -     LH Rears: &nbsp;&nbsp;" & lhRears
        EmailBody = EmailBody & "&nbsp;&nbsp;<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        EmailBody = EmailBody & "197431021     -     RH Rears: &nbsp;" & rhRears
        EmailBody = EmailBody & "<hr>Empties received: " & empties & "&nbsp;&nbsp;"
        EmailBody = EmailBody & "<hr>" & "Report generated on " & Now()
        
        
        sPath = "Y:\Application Data\Microsoft\Signatures\Main.htm"
        If Dir(sPath) <> "" Then
            StrSignature = GetSignature(sPath)
        Else
            StrSignature = ""
        End If

        For X = 71 To 91
            If Cells(X, 4).value <> "" Then
                adressee = adressee & Cells(X, 4).value & "@grupoantolin.com; "
            End If
        Next X

        With NewMail
        .To = adressee
        .cc = ""
        '.BCC = ""
        .Subject = "L551 Door Despatch report"
        .htmlbody = EmailBody & "</BODY>" & vbNewLine & Signature
        '.Attachments.Add (Address)
        '.display
        .send
        End With

    Set NewMail = Nothing
    Set OlApp = Nothing
    
    With Application
    .ScreenUpdating = True
End With
    
    MsgBox "Report was sent successfully", vbOKOnly, "Report was sent"
                        

End Sub
'//===============================================================================================================


Sub L551_despatch()
'// another version - now also redundant

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
    
EmailBody = "<BODY style=font-size:10pt;font-family:Calibri><b>L551 DESPATCH REPORT</b><hr>Trailer arrived at: <br> Trailer left site at: <br>"
EmailBody = EmailBody & "<br>DESPATCHED <hr>"
EmailBody = EmailBody & "197431012-100 - LH Fronts: <br>197431011-100 - RH Fronts: <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;197431022     -     LH Rears: &nbsp;&nbsp;<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;197431021     -     RH Rears: <hr>Empties received: &nbsp;&nbsp;"
EmailBody = EmailBody & "<hr>" & "Report generated on " & Now()

sPath = "Y:\Application Data\Microsoft\Signatures\Main.htm"
If Dir(sPath) <> "" Then
    StrSignature = GetSignature(sPath)
Else
    StrSignature = ""
End If

For X = 71 To 91
    If Cells(X, 4).value <> "" Then
        adressee = adressee & Cells(X, 4).value & "@grupoantolin.com; "
    End If
Next X

On Error Resume Next

    With NewMail
        .To = adressee
        .cc = ""
        .BCC = ""
        .Subject = "L551 Door Despatch report"
        .htmlbody = EmailBody & "</BODY>" & vbNewLine & Signature
        '.Attachments.Add (Address)
        .display
    End With
    'On Error GoTo 0
    Set NewMail = Nothing
    Set OlApp = Nothing

    
End Sub
'//===============================================================================================================

'// sends email with count
Sub stock_count_mail(lh56 As Integer, rh56 As Integer, wad56 As Integer, _
lh54 As Integer, rh54 As Integer, wad54 As Integer, _
lhkab As Integer, rhkab As Integer, _
ftray As Integer, evo As Integer, _
f56wind As Integer, f54wind As Integer)

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

'inserting to log ////////////////////////////////////////////////////

Dim harvest As String
Dim LastRow As Long, ws As Worksheet
Set ws = Sheets("SGI Count Log")
ws.Unprotect

'F56 LH SGI/////////////////////////////////////////
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
ws.Cells(LastRow, 1).value = Now() 'Adds timestamp of request
ws.Cells(LastRow, 4).value = Environ("username")
ws.Cells(LastRow, 2).value = "F56 LH SGI"
If lh56 = 0 Then
    ws.Cells(LastRow, 3).value = "Not Counted"
Else
    ws.Cells(LastRow, 3).value = lh56 * 14
End If
'F56 RH SGI////////////////////////////////////////
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
ws.Cells(LastRow, 1).value = Now() 'Adds timestamp of request
ws.Cells(LastRow, 4).value = Environ("username")
ws.Cells(LastRow, 2).value = "F56 RH SGI"
If rh56 = 0 Then
    ws.Cells(LastRow, 3).value = "Not Counted"
Else
    ws.Cells(LastRow, 3).value = rh56 * 14
End If
'F56 WAD////////////////////////////////////////
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
ws.Cells(LastRow, 1).value = Now() 'Adds timestamp of request
ws.Cells(LastRow, 4).value = Environ("username")
ws.Cells(LastRow, 2).value = "F56 WAD"
If wad56 = 0 Then
    ws.Cells(LastRow, 3).value = "Not Counted"
Else
    ws.Cells(LastRow, 3).value = wad56 * 30
End If
'F54 LH SGI/////////////////////////////////////////
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
ws.Cells(LastRow, 1).value = Now() 'Adds timestamp of request
ws.Cells(LastRow, 4).value = Environ("username")
ws.Cells(LastRow, 2).value = "F54 LH SGI"
If lh54 = 0 Then
    ws.Cells(LastRow, 3).value = "Not Counted"
Else
    ws.Cells(LastRow, 3).value = lh54 * 12
End If
'F54 RH SGI////////////////////////////////////////
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
ws.Cells(LastRow, 1).value = Now() 'Adds timestamp of request
ws.Cells(LastRow, 4).value = Environ("username")
ws.Cells(LastRow, 2).value = "F54 RH SGI"
If rh54 = 0 Then
    ws.Cells(LastRow, 3).value = "Not Counted"
Else
    ws.Cells(LastRow, 3).value = rh54 * 12
End If
'F54 WAD////////////////////////////////////////
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
ws.Cells(LastRow, 1).value = Now() 'Adds timestamp of request
ws.Cells(LastRow, 4).value = Environ("username")
ws.Cells(LastRow, 2).value = "F54 WAD"
If wad54 = 0 Then
    ws.Cells(LastRow, 3).value = "Not Counted"
Else
    ws.Cells(LastRow, 3).value = wad54 * 30
End If

'LH KAB////////////////////////////////////////
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
ws.Cells(LastRow, 1).value = Now() 'Adds timestamp of request
ws.Cells(LastRow, 4).value = Environ("username")
ws.Cells(LastRow, 2).value = "F54 LH Taped KAB"
If lhkab = 0 Then
    ws.Cells(LastRow, 3).value = "Not Counted"
Else
    ws.Cells(LastRow, 3).value = lhkab * 50
End If
'RH KAB////////////////////////////////////////
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
ws.Cells(LastRow, 1).value = Now() 'Adds timestamp of request
ws.Cells(LastRow, 4).value = Environ("username")
ws.Cells(LastRow, 2).value = "F54 RH Taped KAB"
If rhkab = 0 Then
    ws.Cells(LastRow, 3).value = "Not Counted"
Else
    ws.Cells(LastRow, 3).value = rhkab * 50
End If

ws.Protect
'/////////////////////////////////////////////////////////////////////
    
EmailBody = "<BODY style=font-size:10pt;font-family:Calibri><b>R1 Stock Count</b><hr>"
EmailBody = EmailBody & "F56 LH SGI - " & lh56 * 14
EmailBody = EmailBody & "<br>F56 RH SGI - " & rh56 * 14
EmailBody = EmailBody & "<br>F56 WAD - " & wad56 * 30
EmailBody = EmailBody & "<br>F56 W/S Demist - " & f56wind
EmailBody = EmailBody & "<hr>F54 LH SGI - " & lh54 * 12
EmailBody = EmailBody & "<br>F54 RH SGI - " & rh54 * 12
EmailBody = EmailBody & "<br>F54 WAD - " & wad54 * 30
EmailBody = EmailBody & "<br>F54 W/S Demist - " & f54wind
'EmailBody = EmailBody & "<hr>L551 LH Front - " & lhf * 20
'EmailBody = EmailBody & "<br>L551 RH Front - " & rhf * 20
'EmailBody = EmailBody & "<br>L551 LH Rear - " & lhr * 20
'EmailBody = EmailBody & "<br>L551 RH Rear - " & rhr * 20
'EmailBody = EmailBody & "<br>L551 Empties - " & empties
EmailBody = EmailBody & "<hr>" & "LH Taped KAB Guide - " & lhkab * 50
EmailBody = EmailBody & "<br>RH Taped KAB Guide - " & rhkab * 50

'If sa_count = True Then
    'EmailBody = EmailBody & "<hr>F5x LH G/Box Housing - " & lh56housing * 24
    'EmailBody = EmailBody & "<br>F5x LH G/Box Lid - " & lh56lid * 70
    'EmailBody = EmailBody & "<hr>F5x RH G/Box Housing - " & rh56housing * 24
    'EmailBody = EmailBody & "<br>F5x RH G/Box Lid - " & rh56lid * 70
    
    'EmailBody = EmailBody & "<hr>F54 LH G/Box Housing - " & lh54housing * 24
    'EmailBody = EmailBody & "<br>F54 LH G/Box Inner - " & lh54inner * 70
    'EmailBody = EmailBody & "<br>F54 LH G/Box Outer - " & lh54outer * 70
    
    'EmailBody = EmailBody & "<hr>F54 RH G/Box Housing - " & rh54housing * 24
    'EmailBody = EmailBody & "<br>F54 RH G/Box Inner - " & rh54inner * 70
    'EmailBody = EmailBody & "<br>F54 RH G/Box Outer - " & rh54outer * 70
'End If
EmailBody = EmailBody & "<hr>F54 LCI F Carrier - " & ftray * 22
EmailBody = EmailBody & "<br>F5x EVO F Carrier - " & evo * 48

EmailBody = EmailBody & "<hr>" & "Report generated on " & Now()

sPath = "Y:\Application Data\Microsoft\Signatures\Main.htm"
If Dir(sPath) <> "" Then
    StrSignature = GetSignature(sPath)
Else
    StrSignature = ""
End If

On Error Resume Next

For X = 71 To 91 ' sciaganie adresow
    If Cells(X, 12).value <> "" Then
        adressee = adressee & Cells(X, 12).value & "@grupoantolin.com; "
    End If
Next X

    With NewMail
        .To = adressee
        '.To = ""
        .cc = ""
        .BCC = ""
        .Subject = "R1 SGI and stock count"
        .htmlbody = EmailBody & "</BODY>" & vbNewLine & Signature
        '.Attachments.Add (Address)
        .display
    End With
    
    'On Error GoTo 0
    Set NewMail = Nothing
    Set OlApp = Nothing

End Sub
'//===============================================================================================================

Sub delivery_check(deliverka As String)

On Error Resume Next

Dim OlApp As Object
Dim NewMail As Object
Dim EmailBody As String
Dim StrSignature As String
Dim sPath As String
Dim strString
Dim strResult
Dim seqLead As Integer
Dim Address As String

If deliverka = "" Then ' check if anything was entered
    answer = MsgBox("Delivery name was not specified.", vbCritical, "Enter delivery name")
    GoTo Notgood
End If

Application.ScreenUpdating = False
Address = Application.ActiveWorkbook.FullName
    
Set OlApp = CreateObject("Outlook.Application")
Set NewMail = OlApp.CreateItem(0)
'With NewMail
'.display
'End With
    
EmailBody = "<BODY style=font-size:10pt;font-family:Calibri><b>R1 Delivery Check and Post request</b><hr>We have received a delivery from:   <b>" & deliverka & "</b>"
EmailBody = EmailBody & "<br>All parts have been scanned and delivery note has been uploaded.<br> Please check delivery in urgently and do let us know once this has been completed."
EmailBody = EmailBody & "<hr>Request generated on " & Now()
Signature = NewMail.htmlbody

sPath = "Y:\Application Data\Microsoft\Signatures\Main.htm"
If Dir(sPath) <> "" Then
    StrSignature = GetSignature(sPath)
Else
    StrSignature = ""
End If

On Error GoTo Notgood

    With NewMail
        .To = ""
        .cc = ""
        .BCC = ""
        .Subject = "R1 Delivery Check and Post request"
        .htmlbody = EmailBody & "</BODY>" & vbNewLine & Signature
        '.Attachments.Add (Address)
        '.display
        .send
    End With
    
    'NewMail.send
    
    'On Error GoTo 0
    Set NewMail = Nothing
    Set OlApp = Nothing
    
    Sheets("Delivery Log").Unprotect
    Dim LastRow As Long, ws As Worksheet, rfnumber As String
    Set ws = Sheets("Delivery Log")
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = deliverka 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
    Sheets("Delivery Log").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True
    Application.ScreenUpdating = True
    answer = MsgBox(deliverka & " delivery check in has been requested...", vbInformation, "Check in requested.")
    GoTo Finito
    
Notgood:
Application.ScreenUpdating = True
answer = MsgBox("Something has gone wrong, couldn't confirm if request has been posted succesfully.", vbCritical, "Bugger....")

Finito:
Application.ScreenUpdating = True
ThisWorkbook.Save

End Sub
'//===============================================================================================================


Sub msgboxsample()
Dim deliverka As String
deliverka = "PJS"
answer = MsgBox(deliverka & " delivery check in has been requested...", vbInformation, "Check in has been requested.")


End Sub
'//===============================================================================================================

Sub IP_despatch()

Dim OlApp As Object
Dim NewMail As Object
Dim EmailBody As String
Dim StrSignature As String
Dim sPath As String
Dim strString
Dim strResult
Dim seqLead As Integer
Dim Address As String

Dim rng As Range
Set rng = Nothing
Sheets("BRIEF").Unprotect
Set rng = Sheets("BRIEF").Range("D31:K37").SpecialCells(xlCellTypeVisible)
If rng Is Nothing Then
    MsgBox "The selection is not a range or the sheet is protected. " & _
           vbNewLine & "Please correct and try again.", vbOKOnly
    Exit Sub
End If
With Application
    .EnableEvents = False
    .ScreenUpdating = False
End With

Address = Application.ActiveWorkbook.FullName
    
Set OlApp = CreateObject("Outlook.Application")
Set NewMail = OlApp.CreateItem(0)
With NewMail
.display
End With
Signature = NewMail.htmlbody
    
EmailBody = "<BODY style=font-size:10pt;font-family:Calibri><b>F5x IP DESPATCH REPORT</b><hr>"
EmailBody = EmailBody & RangetoHTML(rng)
EmailBody = EmailBody & "<hr>" & "Report generated on " & Now()

sPath = "Y:\Application Data\Microsoft\Signatures\Main.htm"
If Dir(sPath) <> "" Then
    StrSignature = GetSignature(sPath)
Else
    StrSignature = ""
End If


On Error Resume Next

For X = 71 To 91 ' sciaganie adresow
    If Cells(X, 8).value <> "" Then
        adressee = adressee & Cells(X, 8).value & "@grupoantolin.com; "
    End If
Next X

    With NewMail
        .To = adressee
        .cc = ""
        .BCC = ""
        .Subject = "F5x IP Despatch Report"
        .htmlbody = EmailBody & "</BODY>" & vbNewLine & Signature
        '.Attachments.Add (Address)
        .display
    End With
    
With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

Set NewMail = Nothing
Set OlApp = Nothing
Sheets("BRIEF").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True

End Sub
Sub ipnew()

With Application
    .ScreenUpdating = False
End With

Dim OlApp As Object
Dim NewMail As Object
Dim EmailBody As String
Dim StrSignature As String
Dim sPath As String
Dim strString
Dim strResult
Dim seqLead As Integer
Dim Address As String
    
Set OlApp = CreateObject("Outlook.Application")
Set NewMail = OlApp.CreateItem(0)

'With NewMail
'.display
'End With
Signature = NewMail.htmlbody

    
EmailBody = "<BODY style=font-size:10pt;font-family:Calibri><b>F5x IP DESPATCH REPORT</b><hr>"
EmailBody = EmailBody & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=font-size:10pt;font-family:Calibri; text-align:center;>"
EmailBody = EmailBody & "<tr><b><td></td>"
EmailBody = EmailBody & "<td align=""center"">&nbsp;Planned&nbsp;</td>"
EmailBody = EmailBody & "<td align=""center"">&nbsp;Arrival&nbsp;</td>"
EmailBody = EmailBody & "<td align=""center"">&nbsp;Departure&nbsp;</td>"
EmailBody = EmailBody & "<td align=""center"">&nbsp;Last Sequence&nbsp;</td>"
EmailBody = EmailBody & "<td align=""center"">&nbsp;Comments&nbsp;</td></b></tr>"

For i = 1 To 6
    EmailBody = EmailBody & "<tr><td align=""center"">" & i & ".</td>"
    EmailBody = EmailBody & "<td align=""center"">" & Format(Cells(31 + i, 5).value, "hh:mm") & "</td>"
    EmailBody = EmailBody & "<td align=""center"">" & Format(Cells(31 + i, 6).value, "hh:mm") & "</td>"
    EmailBody = EmailBody & "<td align=""center"">" & Format(Cells(31 + i, 7).value, "hh:mm") & "</td>"
    EmailBody = EmailBody & "<td align=""center"">" & Cells(31 + i, 8).value & "</td>"
    EmailBody = EmailBody & "<td align=""center"">" & Cells(31 + i, 9).value & "</td>"
    EmailBody = EmailBody & "</tr>"
Next i

EmailBody = EmailBody & "</table>"
EmailBody = EmailBody & "<hr>" & "Report generated on " & Now()

On Error Resume Next

For X = 71 To 91 ' sciaganie adresow
    If Cells(X, 8).value <> "" Then
        adressee = adressee & Cells(X, 8).value & "@grupoantolin.com; "
    End If
Next X

    With NewMail
        .To = adressee
        .cc = ""
        .BCC = ""
        .Subject = "F5x IP Despatch Report"
        .htmlbody = EmailBody & "</BODY>" & vbNewLine & Signature
        '.Attachments.Add (Address)
        '.display
        .send
    End With
    
With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

Set NewMail = Nothing
Set OlApp = Nothing
Sheets("BRIEF").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True

MsgBox "IP report was sent.", vbOKOnly, "Message sent"

End Sub
'//===============================================================================================================

Sub DS_collections()

With Application
    .ScreenUpdating = False
End With

Dim OlApp As Object
Dim NewMail As Object
Dim EmailBody As String
Dim StrSignature As String
Dim sPath As String
Dim strString
Dim strResult
Dim seqLead As Integer
Dim Address As String
    
Set OlApp = CreateObject("Outlook.Application")
Set NewMail = OlApp.CreateItem(0)

'With NewMail
'.display
'End With
Signature = NewMail.htmlbody

    
EmailBody = "<BODY style=font-size:10pt;font-family:Calibri><b>R1 DS COLLECTIONS REPORT</b><hr>"
EmailBody = EmailBody & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=font-size:10pt;font-family:Calibri; text-align:center;>"
EmailBody = EmailBody & "<tr><b><td align=""center""></td>"
EmailBody = EmailBody & "<td align=""center"">&nbsp;PLANNED&nbsp;</td>"
EmailBody = EmailBody & "<td align=""center"">&nbsp;&nbsp;ZONE&nbsp;3&nbsp;&nbsp;</td>"
EmailBody = EmailBody & "<td align=""center"">&nbsp;&nbsp;DOCK&nbsp;&nbsp;</td>"
EmailBody = EmailBody & "<td align=""center"">&nbsp;DEPARTURE&nbsp;</td>"
EmailBody = EmailBody & "<td align=""center"">&nbsp;COMMENTS&nbsp;</td></b></tr>"


For i = 1 To 7
    EmailBody = EmailBody & "<tr><td align=""center"">" & i & ".</td>"
    EmailBody = EmailBody & "<td align=""center"">" & Format(Cells(39 + i, 5).value, "hh:mm") & "</td>"
    EmailBody = EmailBody & "<td align=""center"">" & Format(Cells(39 + i, 6).value, "hh:mm") & "</td>"
    EmailBody = EmailBody & "<td align=""center"">" & Format(Cells(39 + i, 7).value, "hh:mm") & "</td>"
    EmailBody = EmailBody & "<td align=""center"">" & Format(Cells(39 + i, 8).value, "hh:mm") & "</td>"
    EmailBody = EmailBody & "<td align=""center"">" & Cells(39 + i, 9).value & "</td>"
    EmailBody = EmailBody & "</tr>"
Next i

EmailBody = EmailBody & "</table>"
EmailBody = EmailBody & "<hr>" & "Report generated on " & Now()

On Error Resume Next

For X = 71 To 91 ' sciaganie adresow
    If Cells(X, 6).value <> "" Then
        adressee = adressee & Cells(X, 6).value & "@grupoantolin.com; "
    End If
Next X

    With NewMail
        .To = adressee
        .cc = ""
        .BCC = ""
        .Subject = "R1 DS Collections Report"
        .htmlbody = EmailBody & "</BODY>" & vbNewLine & Signature
        '.Attachments.Add (Address)
        '.display
        .send
    End With
    
With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

MsgBox "DS Report was sent. ", vbOKOnly, "Message sent"

Set NewMail = Nothing
Set OlApp = Nothing
Sheets("BRIEF").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True

End Sub
'//===============================================================================================================

Function RangetoHTML(rng As Range)
' By Ron de Bruin.
    Dim FSO As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         FileName:=TempFile, _
         sheet:=TempWB.Sheets(1).name, _
         source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ts = FSO.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set FSO = Nothing
    Set TempWB = Nothing
End Function
'//===============================================================================================================

Sub deliverses()
    On Error Resume Next
    Sheets("Delivery Log").Select
End Sub
'//===============================================================================================================
Sub gohome()
    On Error Resume Next
    Sheets("BRIEF").Select
End Sub
'//===============================================================================================================