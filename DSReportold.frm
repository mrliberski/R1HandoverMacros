VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DSReportold 
   Caption         =   "BMW Despatch Report"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6588
   OleObjectBlob   =   "DSReportold.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DSReportold"
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
        .To = "pugh, richard; celec, marek; blake, yuell; hussain, ishtiak;"
        .cc = "Liberski, Pawel; partridge, jamie; sim, alina; oliver, yvonne; wootton, louise; leighton, rebecca; cendrowska, alexandra; southall, paula; kaczynska, daria;"
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
