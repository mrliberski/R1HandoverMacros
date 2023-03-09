VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Control Panel"
   ClientHeight    =   4860
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   8004
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'delivery
UserForm1.Hide
deliveryreq.TextBox1.value = ""
deliveryreq.Show
End Sub

Private Sub CommandButton10_Click()
'sorion
ActiveWorkbook.FollowHyperlink Address:="http://re2vmsor01/orionukl/reporting/home.php", NewWindow:=True
UserForm1.Hide
End Sub

Private Sub CommandButton11_Click()
'qms
ActiveWorkbook.FollowHyperlink Address:="http://qmslive/", NewWindow:=True
UserForm1.Hide
End Sub

Private Sub CommandButton12_Click()
'beone
ActiveWorkbook.FollowHyperlink Address:="http://beone.grupoantolin.com/Pages/default.aspx", NewWindow:=True
UserForm1.Hide

End Sub

Private Sub CommandButton13_Click()
UserForm1.Hide
Call send_handover
End Sub

Private Sub CommandButton14_Click()
'attendance
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\Attendance\Red1-Employees Attendance Sheet.xlsm")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////

End Sub

Private Sub CommandButton15_Click()
'headcount
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-PRODUCTION\Redditch Production Data\Redditch 1\Redditch 1 Daily Headcount", vbNormalFocus)
UserForm1.Hide

End Sub

Private Sub CommandButton16_Click()
'overtime tracker
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\public\Pub-COMMON\OVERTIME\Overtime planner (PRODUCTION & LOGISTICS).xlsx", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton17_Click()
'Mannin online
ActiveWorkbook.FollowHyperlink Address:="https://forms.office.com/Pages/ResponsePage.aspx?id=PYnoqaTxjU2Xj4Yu39Noavt9AYia845OsbsgcDrGelxUOE4wNzhRS1Y5R1o1OFQ2WkhIMjlWOTFZMS4u", NewWindow:=True
UserForm1.Hide
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton18_Click()
'temp timesheet
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\Attendance\Red1-Temporary Employees TimeSheet.xlsx")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton19_Click()
'adp
ActiveWorkbook.FollowHyperlink Address:="https://ihcm.adp.com/", NewWindow:=True
UserForm1.Hide
End Sub

Private Sub CommandButton2_Click()
UserForm1.Hide
'StockCount.OptionButton1.value = False
StockCount.Show
End Sub

Private Sub CommandButton20_Click()
'rotation sheets
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\Shifts", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton21_Click()
'fire register
On Error Resume Next
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\Shifts\RED-EHS-034 ATTENDANCE REGISTER FIRE ROLL CALL SHEET.xls")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
End Sub

Private Sub CommandButton22_Click()
'brief signoff
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("J:\Pub-LOGISTICS\Shift Folder New\2019\Shift Brief Sign off.xlsx")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton23_Click() 'hide
UserForm1.Hide
End Sub

Private Sub CommandButton24_Click() 'hide
UserForm1.Hide
End Sub

Private Sub CommandButton25_Click() 'hide
UserForm1.Hide
End Sub

Private Sub CommandButton26_Click()
'empty pck
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\Customer\Red1-Empty Packaging Tracker Log.xlsm")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton27_Click()
'cms
'////////////////////////////////////////////////////////
'Dim wkb As Workbook
'Dim name As String
'name = ThisWorkbook.name
'Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\Customer\Redditch 1-CMS Count Log.xlsm")
UserForm1.Hide
'Workbooks(name).Activate
'ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
PackDaily.Show
End Sub

Private Sub CommandButton28_Click()
'supplier return
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\Daily Doc's\Packaging\RED-LOG-060 - Supplier Empties Return Sheet.xlsx")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
End Sub

Private Sub CommandButton29_Click()
'trailer check
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\Daily Doc's\R1-Trailer Check Sheet.xlsx")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton3_Click()

'show copyright
Polymer.Label140.Caption = "Copyright " & Chr(169) & " Pawel Liberski 2019"
'hide control page
UserForm1.Hide

' reset all checkboxes////////////////////////
Dim contr As Control
For Each contr In Polymer.Controls
    If TypeName(contr) = "CheckBox" Then
        contr.value = False
    End If
Next

'show polymer form
Polymer.Show

End Sub

Private Sub CommandButton30_Click()
'flt repairs
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Packaging\MHE Repair log NEW.xlsx")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton31_Click()
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\UET", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton32_Click()
'low stock
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\Low Stock Warning Form New.xls")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton33_Click()
'racking labels
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\Racking Labels", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton34_Click()
'ds bingo sheet
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\Daily Doc's\RED-2045-056 - DS Loading Bingo Sheet.xlsx")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton35_Click()
'packaging tally
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\Daily Doc's\Packaging\R1-BMW Returnable Packaging Tally Sheet.xlsx")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton36_Click()
'matrix
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\Matrix", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton37_Click()
'battery top up
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\Daily Doc's\FLT Battries TOP UP MATRIX 2018.xlsx")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton38_Click()
'DS UET BOARD
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Direct Supply\UET\RED1", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton39_Click()
'manual gdn
'\\RE2VMFIL02\Public\Pub-LOGISTICS\GDN
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\GDN", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton4_Click()
'ip

answer = MsgBox("IP report will be sent now. Do you wish to continue?", vbOKCancel, "Confirmation required")
If answer = vbCancel Then Exit Sub

UserForm1.Hide
Call ipnew

End Sub

Private Sub CommandButton40_Click()
'odette label

UserForm1.Hide

'Odette.TextBox1.Value = ""
'Odette.TextBox2.Value = ""
'Odette.TextBox3.Value = ""
'Odette.TextBox4.Value = ""
'Odette.TextBox5.Value = ""
'Odette.TextBox6.Value = ""

Odette.Show


'////////////////////////////////////////////////////////
'Dim wkb As Workbook
'Dim name As String
'name = ThisWorkbook.name
'Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\Daily Doc's\R1 - Odette label.xlsm")
'UserForm1.Hide
'Workbooks(name).Activate
'ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton41_Click()
'5S folder
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\UET\5'S", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton42_Click()
'layout plan
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\Layout Plans", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton43_Click()
'logistics folder
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton44_Click()
'UET boards
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\UET", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton45_Click()
'monthly count
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\Daily Doc's\Packaging\Monthly Returnable Packaging Count Sheet.xlsx")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton46_Click()
'transports
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\Transport\Transport Bookings", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton47_Click()
'lateness form
Dim name As String
name = ThisWorkbook.name
ActiveWorkbook.FollowHyperlink "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\HR\2018\Lateness Form.docx"
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
End Sub

Private Sub CommandButton48_Click()
'back to work
Dim name As String
name = ThisWorkbook.name
ActiveWorkbook.FollowHyperlink "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\HR\2018\Self Certification Back to work interview.pdf"
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
End Sub

Private Sub CommandButton49_Click()
'auth absence
Dim name As String
name = ThisWorkbook.name
ActiveWorkbook.FollowHyperlink "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\HR\2018\AUTHORISED ABSENCE FORM.doc"
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
End Sub

Private Sub CommandButton5_Click()
'ds report

answer = MsgBox("DS report will be sent now. Do you wish to continue?", vbOKCancel, "Confirmation required")
If answer = vbCancel Then Exit Sub

UserForm1.Hide
Call DS_collections

End Sub

Private Sub CommandButton50_Click()
'////////////////////////////////////////////////////////
'Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
'Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\HR\2018\MOD Uninform Order Form.pdf")
ActiveWorkbook.FollowHyperlink "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\HR\2018\MOD Uninform Order Form.pdf"
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton51_Click()
'postit
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\POST IT", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton52_Click()
'sop folder
Call Shell("explorer.exe" & " " & "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\SOP Folder", vbNormalFocus)
UserForm1.Hide
End Sub

Private Sub CommandButton54_Click()
'damaged packaging
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Direct Supply\Packaging\BMW Damaged PKG LOG.xlsx")
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub

Private Sub CommandButton55_Click()
'unlock all
On Error Resume Next
Sheets("BRIEF").Unprotect
Sheets("Poly Req Log").Unprotect
Sheets("Delivery Log").Unprotect

End Sub

Private Sub CommandButton56_Click()
'lock all
On Error Resume Next
Sheets("BRIEF").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True
Sheets("Poly Req Log").Protect
Sheets("Delivery Log").Protect
End Sub

Private Sub CommandButton57_Click()
'unhide all
On Error Resume Next
Application.ScreenUpdating = False
Sheets("Sheet1").Visible = True
Sheets("Sheet2").Visible = True
Sheets("Products").Visible = True
Sheets("Odette").Visible = True
Sheets("Folder Labels").Visible = xlVisible
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton58_Click()
'hide all
On Error Resume Next
Sheets("Sheet1").Visible = xlVeryHidden
Sheets("Sheet2").Visible = xlHidden
Sheets("Products").Visible = xlVeryHidden
Sheets("Odette").Visible = xlVeryHidden
End Sub

Private Sub CommandButton6_Click()
'l551
UserForm1.Hide
L551_Report.Show
'Call L551_despatch
End Sub

Private Sub CommandButton62_Click()
'witness statement
'\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\HR\Interview Record - E Template.docx
Dim name As String
name = ThisWorkbook.name
ActiveWorkbook.FollowHyperlink "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\HR\Interview Record - E Template.docx"
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
End Sub

Private Sub CommandButton63_Click()
'yellow postit
Dim name As String
name = ThisWorkbook.name
ActiveWorkbook.FollowHyperlink "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\POST IT\Post It Internal - Yellow.xslx"
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
End Sub

Private Sub CommandButton64_Click()
'green post it
Dim name As String
name = ThisWorkbook.name
ActiveWorkbook.FollowHyperlink "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\POST IT\HSE  POST IT 2018-001 DANGEROUS OCCURRENCE - TEMPLATE.xls"
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
End Sub

Private Sub CommandButton65_Click()
'red post it
Dim name As String
name = ThisWorkbook.name
ActiveWorkbook.FollowHyperlink "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\POST IT\HSE  POST IT 2018-001 ACCIDENT - TEMPLATE.xls"
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
End Sub

Private Sub CommandButton66_Click()
'ADP
ActiveWorkbook.FollowHyperlink Address:="https://online.elearn365.co.uk/", NewWindow:=True
UserForm1.Hide
End Sub

Private Sub CommandButton67_Click()

End Sub

Private Sub CommandButton68_Click() ' edit adressee
' hideme Macro


If CommandButton68.Caption = "Show Recipients" Then

    ActiveWindow.SmallScroll Down:=39
    Rows("69:90").Select
    Range("B69").Activate
    ActiveSheet.Unprotect
    Selection.EntireRow.Hidden = False
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True
    CommandButton68.Caption = "Hide Recipients"
    UserForm1.Hide
    
ElseIf CommandButton68.Caption = "Hide Recipients" Then

    ActiveWindow.SmallScroll Down:=39
    Rows("69:90").Select
    Range("B69").Activate
    ActiveSheet.Unprotect
    Selection.EntireRow.Hidden = True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True
    CommandButton68.Caption = "Show Recipients"
    UserForm1.Hide
    Range("A1").Select
    ActiveWindow.SmallScroll Up:=0
    
Else: CommandButton68.Caption = "Hide Recipients"
End If



End Sub

Private Sub CommandButton69_Click()
'racking inspection
Dim name As String
name = ThisWorkbook.name
ActiveWorkbook.FollowHyperlink "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\Daily Doc's\Red1-Racking inspection 2020.xlsx"
UserForm1.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal

End Sub

Private Sub CommandButton7_Click()
'ADP
ActiveWorkbook.FollowHyperlink Address:="https://ihcm.adp.com/", NewWindow:=True
UserForm1.Hide
End Sub

Private Sub CommandButton70_Click()
'racvking labelz
On Error Resume Next
Dim Address
Address = "\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2020\REDDITCH 1\Stores"
Call Shell("explorer.exe" & " " & Address, vbNormalFocus)
UserForm1.Hide

End Sub

Private Sub CommandButton71_Click()
'delivery log
On Error Resume Next
Sheets("Delivery Log").Select
UserForm1.Hide
End Sub

Private Sub CommandButton72_Click()
'booking
UserForm1.Hide

'=============================== CHECK IF DATABASE IS ACCESSIBLE
Dim FPath As String
FPath = "J:\Pub-LOGISTICS\Packaging\Packaging.accdb"
If Dir(FPath) = "" Then
    MsgBox "Database is not accessible. Please try again later.", vbOKOnly, "Could not find database."
    Exit Sub
End If
'======================================================================

Booking.Show

End Sub

Private Sub CommandButton73_Click()
UserForm1.Hide
PackagingBrowse.Show
End Sub

Private Sub CommandButton74_Click()
UserForm1.Hide
'==============================
Dim rownum
rownum = 2 ' first row of list
Editor.TextBox1.text = Sheet18.Cells(rownum, 1).value
Editor.Show
End Sub

Private Sub CommandButton75_Click()
Application.ThisWorkbook.Save
UserForm1.Hide
End Sub


Private Sub CommandButton76_Click()
UserForm1.Hide

If UCase(Environ("username")) = "PAWEL.LIBERSKI" Then
    AssesmentAdd.CheckBox1.Enabled = True
        AssesmentAdd.CheckBox1.value = True
    AssesmentAdd.CheckBox2.Enabled = True
    AssesmentAdd.CheckBox3.Enabled = True
    AssesmentAdd.CheckBox4.Enabled = True
    AssesmentAdd.CheckBox5.Enabled = True

Else
    AssesmentAdd.CheckBox1.Enabled = False
    AssesmentAdd.CheckBox2.Enabled = True
        AssesmentAdd.CheckBox2.value = True ' default R1
    AssesmentAdd.CheckBox3.Enabled = False
    AssesmentAdd.CheckBox4.Enabled = False
    AssesmentAdd.CheckBox5.Enabled = False
End If

AssesmentAdd.Show

End Sub

Private Sub CommandButton77_Click()
    UserForm1.Hide
    Odette.Show
End Sub

Private Sub CommandButton78_Click()
UserForm1.Hide
UserForm2.Show
End Sub

Private Sub CommandButton79_Click()
'timesheet 2021
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("J:\Pub-LOGISTICS\Shift Folder New\2021\Attendance\Redditch 2\Redditch 2 Timesheet 2022.xlsx")
UserForm1.Hide
End Sub

Private Sub CommandButton8_Click()
'line lead
ActiveWorkbook.FollowHyperlink Address:="http://re2vmrep01/EndOfLine/Red1LeadChartAll.aspx?&back=1", NewWindow:=True
UserForm1.Hide
End Sub

Private Sub CommandButton80_Click()
 UserForm1.Hide
 DSReport.Show
End Sub

Private Sub CommandButton9_Click()
'it ticket
'https://bit.ly/2RkOS4b
ActiveWorkbook.FollowHyperlink Address:="https://bit.ly/2RkOS4b", NewWindow:=True
UserForm1.Hide
End Sub

