VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Odette 
   Caption         =   "Generate Odette Label "
   ClientHeight    =   2490
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5328
   OleObjectBlob   =   "Odette.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Odette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Odette.Hide
End Sub

Private Sub CommandButton2_Click()
'MsgBox "Not ready yet..."
On Error Resume Next
Application.ScreenUpdating = False
Sheets("Products").Visible = True
Sheets("Odette").Visible = True
Sheets("Odette").Activate


If Odette.TextBox1.value = "" Then
    Sheets("Odette").Cells(4, 1).value = "VOID"
Else: Sheets("Odette").Cells(4, 1).value = UCase(Odette.TextBox1.value)
End If
If Odette.TextBox2.value = "" Then
    Sheets("Odette").Cells(8, 1).value = "VOID"
Else: Sheets("Odette").Cells(8, 1).value = UCase(Odette.TextBox2.value)
End If
If Odette.TextBox3.value = "" Then
    Sheets("Odette").Cells(11, 1).value = "VOID"
Else: Sheets("Odette").Cells(11, 1).value = UCase(Odette.TextBox3.value)
End If
If Odette.TextBox4.value = "" Then
    Sheets("Odette").Cells(28, 1).value = "VOID"
Else: Sheets("Odette").Cells(28, 1).value = UCase(Odette.TextBox4.value)
End If
If Odette.TextBox5.value = "" Then
    Sheets("Odette").Cells(32, 1).value = "VOID"
Else: Sheets("Odette").Cells(32, 1).value = UCase(Odette.TextBox5.value)
End If
If Odette.TextBox6.value = "" Then
    Sheets("Odette").Cells(35, 1).value = "VOID"
Else: Sheets("Odette").Cells(35, 1).value = UCase(Odette.TextBox6.value)
End If

Call RenderQRCode("Odette", "A5", "N" & Cells(4, 1).value)
Call RenderQRCode("Odette", "A9", "P" & Cells(8, 1).value)
Call RenderQRCode("Odette", "A13", "Q" & Cells(11, 1).value)

Call RenderQRCode("Odette", "A29", "N" & Cells(28, 1).value)
Call RenderQRCode("Odette", "A33", "P" & Cells(32, 1).value)
Call RenderQRCode("Odette", "A37", "Q" & Cells(35, 1).value)

ActiveSheet.PageSetup.PrintArea = "$A$1:$E$46"
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
            IgnorePrintAreas:=False


Sheets("Odette").Cells(4, 1).value = "No Data"
Sheets("Odette").Cells(8, 1).value = "No Data"
Sheets("Odette").Cells(11, 1).value = "No Data"
Sheets("Odette").Cells(28, 1).value = "No Data"
Sheets("Odette").Cells(32, 1).value = "No Data"
Sheets("Odette").Cells(35, 1).value = "No Data"

Call RenderQRCode("Odette", "A5", "N" & Cells(4, 1).value)
Call RenderQRCode("Odette", "A9", "P" & Cells(8, 1).value)
Call RenderQRCode("Odette", "A13", "Q" & Cells(11, 1).value)

Call RenderQRCode("Odette", "A29", "N" & Cells(28, 1).value)
Call RenderQRCode("Odette", "A33", "P" & Cells(32, 1).value)
Call RenderQRCode("Odette", "A37", "Q" & Cells(35, 1).value)

Sheets("Products").Visible = xlVeryHidden
Sheets("Odette").Visible = xlVeryHidden
Sheets("BRIEF").Activate
Application.ScreenUpdating = True
Odette.Hide
UserForm1.Show

End Sub

Private Sub CommandButton3_Click()
'////////////////////////////////////////////////////////
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2018\Redditch 1\Daily Doc's\R1 - Odette label.xlsm")
Odette.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal
'///////////////////////////////////////////////////////
End Sub
