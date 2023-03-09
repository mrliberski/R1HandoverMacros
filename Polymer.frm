VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Polymer 
   Caption         =   "Polymer Request Form"
   ClientHeight    =   10260
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   7188
   OleObjectBlob   =   "Polymer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Polymer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'open classic polymer request
On Error Resume Next
Dim wkb As Workbook
Dim name As String
name = ThisWorkbook.name
Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\Polymer Request Form.xlsm")
Polymer.Hide
Workbooks(name).Activate
ActiveWindow.WindowState = xlNormal

End Sub


Private Sub CommandButton2_Click()
'On Error Resume Next
Sheets("Poly Req Log").Unprotect


Dim harvest As String
Dim LastRow As Long, ws As Worksheet, rfnumber As String
Set ws = Sheets("Poly Req Log")

If Polymer.CheckBox1.value = True Then
    harvest = harvest & "RF/80347/000&nbsp;&nbsp;&nbsp;&nbsp;<b>PP Glass (Jaguar Toppers)</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;950 kg"
    rfnumber = "RF/80347/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox2.value = True Then
    harvest = harvest & "<br>RF/24847/000&nbsp;&nbsp;&nbsp;&nbsp;<b>ABS  Magnum</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Natural</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/24847/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox3.value = True Then
    harvest = harvest & "<br>RF/24747/000&nbsp;&nbsp;&nbsp;&nbsp;<b>ABS  H604</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;900 kg"
    rfnumber = "RF/24747/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox4.value = True Then
    harvest = harvest & "<br>RF/81190/000&nbsp;&nbsp;&nbsp;&nbsp;<b>FINALLOY</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/81190/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox5.value = True Then
    harvest = harvest & "<br>RF/24845/000&nbsp;&nbsp;&nbsp;&nbsp;<b>PC/ABS Pulse A35-105</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/24845/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox6.value = True Then
    harvest = harvest & "<br>RF/27590/000&nbsp;&nbsp;&nbsp;&nbsp;<b>PP TRC 333N</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1100 kg"
    rfnumber = "RF/27590/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox7.value = True Then
    harvest = harvest & "<br>RF/24845/000&nbsp;&nbsp;&nbsp;&nbsp;<b>PC/ABS Pulse A35-105</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/24845/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox8.value = True Then
    harvest = harvest & "<br>RF/27081/000&nbsp;&nbsp;&nbsp;&nbsp;<b>NYLON (Akulon P66)</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/27081/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox9.value = True Then
    harvest = harvest & "<br>RF/25856/000&nbsp;&nbsp;&nbsp;&nbsp;<b>PA/ABS  Terblend NM 19</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/25856/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox10.value = True Then
    harvest = harvest & "<br>RF/81042/000&nbsp;&nbsp;&nbsp;&nbsp;<b>PA/ABS  Terblend NNG 02EF</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/81042/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox11.value = True Then
    harvest = harvest & "<br>RF/27606/000&nbsp;&nbsp;&nbsp;&nbsp;<b>PC/ABS GX 50</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/27606/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox12.value = True Then
    harvest = harvest & "<br>RF/80357/000&nbsp;&nbsp;&nbsp;&nbsp;<b>TKG 300N BASELL PP</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Natural</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/80357/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox13.value = True Then
    harvest = harvest & "<br>RF/24849/000&nbsp;&nbsp;&nbsp;&nbsp;<b>TF 1305 PP</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/24849/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox14.value = True Then
    harvest = harvest & "<br>RF/28574/000&nbsp;&nbsp;&nbsp;&nbsp;<b>PA/ASA</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/28574/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox15.value = True Then
    harvest = harvest & "<br>RF/82163/000&nbsp;&nbsp;&nbsp;&nbsp;<b>HKC 431N BASELL PP</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/82163/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox16.value = True Then
    harvest = harvest & "<br>RF/24262/000&nbsp;&nbsp;&nbsp;&nbsp;<b>L-319 - H604 EBONY</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;900 kg"
    rfnumber = "RF/24262/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox17.value = True Then
    harvest = harvest & "<br>RF/80354/000&nbsp;&nbsp;&nbsp;&nbsp;<b>TKG 300N BASELL PP</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Jet Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/80354/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox18.value = True Then
    harvest = harvest & "<br>RF/29684/000&nbsp;&nbsp;&nbsp;&nbsp;<b>Fibre glass 33% for SGI</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/29684/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox19.value = True Then
    harvest = harvest & "<br>RF/27591/000&nbsp;&nbsp;&nbsp;&nbsp;<b>PP  TRC 333N</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Grey</i>&nbsp;&nbsp;&nbsp;&nbsp;1100 kg"
    rfnumber = "RF/27591/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox20.value = True Then
    harvest = harvest & "<br>RF/28590/000&nbsp;&nbsp;&nbsp;&nbsp;<b>Foaming agent 1.8% for SGI</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Natural</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/28590/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox21.value = True Then
    harvest = harvest & "<br>RF/27631/000&nbsp;&nbsp;&nbsp;&nbsp;<b>Masterbatch 2% for Tailgates</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;200 kg"
    rfnumber = "RF/27631/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox22.value = True Then
    harvest = harvest & "<br>RF/27608/000&nbsp;&nbsp;&nbsp;&nbsp;<b>2K Unit Santropene</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Natural</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/27608/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox23.value = True Then
    harvest = harvest & "<br>RF/28499/000&nbsp;&nbsp;&nbsp;&nbsp;<b>Masterbatch 2% for 2K Unit Santropene</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "RF/28499/000"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox24.value = True Then
    harvest = harvest & "<br>115009840/001&nbsp;&nbsp;&nbsp;&nbsp;<b>TKC 451 for L551</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Black</i>&nbsp;&nbsp;&nbsp;&nbsp;1100 kg"
    rfnumber = "115009840/001"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
If Polymer.CheckBox25.value = True Then
    harvest = harvest & "<br>PURGEX&nbsp;&nbsp;&nbsp;&nbsp;<b>Cleaning agent</b>&nbsp;&nbsp;&nbsp;&nbsp;<i>Natural</i>&nbsp;&nbsp;&nbsp;&nbsp;1000 kg"
    rfnumber = "PURGEX"
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    ws.Range("A" & LastRow).value = Now() 'Adds timestamp of request
    ws.Range("B" & LastRow).value = rfnumber 'Adds part number requested
    ws.Range("C" & LastRow).value = Environ("username")
End If
Sheets("Poly Req Log").Protect
Call polymer_request(harvest)
Polymer.Hide

End Sub

Private Sub CommandButton3_Click()
'cancel button
Polymer.Hide
ThisWorkbook.Save
End Sub
