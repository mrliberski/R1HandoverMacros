VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PackDaily 
   Caption         =   "Packaging count"
   ClientHeight    =   8505.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5628
   OleObjectBlob   =   "PackDaily.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PackDaily"
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
Private Sub Com()

    'sort out blank fields and set them to zero, otherwise it crashes
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.name Like "Text*" Then
            If ctrl = NullString Then ctrl = 0
        End If
    Next

    'read fields and do some maths
    Dim count3083480 As Integer
        count3083480 = PackDaily.TextBox1.value
        PackDaily.TextBox1.value = NullString
    Dim ppc3083480 As Integer
        ppc3083480 = 20 'parts per container
    Dim cpp3083480 As Integer
        cpp3083480 = 1 ' containers per pallet
    Dim cycle3083480 As Integer
        cycle3083480 = 60 / 1 'cycle time
    Dim coverage3083480 As Integer
        coverage3083480 = (count3083480 * ppc3083480 * cpp3083480) / (3600 / cycle3083480)
    
    Dim count3084151 As Integer
        count3084151 = PackDaily.TextBox2.value
        PackDaily.TextBox2.value = NullString
    Dim ppc3084151 As Integer
        ppc3084151 = 20 'parts per container
    Dim cpp3084151 As Integer
        cpp3084151 = 1 ' containers per pallet
    Dim cycle3084151 As Integer
        cycle3084151 = 60 / 1 'cycle time
    Dim coverage3084151 As Integer
        coverage3084151 = (count3084151 * ppc3084151 * cpp3084151) / (3600 / cycle3084151)
    
    Dim count3084308 As Integer
        count3084308 = PackDaily.TextBox3.value
        PackDaily.TextBox3.value = NullString
    Dim ppc3084308 As Integer
        ppc3084308 = 80 'parts per container
    Dim cpp3084308 As Integer
        cpp3084308 = 1 ' containers per pallet
    Dim cycle3084308 As Integer
        cycle3084308 = 50 / 2 'cycle time
    Dim coverage3084308 As Integer
        coverage3084308 = (count3084308 * ppc3084308 * cpp3084308) / (3600 / cycle3084308)
    
    Dim count3084309 As Integer
        count3084309 = PackDaily.TextBox4.value
        PackDaily.TextBox4.value = NullString
    Dim ppc3084309 As Integer
        ppc3084309 = 27 'parts per container
    Dim cpp3084309 As Integer
        cpp3084309 = 8 ' containers per pallet
    Dim cycle3084309 As Integer
        cycle3084309 = 70 / 2 'cycle time
    Dim coverage3084309 As Integer
        coverage3084309 = (count3084309 * ppc3084309 * cpp3084309) / (3600 / cycle3084309)
    
    Dim count3084310 As Integer
        count3084310 = PackDaily.TextBox5.value
        PackDaily.TextBox5.value = NullString
    Dim ppc3084310 As Integer
        ppc3084310 = 24 'parts per container
    Dim cpp3084310 As Integer
        cpp3084310 = 4 ' containers per pallet
    Dim cycle3084310 As Integer
        cycle3084310 = 56 / 2 'cycle time
    Dim coverage3084310 As Integer
        coverage3084310 = (count3084310 * ppc3084310 * cpp3084310) / (3600 / cycle3084310)
    
    Dim count3084317 As Integer
        count3084317 = PackDaily.TextBox6.value
        PackDaily.TextBox6.value = NullString
    Dim ppc3084317 As Integer
        ppc3084317 = 48 'parts per container
    Dim cpp3084317 As Integer
        cpp3084317 = 1 ' containers per pallet
    Dim cycle3084317 As Integer
        cycle3084317 = 60 / 4 'cycle time
    Dim coverage3084317 As Integer
        coverage3084317 = (count3084317 * ppc3084317 * cpp3084317) / (3600 / cycle3084317)
    
    Dim count3084320 As Integer
        count3084320 = PackDaily.TextBox7.value
        PackDaily.TextBox7.value = NullString
    Dim ppc3084320 As Integer
        ppc3084320 = 48 'parts per container
    Dim cpp3084320 As Integer
        cpp3084320 = 1 ' containers per pallet
    Dim cycle3084320 As Integer
        cycle3084320 = 60 'cycle time
    Dim coverage3084320 As Integer
        coverage3084320 = (count3084320 * ppc3084320 * cpp3084320) / (3600 / cycle3084320)
    
 'missing two divider numbers
 
    Dim count3103210 As Integer
        count3103210 = PackDaily.TextBox8.value
        PackDaily.TextBox8.value = NullString
    Dim count3104802 As Integer
        count3104802 = PackDaily.TextBox9.value
        PackDaily.TextBox9.value = NullString
 
    Dim count6200294 As Integer
        count6200294 = PackDaily.TextBox10.value
        PackDaily.TextBox10.value = NullString
    Dim ppc6200294 As Integer
        ppc6200294 = 10 'parts per container
    Dim cpp6200294 As Integer
        cpp6200294 = 4 ' containers per pallet
    Dim cycle6200294 As Integer
        cycle6200294 = 54 / 2 'cycle time
    Dim coverage6200294 As Integer
        coverage6200294 = (count6200294 * ppc6200294 * cpp6200294) / (3600 / cycle6200294)
    
    Dim count6200852 As Integer
        count6200852 = PackDaily.TextBox11.value
        PackDaily.TextBox11.value = NullString
    Dim ppc6200852 As Integer
        ppc6200852 = 54 'parts per container
    Dim cpp6200852 As Integer
        cpp6200852 = 1 ' containers per pallet
    Dim cycle6200852 As Integer
        cycle6200852 = 60 / 2 'cycle time / cavities
    Dim coverage6200852 As Integer
        coverage6200852 = (count6200852 * ppc6200852 * cpp6200852) / (3600 / cycle6200852)
    
    Dim count6200854 As Integer
        count6200854 = PackDaily.TextBox12.value
        PackDaily.TextBox12.value = NullString
    Dim ppc6200854 As Integer
        ppc6200854 = 40 'parts per container
    Dim cpp6200854 As Integer
        cpp6200854 = 1 ' containers per pallet
    Dim cycle6200854 As Integer
        cycle6200854 = 65 / 1 'cycle time / cavities
    Dim coverage6200854 As Integer
        coverage6200854 = (count6200854 * ppc6200854 * cpp6200854) / (3600 / cycle6200854)
    
    Dim count6200858 As Integer
        count6200858 = PackDaily.TextBox13.value
        PackDaily.TextBox13.value = NullString
    Dim ppc6200858 As Integer
        ppc6200858 = 28 'parts per container
    Dim cpp6200858 As Integer
        cpp6200858 = 8 ' containers per pallet
    Dim cycle6200858 As Integer
        cycle6200858 = 60 / 2 'cycle time / cavities
    Dim coverage6200858 As Integer
        coverage6200858 = (count6200858 * ppc6200858 * cpp6200858) / (3600 / cycle6200858)
    
    Dim count6202426 As Integer
        count6202426 = PackDaily.TextBox14.value
        PackDaily.TextBox14.value = NullString
    Dim ppc6202426 As Integer
        ppc6202426 = 78 'parts per container
    Dim cpp6202426 As Integer
        cpp6202426 = 1 ' containers per pallet
    Dim cycle6202426 As Integer
        cycle6202426 = 52 / 2 'cycle time / cavities
    Dim coverage6202426 As Integer
        coverage6202426 = (count6202426 * ppc6202426 * cpp6202426) / (3600 / cycle6202426)
    
    Dim count6202434 As Integer
        count6202434 = PackDaily.TextBox15.value
        PackDaily.TextBox15.value = NullString
    Dim ppc6202434 As Integer
        ppc6202434 = 60 / 2 'parts per container
    Dim cpp6202434 As Integer
        cpp6202434 = 1 ' containers per pallet
    Dim cycle6202434 As Integer
        cycle6202434 = 60 / 2 'cycle time / cavities
    Dim coverage6202434 As Integer
        coverage6202434 = (count6202434 * ppc6202434 * cpp6202434) / (3600 / cycle6202434)
    
    Dim count6202789 As Integer
        count6202789 = PackDaily.TextBox16.value
        PackDaily.TextBox16.value = NullString
    Dim ppc6202789 As Integer
        ppc6202789 = 14 'parts per container
    Dim cpp6202789 As Integer
        cpp6202789 = 12 ' containers per pallet
    Dim cycle6202789 As Integer
        cycle6202789 = 54 / 2 'cycle time / cavities
    Dim coverage6202789 As Integer
        coverage6202789 = (count6202789 * ppc6202789 * cpp6202789) / (3600 / cycle6202789)
    
    Dim count6202790 As Integer
        count6202790 = PackDaily.TextBox17.value
        PackDaily.TextBox17.value = NullString
    Dim ppc6202790 As Integer
        ppc6202790 = 12 'parts per container
    Dim cpp6202790 As Integer
        cpp6202790 = 12 ' containers per pallet
    Dim cycle6202790 As Integer
        cycle6202790 = 54 / 2 'cycle time / cavities
    Dim coverage6202790 As Integer
        coverage6202790 = (count6202790 * ppc6202790 * cpp6202790) / (3600 / cycle6202790)
    
    Dim count6203991 As Integer
        count6203991 = PackDaily.TextBox18.value
        PackDaily.TextBox18.value = NullString
    Dim ppc6203991 As Integer
        ppc6203991 = 40 'parts per container
    Dim cpp6203991 As Integer
        cpp6203991 = 1 ' containers per pallet
    Dim cycle6203991 As Integer
        cycle6203991 = 60 / 1 'cycle time / cavities
    Dim coverage6203991 As Integer
        coverage6203991 = (count6203991 * ppc6203991 * cpp6203991) / (3600 / cycle6203991)
    
    Dim count6204136 As Integer
        count6204136 = PackDaily.TextBox19.value
        PackDaily.TextBox19.value = NullString
    Dim ppc6204136 As Integer
        ppc6204136 = 26 'parts per container
    Dim cpp6204136 As Integer
        cpp6204136 = 8 ' containers per pallet
    Dim cycle6204136 As Integer
        cycle6204136 = 54 / 2 'cycle time / cavities
    Dim coverage6204136 As Integer
        coverage6204136 = (count6204136 * ppc6204136 * cpp6204136) / (3600 / cycle6204136)
    
    Dim count6206723 As Integer
        count6206723 = PackDaily.TextBox20.value
        PackDaily.TextBox20.value = NullString
    Dim ppc6206723 As Integer
        ppc6206723 = 18 'parts per container
    Dim cpp6206723 As Integer
        cpp6206723 = 8 ' containers per pallet
    Dim cycle6206723 As Integer
        cycle6206723 = 40 'cycle time / cavities
    Dim coverage6206723 As Integer
        coverage6206723 = (count6206723 * ppc6206723 * cpp6206723) / (3600 / cycle6206723)
                    
    Dim count6207233 As Integer
        count6207233 = PackDaily.TextBox21.value
        PackDaily.TextBox21.value = NullString
    Dim ppc6207233 As Integer
        ppc6207233 = 12 'parts per container
    Dim cpp6207233 As Integer
        cpp6207233 = 12 ' containers per pallet
    Dim cycle6207233 As Integer
        cycle6207233 = 40 'cycle time / cavities
    Dim coverage6207233 As Integer
        coverage6207233 = (count6207233 * ppc6207233 * cpp6207233) / (3600 / cycle6207233)
                    
    Dim count551 As Integer
        count551 = PackDaily.TextBox22.value
        PackDaily.TextBox22.value = NullString
    Dim ppc551 As Integer
        ppc551 = 6 'parts per container
    Dim cpp551 As Integer
        cpp551 = 1 ' containers per pallet
    Dim cycle551 As Integer
        cycle551 = 65 / 1 'cycle time / cavities
    Dim coverage551 As Integer
        coverage551 = (count551 * ppc551 * cpp551) / (3600 / cycle551)
    
    Dim count3100662 As Integer
        count3100662 = PackDaily.TextBox23.value
        PackDaily.TextBox23.value = NullString
    Dim ppc3100662 As Integer
        ppc3100662 = 1 'parts per container
    Dim cpp3100662 As Integer
        cpp3100662 = 1 ' containers per pallet
    Dim cycle3100662 As Integer
        cycle3100662 = 60 'cycle time / cavities
    Dim coverage3100662 As Integer
        coverage3100662 = (count3100662 * ppc3100662 * cpp3100662) / (3600 / cycle3100662)
                    
    Dim count6200856 As Integer
        count6200856 = PackDaily.TextBox24.value
        PackDaily.TextBox24.value = NullString
    Dim ppc6200856 As Integer
        ppc6200856 = 26 'parts per container
    Dim cpp6200856 As Integer
        cpp6200856 = 24 ' containers per pallet
    Dim cycle6200856 As Integer
        cycle6200856 = 60 / 3 'cycle time / cavities
    Dim coverage6200856 As Integer
        coverage6200856 = (count6200856 * ppc6200856 * cpp6200856) / (3600 / cycle6200856)
                                                           
    Dim countGb As Integer
        countGb = PackDaily.TextBox25.value
        PackDaily.TextBox25.value = NullString
    Dim ppcGb As Integer
        ppcGb = 24 'parts per container
    Dim cppGb As Integer
        cppGb = 1 ' containers per pallet
    Dim cycleGb As Integer
        cycleGb = 70 / 2 'cycle time / cavities
    Dim coverageGb As Integer
        coverageGb = (countGb * ppcGb * cppGb) / (3600 / cycleGb)
'////LOGGING ON TO CMS
    Application.ScreenUpdating = False
    Dim wkb As Workbook
    Set wkb = Workbooks.Open("\\RE2VMFIL02\Public\Pub-LOGISTICS\Shift Folder New\2019\Customer\Redditch 1-CMS Count Log.xlsm")
    wkb.Sheets("Counts").Select
    For i = 5 To 400
        If Cells(6, i).value = Date Then
            Cells(8, i).value = count3083480
            Cells(9, i).value = count3084151
            Cells(10, i).value = count3084308
            Cells(11, i).value = count3084309
            Cells(12, i).value = count3084310
            Cells(13, i).value = count3084317
            Cells(14, i).value = count3084320
            Cells(15, i).value = count3103210
            Cells(16, i).value = count3104802
            Cells(17, i).value = count6200294
            Cells(18, i).value = count6200852
            Cells(19, i).value = count6200854
            Cells(20, i).value = count6200858
            Cells(21, i).value = count6202426
            Cells(22, i).value = count6202434
            Cells(23, i).value = count6202789
            Cells(24, i).value = count6202790
            Cells(25, i).value = count6203991
            Cells(26, i).value = count6204136
            Cells(27, i).value = count6206723
            Cells(28, i).value = count6207233
            'Cells(29, i).value = count551
            Cells(30, i).value = count3100662
            Cells(31, i).value = count6200856
            Cells(32, i).value = countGb
        End If
    Next i
    wkb.Save
    wkb.Close
    
    
    
    Dim name As String
    name = ThisWorkbook.name
    Workbooks(name).Activate
    Application.ScreenUpdating = True
        
                       
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
    
EmailBody = "<BODY style=font-size:10pt;font-family:Calibri; font-color :#3d3d40;><h3>Redditch 1 Daily Packaging Status:</h3>"

Dim newBody
Dim flag As Integer
flag = 24

newBody = " <html><head><style>body {color: #3d3d40;font-size:10pt;font-family:Calibri;}; <style>table, th, td {border: 1px solid #3d3d40;border-collapse: collapse;text-align: center;}</style></style></head><body>"
newBody = newBody + "<h3> " & Cells(1, 4).value & " Redditch 1 Daily Packaging Report</h3>"

newBody = newBody & "These are indicative data only and may contain errors. Cycle times are not yet verified. Please flag up any concerns. <br><br>"

newBody = newBody & "<table border=""1"" cellspacing=""0"" cellpadding=""0"" style=font-size:10pt;font-family:Calibri; border-collapse: collapse; text-align:center;>"
newBody = newBody & "<tr><b>"

newBody = newBody & "<td align=""center"">&nbsp;Packaging number&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Description&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Full pallets counted&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Production coverage (hrs)&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Standard cycle time&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Parts per container&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Containers per pallet&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Parts per pallet&nbsp;</td></b></tr>"

'///////// 3083480
newBody = newBody & "<tr>"

newBody = newBody & "<td align=""center"">&nbsp;3083480&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;LU Covering tailgate lower inner lower&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count3083480 & "&nbsp;</td>"
    If coverage3083480 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage3083480 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage3083480 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle3083480 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3083480 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp3083480 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3083480 * cpp3083480 & "&nbsp;</td></tr>"


'///////// 3084151
newBody = newBody & "<tr>"

newBody = newBody & "<td align=""center"">&nbsp;3084151&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;F5X Glovebox Assy&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count3084151 & "&nbsp;</td>"
    If coverage3084151 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage3084151 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage3084151 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle3084151 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3084151 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp3084151 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3084151 * cpp3084151 & "&nbsp;</td></tr>"

'///////// 3084308
newBody = newBody & "<tr>"

newBody = newBody & "<td align=""center"">&nbsp;3084308&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;LU -Finisher B - Post inner Upper&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count3084308 & "&nbsp;</td>"
    If coverage3084308 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage3084308 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage3084308 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle3084308 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3084308 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp3084308 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3084308 * cpp3084308 & "&nbsp;</td></tr>"

'///////// 3084309
newBody = newBody & "<tr>"

newBody = newBody & "<td align=""center"">&nbsp;3084309&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Ambient Vent&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count3084309 & "&nbsp;</td>"
    If coverage3084309 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage3084309 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage3084309 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle3084309 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3084309 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp3084309 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3084309 * cpp3084309 & "&nbsp;</td></tr>"
   
   
'///////// 3084310
newBody = newBody & "<tr>"

newBody = newBody & "<td align=""center"">&nbsp;3084310&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;LU - Finisher C Post inner&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count3084310 & "&nbsp;</td>"
    If coverage3084310 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage3084310 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage3084310 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle3084310 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3084310 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp3084310 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3084310 * cpp3084310 & "&nbsp;</td></tr>"
    
'///////// 3084317
newBody = newBody & "<tr>"

newBody = newBody & "<td align=""center"">&nbsp;3084317&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;LU - Side cover instrument Panel&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count3084317 & "&nbsp;</td>"
    If coverage3084317 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage3084317 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage3084317 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle3084317 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3084317 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp3084317 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3084317 * cpp3084317 & "&nbsp;</td></tr>"
    
    
'///////// 3084320
newBody = newBody & "<tr>"

newBody = newBody & "<td align=""center"">&nbsp;3084320&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;LU - Trim tailgate / rear Hatch upper&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count3084320 & "&nbsp;</td>"
    If coverage3084320 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage3084320 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage3084320 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle3084320 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3084320 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp3084320 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3084320 * cpp3084320 & "&nbsp;</td></tr>"
    
   
'///////// 6200294
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6200294&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;F55 BLENDE B-SÄULE INNEN OBEN&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6200294 & "&nbsp;</td>"
    If coverage6200294 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6200294 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6200294 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6200294 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6200294 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6200294 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6200294 * cpp6200294 & "&nbsp;</td></tr>"
    
    
'///////// 6200852
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6200852&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Function carrier centre stack&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6200852 & "&nbsp;</td>"
    If coverage6200852 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6200852 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6200852 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6200852 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6200852 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6200852 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6200852 * cpp6200852 & "&nbsp;</td></tr>"
    
'///////// 6200854
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6200854&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Trim Panel DS- LWR /Footwell Knee Bolst&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6200854 & "&nbsp;</td>"
    If coverage6200854 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6200854 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6200854 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6200854 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6200854 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6200854 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6200854 * cpp6200854 & "&nbsp;</td></tr>"
   
'///////// 6200858
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6200858&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;TRIM INNER TAILGATE / REAR HATCH SIDE&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6200858 & "&nbsp;</td>"
    If coverage6200858 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6200858 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6200858 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6200858 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6200858 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6200858 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6200858 * cpp6200858 & "&nbsp;</td></tr>"
    
'///////// 6202426
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6202426&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;F55 Finisher D-Pillar inner upper&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6202426 & "&nbsp;</td>"
    If coverage6202426 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6202426 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6202426 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6202426 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6202426 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6202426 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6202426 * cpp6202426 & "&nbsp;</td></tr>"
    
'///////// 6202434
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6202434&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;F55 Trim inner tailgate/ rear hatch side&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6202434 & "&nbsp;</td>"
    If coverage6202434 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6202434 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6202434 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6202434 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6202434 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6202434 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6202434 * cpp6202434 & "&nbsp;</td></tr>"
   
'///////// 6202789
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6202789&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Finisher C Post inner LU&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6202789 & "&nbsp;</td>"
    If coverage6202789 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6202789 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6202789 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6202789 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6202789 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6202789 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6202789 * cpp6202789 & "&nbsp;</td></tr>"
   
'///////// 6202790
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6202790&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;F54 Finisher D-Pillar Upper&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6202790 & "&nbsp;</td>"
    If coverage6202790 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6202790 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6202790 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6202790 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6202790 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6202790 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6202790 * cpp6202790 & "&nbsp;</td></tr>"
   
'///////// 6203991
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6203991&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;F54 Glovebox Assembly&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6203991 & "&nbsp;</td>"
    If coverage6203991 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6203991 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6203991 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6203991 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6203991 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6203991 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6203991 * cpp6203991 & "&nbsp;</td></tr>"
    
'///////// 6204136
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6204136&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;F55 finisher C-Post inner&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6204136 & "&nbsp;</td>"
    If coverage6204136 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6204136 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6204136 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6204136 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6204136 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6204136 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6204136 * cpp6204136 & "&nbsp;</td></tr>"
   
   
'///////// 6206723
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6206723&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;F54 Hevac Cover&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6206723 & "&nbsp;</td>"
    If coverage6206723 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6206723 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6206723 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6206723 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6206723 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6206723 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6206723 * cpp6206723 & "&nbsp;</td></tr>"
   
'///////// 6207233
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6207233&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;F54 Hevac Cover&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6207233 & "&nbsp;</td>"
    If coverage6207233 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6207233 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6207233 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6207233 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6207233 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6207233 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6207233 * cpp6207233 & "&nbsp;</td></tr>"
   
   
'///////// 3100662
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;3100662&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Green Gitter Bin&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count3100662 & "&nbsp;</td>"
    If coverage3100662 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage3100662 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage3100662 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle3100662 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3100662 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp3100662 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc3100662 * cpp3100662 & "&nbsp;</td></tr>"
  

'///////// 6200856
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6200856&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;F5X Media Finisher</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count6200856 & "&nbsp;</td>"
    If coverage6200856 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage6200856 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage6200856 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle6200856 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6200856 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp6200856 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc6200856 * cpp6200856 & "&nbsp;</td></tr>"
  
  
'///////// gb
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;G/b Housing&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;Glovebox WIP Stillage</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & countGb & "&nbsp;</td>"
    If coverageGb > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverageGb & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverageGb & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycleGb & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppcGb & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cppGb & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppcGb * cppGb & "&nbsp;</td></tr>"
  
'///////// 551
newBody = newBody & "<tr>"
newBody = newBody & "<td align=""center"">&nbsp;6206490&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;BMW Oxfrod IP Stillage</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & count551 & "&nbsp;</td>"
    If coverage551 > flag Then
        newBody = newBody & "<td align=""center""; style='background-color:#9af246'>&nbsp;" & coverage551 & "&nbsp;</td>"
    Else
        newBody = newBody & "<td align=""center""; style='background-color:#FF9673'>&nbsp;" & coverage551 & "&nbsp;</td>"
    End If
newBody = newBody & "<td align=""center"">&nbsp;" & cycle551 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc551 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & cpp551 & "&nbsp;</td>"
newBody = newBody & "<td align=""center"">&nbsp;" & ppc551 * cpp551 & "&nbsp;</td></tr>"
  
  
  
  
  
  
   
   
    
    
    
newBody = newBody & "</table>"
newBody = newBody & "<br>" & "Report version 09.04.2021 - generated on " & Now()
newBody = newBody & "</html>"

sPath = "Y:\Application Data\Microsoft\Signatures\Main.htm"
'sPath = "C:\Users\Liberski, Pawel\AppData\Roaming\Microsoft\Signatures\main.htm"
If Dir(sPath) <> "" Then
    StrSignature = GetSignature(sPath)
Else
    StrSignature = ""
End If

'On Error Resume Next

    With NewMail
        .To = "Vitkauskas, Arnoldas; Hussain, Israr; porebska, agnieska; krol, radek; Leighton, Rebecca; Wootton, Louise; Sim, Alina; Celec, Marek; Pugh, Richard; Duggins, Jamie; Konieczniak, Michal; Sliwa, Karolina; Masood, Tayyab; Marczynski, Krystian"
        .cc = "Liberski, Pawel;  partridge, jamie; bennett, chris; rushton, craig; kaczynska, daria;"
        .BCC = ""
        .Subject = "Redditch 1 Packaging Count - Trial version " & Now()
        '.htmlbody = EmailBody & "</BODY>" & vbNewLine & Signature
        .htmlbody = newBody & "</html></BODY>" '& vbNewLine & Signature
        '.Attachments.Add (Address)
        .display
    End With
    
    'On Error GoTo 0
    Set NewMail = Nothing
    Set OlApp = Nothing
    
PackDaily.Hide
    
End Sub

Private Sub CommandButton2_Click()
    PackDaily.Hide
End Sub

Private Sub CommandButton3_Click()
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.name Like "Text*" Then ctrl = NullString
    Next
    
    Me.Repaint
    
End Sub


Private Sub CommandButton1_Click()

If ComboBox1.value = NullString Then
    MsgBox "You must select site name first.", vbOKOnly, "Site name not selected"
    Exit Sub
End If


answer = MsgBox("Database will be updated, do you wish to continue?", vbOKCancel, "Confirmation required")
If answer = vbCancel Then Exit Sub

'On Error Resume Next
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
    
    
            
    sSQL = "SELECT * FROM PackCount"     'sql query - select * from selected database
    
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
'===========================================
rst!CountDate = Date
rst!Site = ComboBox1.value
rst!AddedBy = Environ("username")

'on error resume next
If TextBox1.value = NullString Then rst!P3083480 = 0 Else rst!P3083480 = TextBox1.value
If TextBox2.value = NullString Then rst!P3084151 = 0 Else rst!P3084151 = TextBox2.value
If TextBox3.value = NullString Then rst!P3084308 = 0 Else rst!P3084308 = TextBox3.value
If TextBox4.value = NullString Then rst!P3084309 = 0 Else rst!P3084309 = TextBox4.value
If TextBox5.value = NullString Then rst!P3084310 = 0 Else rst!P3084310 = TextBox5.value
If TextBox6.value = NullString Then rst!P3084317 = 0 Else rst!P3084317 = TextBox6.value
If TextBox7.value = NullString Then rst!P3084320 = 0 Else rst!P3084320 = TextBox7.value
If TextBox8.value = NullString Then rst!P3103210 = 0 Else rst!P3103210 = TextBox8.value
If TextBox9.value = NullString Then rst!P3104802 = 0 Else rst!P3104802 = TextBox9.value
If TextBox10.value = NullString Then rst!P6200294 = 0 Else rst!P6200294 = TextBox10.value
If TextBox11.value = NullString Then rst!P6200852 = 0 Else rst!P6200852 = TextBox11.value
If TextBox12.value = NullString Then rst!P6200854 = 0 Else rst!P6200854 = TextBox12.value
If TextBox13.value = NullString Then rst!P6200858 = 0 Else rst!P6200858 = TextBox13.value
If TextBox14.value = NullString Then rst!P6202426 = 0 Else rst!P6202426 = TextBox14.value
If TextBox15.value = NullString Then rst!P6202434 = 0 Else rst!P6202434 = TextBox15.value
If TextBox16.value = NullString Then rst!P6202789 = 0 Else rst!P6202789 = TextBox16.value
If TextBox17.value = NullString Then rst!P6202790 = 0 Else rst!P6202790 = TextBox17.value
If TextBox18.value = NullString Then rst!P6203991 = 0 Else rst!P6203991 = TextBox18.value
If TextBox19.value = NullString Then rst!P6204136 = 0 Else rst!P6204136 = TextBox19.value
If TextBox20.value = NullString Then rst!P6206723 = 0 Else rst!P6206723 = TextBox20.value
If TextBox21.value = NullString Then rst!P6207233 = 0 Else rst!P6207233 = TextBox21.value
If TextBox22.value = NullString Then rst!P6206490 = 0 Else rst!P6206490 = TextBox22.value
If TextBox23.value = NullString Then rst!P3100662 = 0 Else rst!P3100662 = TextBox23.value
If TextBox24.value = NullString Then rst!P6200856 = 0 Else rst!P6200856 = TextBox24.value
If TextBox25.value = NullString Then rst!GboxWIP = 0 Else rst!GboxWIP = TextBox25.value

rst.Update
rst.Close

Set rst = Nothing
cnn.Close
Set cnn = Nothing

'=====================================
'On Error Resume Next
Call Com

'=====================================
 
MsgBox "Database was updated.", vbInformation, "Procedure completed."

'=====================================
PackDaily.Repaint
PackDaily.Hide
  End If
End Sub

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Activate()
    On Error GoTo 0
    Debug.Print "page loaded"
    Call isAvailable("J:\Pub-LOGISTICS\Packaging\Packaging.accdb")
    ComboBox1.value = "Redditch 1"
    
    
End Sub
Private Sub UserForm_Click()

End Sub
Private Sub UserForm_Initialize()

    ComboBox1.AddItem "Redditch 1"
    ComboBox1.AddItem "Redditch 2 Main"
    ComboBox1.AddItem "Redditch 2 DS"
    ComboBox1.AddItem "Barton"
    
End Sub
Private Sub isAvailable(ByVal adres As String)
    On Error Resume Next
    If Dir(adres) = NullString Then
        Label2.Caption = "Database availability Status: Unavailable"
        CommandButton1.Enabled = False
    Else
        Label2.Caption = "Database availability Status: Available"
    End If
    Label3.Caption = "User: " & Environ("username")
    CommandButton1.Enabled = True
    
End Sub
