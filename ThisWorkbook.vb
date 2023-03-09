'// ======================================================
'// Code triggered on file opening event
'// Author: Pawel Liberski 2018 - 2023
'// MrLiberski@outlook.com
'// Last modified: 09/03/2022 (tidy up mess code and removed redundant stuff)
'// ======================================================
Private Sub Workbook_Open()

    Call setTimesheetAccess()
    Call hideSheets()
    Call SetScreen()
    Call welcomeBit()
    Call activityLogger() '// This script is stored in separate file
    Call fffcaller()      '// calling external reference file

End sub
'// ======================================================

'// resizes widow to full screen etc
Private sub SetScreen()

    With Application
        .DisplayFullScreen = True
        .CommandBars("Full Screen").Visible = False
        .ScreenUpdating = True
        .ActiveWindow.DisplayHeadings = False
        .EnableEvents = True
    End With

End Sub
'// ======================================================

'//  I was playing with tex-to-speech but have settled for welcome message for me only
sub welcomeBit()

    dim welcomeMessage
    mess = "Hi"

    If UCase(Environ("UserName")) = "PAWEL" Then
        'Application.Speech.Speak mess, True, True, True
        MsgBox "This is the version of document updated on 1st June 2020.", vbOKOnly, "Authorized successfully."
    End If

end sub
'// ======================================================


'//  Ensure timesheet sheet is hidden away from those who should not see it
Sub setTimesheetAccess()

    If UCase(Environ("username")) = "MAREK" Or _
        UCase(Environ("username")) = "PAWEL" Or _
        UCase(Environ("username")) = "MARIA" Then
        Sheets("Timesheet").Visible = xlSheetVisible 'show timesheet to who can access it
    Else
        Sheets("Timesheet").Visible = xlSheetVeryHidden 'hide from unauthorized people
    End If

End sub
'// ======================================================


'// Hide sheets that are better left hidden
Sub hideSheets()

    On Error Resume Next
    Sheets("Sheet1").Visible = xlVeryHidden
    Sheets("Products").Visible = xlVeryHidden
    Sheets("Odette").Visible = xlVeryHidden
    Sheets("Sheet2").Visible = xlVeryHidden
    Sheets("Folder Labels").Visible = xlVeryHidden
        
    Sheets("BRIEF").Unprotect
    Sheets("BRIEF").Cells(57, 2).value = UCase(Environ("username"))
    Sheets("BRIEF").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True

End Sub
'// ======================================================
'// ======================================================
'// ======================================================
'// ======================================================
'// ======================================================