Attribute VB_Name = "examples"
'=================================================================================
'This module contains code examples only
'=================================================================================
'This code will unhide all sheets in the workbook
Sub UnhideAllWoksheets()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Visible = xlSheetVisible
Next ws
End Sub
'=================================================================================
'This macro will hide all the worksheet except the active sheet
Sub HideAllExceptActiveSheet()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
If ws.name <> ActiveSheet.name Then ws.Visible = xlSheetHidden
Next ws
End Sub
'=================================================================================
'This code will sort the worksheets alphabetically
Sub SortSheetsTabName()
Application.ScreenUpdating = False
Dim ShCount As Integer, i As Integer, j As Integer
ShCount = Sheets.Count
For i = 1 To ShCount - 1
For j = i + 1 To ShCount
If Sheets(j).name < Sheets(i).name Then
Sheets(j).Move before:=Sheets(i)
End If
Next j
Next i
Application.ScreenUpdating = True
End Sub
'=================================================================================
'This code will protect all the sheets at one go
Sub ProtectAllSheets()
Dim ws As Worksheet
Dim password As String
password = "Test123" 'replace Test123 with the password you want
For Each ws In Worksheets
   ws.Protect password:=password
Next ws
End Sub
'=================================================================================
'This code will protect all the sheets at one go
Sub Protect_AllSheets()
Dim ws As Worksheet
Dim password As String
password = "Test123" 'replace Test123 with the password you want
For Each ws In Worksheets
ws.Unprotect password:=password
Next ws
End Sub
'=================================================================================
'This code will unmerge all the merged cells
Sub UnmergeAllCells()
ActiveSheet.Cells.UnMerge
End Sub
'=================================================================================
'This code will Save the File With a Timestamp in its name
Sub SaveWorkbookWithTimeStamp()
Dim timestamp As String
timestamp = Format(Date, "dd-mm-yyyy") & "_" & Format(Time, "hh-ss")
ThisWorkbook.SaveAs "C:UsersUsernameDesktopWorkbookName" & timestamp
End Sub
'=================================================================================
'This code will save each worsheet as a separate PDF
Sub SaveWorkshetAsPDF()
Dim ws As Worksheet
For Each ws In Worksheets
ws.ExportAsFixedFormat xlTypePDF, "C:UsersSumitDesktopTest" & ws.name & ".pdf"
Next ws
End Sub
'=================================================================================
'This code will save the entire workbook as PDF
Sub SaveWorkshetxAsPDF()
ThisWorkbook.ExportAsFixedFormat xlTypePDF, "C:UsersSumitDesktopTest" & ThisWorkbook.name & ".pdf"
End Sub
'=================================================================================
'This code will convert all formulas into values
Sub ConvertToValues()
With ActiveSheet.UsedRange
.value = .value
End With
End Sub
'=================================================================================
'This macro code will lock all the cells with formulas
Sub LockCellsWithFormulas()
With ActiveSheet
   .Unprotect
   .Cells.Locked = False
   .Cells.SpecialCells(xlCellTypeFormulas).Locked = True
   .Protect AllowDeletingRows:=True
End With
End Sub
'=================================================================================
'This code will protect all sheets in the workbook
Sub ProtectxAllSheets()
Dim ws As Worksheet
For Each ws In Worksheets
ws.Protect
Next ws
End Sub
'=================================================================================
'This code will insert a row after every row in the selection
Sub InsertAlternateRows()
Dim rng As Range
Dim CountRow As Integer
Dim i As Integer
Set rng = Selection
CountRow = rng.EntireRow.Count
For i = 1 To CountRow
ActiveCell.EntireRow.Insert
ActiveCell.Offset(2, 0).Select
Next i
End Sub
'=================================================================================
'This code will insert a timestamp in the adjacent cell
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo Handler
If Target.Column = 1 And Target.value <> "" Then
Application.EnableEvents = False
Target.Offset(0, 1) = Format(Now(), "dd-mm-yyyy hh:mm:ss")
Application.EnableEvents = True
End If
Handler:
End Sub
'=================================================================================
'This code would highlight alternate rows in the selection
Sub HighlightAlternateRows()
Dim Myrange As Range
Dim Myrow As Range
Set Myrange = Selection
For Each Myrow In Myrange.Rows
   If Myrow.Row Mod 2 = 1 Then
      Myrow.Interior.color = vbCyan
   End If
Next Myrow
End Sub
'=================================================================================
'This code will highlight the cells that have misspelled words
Sub HighlightMisspelledCells()
Dim cl As Range
For Each cl In ActiveSheet.UsedRange
If Not Application.CheckSpelling(Word:=cl.text) Then
cl.Interior.color = vbRed
End If
Next cl
End Sub
'=================================================================================
'This code will refresh all the Pivot Table in the Workbook
Sub RefreshAllPivotTables()
Dim PT As PivotTable
For Each PT In ActiveSheet.PivotTables
PT.RefreshTable
Next PT
End Sub
'=================================================================================
'This code will change the Selection to Upper Case
Sub ChangeCase()
Dim rng As Range
For Each rng In Selection.Cells
If rng.HasFormula = False Then
rng.value = UCase(rng.value)
End If
Next rng
End Sub
'=================================================================================
'This code will highlight cells that have comments`
Sub HighlightCellsWithComments()
ActiveSheet.Cells.SpecialCells(xlCellTypeComments).Interior.color = vbBlue
End Sub
'=================================================================================
'This code will highlight all the blank cells in the dataset
Sub HighlightBlankCells()
Dim Dataset As Range
Set Dataset = Selection
Dataset.SpecialCells(xlCellTypeBlanks).Interior.color = vbRed
End Sub
'=================================================================================
Sub SortDataHeader()
Range("DataRange").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
End Sub
'=================================================================================
Sub SortMultipleColumns()
With ActiveSheet.Sort
 .SortFields.Add key:=Range("A1"), Order:=xlAscending
 .SortFields.Add key:=Range("B1"), Order:=xlAscending
 .SetRange Range("A1:C13")
 .Header = xlYes
 .Apply
End With
End Sub
'=================================================================================
'This VBA code will create a function to get the numeric part from a string
Function GetNumeric(CellRef As String)
Dim StringLength As Integer
StringLength = Len(CellRef)
For i = 1 To StringLength
If IsNumeric(Mid(CellRef, i, 1)) Then result = result & Mid(CellRef, i, 1)
Next i
GetNumeric = result
End Function
'=================================================================================
'This VBA code will create a function to get the text part from a string
Function GetText(CellRef As String)
Dim StringLength As Integer
StringLength = Len(CellRef)
For i = 1 To StringLength
If Not (IsNumeric(Mid(CellRef, i, 1))) Then result = result & Mid(CellRef, i, 1)
Next i
GetText = result
End Function
'=================================================================================
Sub AddSerialNumbers()
Dim i As Integer
On Error GoTo Last
i = InputBox("Enter Value", "Enter Serial Numbers")
For i = 1 To i
ActiveCell.value = i
ActiveCell.Offset(1, 0).Activate
Next i
Last: Exit Sub
End Sub
'=================================================================================
'2. Insert Multiple Columns
'Once you run this macro it will show an input box and you need to enter the number of columns you want to insert.
Sub InsertMultipleColumns()
Dim i As Integer
Dim j As Integer
ActiveCell.EntireColumn.Select
On Error GoTo Last
i = InputBox("Enter number of columns to insert", "Insert Columns")
For j = 1 To i
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightorAbove Next j
Last: Exit Sub
End Sub
'=================================================================================
Sub OpenCalculator()
Application.ActivateMicrosoftApp Index:=0
End Sub
'=================================================================================
Sub PasteAsPicture() ' copy selection as image
Application.CutCopyMode = False
Selection.Copy
ActiveSheet.Pictures.Paste.Select
End Sub
'=================================================================================
Sub Speak()
Selection.Speak
End Sub
'=================================================================================
Sub talktome()
Dim mess As String
mess = "Good morning satan, pleased to meet you"
Application.Speech.Speak (mess)

End Sub
