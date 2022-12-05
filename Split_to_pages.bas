'Split sheet content to new sheets based on unique values in a column

Public splitsheet As String
Public lr As Long
Public lc As Long
Public fr As Long
Public fc As Long
Public ws As Worksheet
Public wb As Workbook
Public splitcol, i As Integer
Public icol As Long
Public vcol As Long
Public myarr As Variant
Public title As String


'Name of the sheet to split
splitsheet = "Sheet1"

'Header range with columns to split
title = "B2:D2"

'Column relative position in table to split
splitcol = 2

Set ws = Sheets(splitsheet)

fr = ws.Range(title).Cells(1).Row
fc = ws.Range(title).Cells(1).Column
lr = ws.Cells(ws.Rows.Count, fc).End(xlUp).Row
lc = ws.Cells(fr, ws.Columns.Count).End(xlToLeft).Column
icol = ws.Columns.Count
ws.Cells(fr, icol) = "Unique"
vcol = splitcol + fc - 1

'Replace blank values with "(blank)" text in the splitted column
ws.Range(Cells(fr, vcol), Cells(lr, vcol)).Select
Selection.Replace What:="", Replacement:="(blank)", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

For i = fr + 1 To lr
On Error Resume Next
If ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(ws.Cells(i, vcol), ws.Columns(icol), 0) = 0 Then
ws.Cells(ws.Rows.Count, icol).End(xlUp).Offset(1) = ws.Cells(i, vcol)
End If
Next

myarr = Application.WorksheetFunction.Transpose(ws.Columns(icol).SpecialCells(xlCellTypeConstants))
ws.Columns(icol).Clear

For i = 2 To UBound(myarr)
ws.Range(title).AutoFilter field:=splitcol, Criteria1:=myarr(i) & ""
If Not Evaluate("=ISREF('" & myarr(i) & "'!A1)") Then
Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = Left(myarr(i), 20) & ""
Else
Sheets(Left(myarr(i), 20) & "").Move After:=Worksheets(Worksheets.Count)
End If

ws.Select
ws.Range(Cells(fr, fc), Cells(lr, lc)).Copy Sheets(Left(myarr(i), 20) & "").Range("A1")
Sheets(Left(myarr(i), 20) & "").Columns.AutoFit

Next
ws.AutoFilterMode = False
ws.Activate
End Sub

