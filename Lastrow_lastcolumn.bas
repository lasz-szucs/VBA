'Find used range: last row and last column in the table

Public lastrow As Integer
Public lastcol As Integer

With ActiveSheet
        lastrow = .UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
        lastcol = .UsedRange.Columns(ActiveSheet.UsedRange.Columns.Count).Column
End With