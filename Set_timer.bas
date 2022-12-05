'Measure macro running time with a pop up text box with results at the end

Public StartTime As Double
Public MinutesElapsed As String

'Start timer in the first line

StartTime = Timer

'Display result in the last line

MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
MsgBox "Successful run in " & MinutesElapsed & " minutes!", vbInformation