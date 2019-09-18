'Here is the macros code needed to automate any of the spreadsheets in the feedback sections

Sub deleteColumn()
Range("A:H, J:J, L:O").Delete
End Sub
Sub deleteRow()
Rows(1).Font.Bold = True
On Error Resume Next
Columns("A").SpecialCells(xlBlanks).EntireRow.Delete

End Sub
Sub Sort()
'Sort all the data
With Range("A1").CurrentRegion
.Sort Key1:=Range("A1"), Key2:=Range("B1"), Header:=xlYes
End With
End Sub
Sub countOccurrences()

'Declare variables
Dim venjoyCount As Long, enjoyCount As Long, mehCount As Long, notenjoyCount As Long, badCount As Long

venjoyCount = Application.CountIf(Range("A:A"), "Very Enjoyable")
enjoyCount = Application.CountIf(Range("A:A"), "Enjoyable")
mehCount = Application.CountIf(Range("A:A"), "Meh")
notenjoyCount = Application.CountIf(Range("A:A"), "Not Enjoyable")
badCount = Application.CountIf(Range("A:A"), "It was quite bad...")

'Message Boxes
MsgBox ("Very Enjoyable: " & venjoyCount)
MsgBox ("Enjoyable: " & enjoyCount)
MsgBox ("Meh: " & mehCount)
MsgBox ("Not Enjoyable: " & notenjoyCount)
MsgBox ("Bad: " & badCount)
End Sub
Sub Run()
Call deleteColumn
Call deleteRow
Call Sort
Call countOccurrences
End Sub