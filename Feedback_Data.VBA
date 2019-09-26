'Here is the macros code needed to automate any of the spreadsheets in the feedback sections

' name: deleteColumn
' in: n/a
' out: n/a
' description: deletes rows A:H, J, and L:O
'		(columns unneeded for scraping the data)

Sub deleteColumn()
Range("A:H, J:J, L:O").Delete
End Sub


' name: deleteRow
' in: n/a
' out: n/a
' description: deletes headers in first row, and any blank cells in column A

Sub deleteRow()
Rows(1).Font.Bold = True
On Error Resume Next
Columns("A").SpecialCells(xlBlanks).EntireRow.Delete
End Sub


' name: Sort
' in: n/a
' out: n/a
' description: first links each row in the columns a b e, and then sorts A alphabetically

Sub Sort()
'Sort all the data
With Range("A1").CurrentRegion
.Sort Key1:=Range("A1"), Key2:=Range("B1"), Header:=xlYes
End With
End Sub


' name: countOccurrences
' in: n/a
' out: n/a
' description: counts instances of each rating, and then displays a series of 5 messages boxes displaying the rankings

Sub countOccurrences()
'Declare variables
Dim venjoyCount As Long, enjoyCount As Long, mehCount As Long, notenjoyCount As Long, badCount As Long
venjoyCount = Application.CountIf(Range("A:A"), "Very Enjoyable")
enjoyCount = Application.CountIf(Range("A:A"), "Enjoyable")
mehCount = Application.CountIf(Range("A:A"), "Meh")
notenjoyCount = Application.CountIf(Range("A:A"), "Not Enjoyable")
badCount = Application.CountIf(Range("A:A"), "It was quite bad...")

'Message Boxes, from best to worst
MsgBox ("Very Enjoyable: " & venjoyCount)
MsgBox ("Enjoyable: " & enjoyCount)
MsgBox ("Meh: " & mehCount)
MsgBox ("Not Enjoyable: " & notenjoyCount)
MsgBox ("Bad: " & badCount)
End Sub



`<<<MAIN START>>>

Sub Run()
Call deleteColumn
Call deleteRow
Call Sort
Call countOccurrences
End Sub