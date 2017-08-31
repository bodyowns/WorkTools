Sub FormatText()
    ' DelShift Rows 1 to 10

'    Rows("1:10").Select
'    Selection.Delete Shift:=xlUp
    ' Analize and find Text Store
Dim x As String
x = "Store"
' Select cell B1
Range("B1").Select
' Set Do loop to stop when Store text in cell is reached.
Do Until ActiveCell.Value = "Store"
    nx = 1
    For nx = 1 To 35:
    
    If ActiveCell.Value = x Or nx = 32 Then

        Exit Do
    End If
    ActiveCell.Offset(1, 0).Select
    Next nx
    Loop
    
' Select range to the top stop

      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Selection.Delete Shift:=xlToLeft
Dim y As String
y = "Pick Face"
' Select cell B1
Range("F1").Select
' Set Do loop to stop when Store text in cell is reached.
Do Until ActiveCell.Value = "Pick Face"
    If ActiveCell.Value = y Then
        Exit Do
    End If
    ActiveCell.Offset(1, 0).Select
    Loop
    
' Select range to the top

      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Selection.Delete Shift:=xlToLeft
    
Dim z As String
z = "Priority"
' Select cell B1
Range("L1").Select
' Set Do loop to stop when Store text in cell is reached.
Do Until ActiveCell.Value = "Priority"
    If ActiveCell.Value = z Then
        Exit Do
    End If
    ActiveCell.Offset(1, 0).Select
    Loop
    
' Select range to the top

      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Range(Selection, Selection.End(xlUp)).Select
      Selection.Delete Shift:=xlToLeft
    


  
    Columns("A:F").Select
    ActiveWindow.SmallScroll Down:=-33
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:AV").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.NumberFormat = "@"
    Columns("B:B").EntireColumn.AutoFit
    Selection.ColumnWidth = 16.14
    Selection.NumberFormat = "0"
    ActiveWindow.SmallScroll Down:=-87
    Cells.Select
    Cells.EntireColumn.AutoFit

End Sub
End Sub