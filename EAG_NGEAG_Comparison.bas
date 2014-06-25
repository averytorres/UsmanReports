

Sub Make_EAG_NGEAG_Comp()

Dim pt As PivotTable
Dim cacheOfpt As PivotCache  'source data for pt
Dim pf As PivotField
Dim pi As PivotItem
Dim WSD As Worksheet
Dim PRange As Range
Dim wsSheet1 As Worksheet

'try to rename current workbook 'Data' if it is not already
On Error Resume Next
Set wsSheet1 = Sheets("EAG_NGEAG_COMP")
On Error GoTo 0
If Not wsSheet1 Is Nothing Then
    Sheets("EAG_NGEAG_COMP").Select ' does exist
Else
    Sheets.Add.Name = "EAG_NGEAG_COMP" 'does not exist
End If


'add text
Sheets("EAG_NGEAG_COMP").Range("A3").Value = "EAG May 2014"
Sheets("EAG_NGEAG_COMP").Range("A3").Font.Bold = True
Sheets("EAG_NGEAG_COMP").Range("A3:B3").Interior.ColorIndex = 16

Sheets("EAG_NGEAG_COMP").Range("A4").Value = "Minimum Volume Day"
Sheets("EAG_NGEAG_COMP").Range("A5").Value = "Maximum Volume Day"
Sheets("EAG_NGEAG_COMP").Range("A6").Value = "Average Volume Day"

Sheets("EAG_NGEAG_COMP").Range("B4").Value = "93,298"
Sheets("EAG_NGEAG_COMP").Range("B5").Value = "446,510"
Sheets("EAG_NGEAG_COMP").Range("B6").Value = "318,069"

Sheets("EAG_NGEAG_COMP").Range("A4:B6").Interior.ColorIndex = 42

Sheets("EAG_NGEAG_COMP").Range("A8").Value = "NGEAG Day-X Transactions"
Sheets("EAG_NGEAG_COMP").Range("A8").Font.Bold = True
Sheets("EAG_NGEAG_COMP").Range("A8:B8").Interior.ColorIndex = 35

Sheets("EAG_NGEAG_COMP").Range("A9").Value = "Volume for XX/XX/XXXX"
Sheets("EAG_NGEAG_COMP").Range("B9").Value = Sheets("Result_SC").Range("L9").Value

Sheets("EAG_NGEAG_COMP").Range("A9:B9").Interior.ColorIndex = 42

Sheets("EAG_NGEAG_COMP").Range("B17").Value = "NGEAG"
Sheets("EAG_NGEAG_COMP").Range("B17").Font.Bold = True
Sheets("EAG_NGEAG_COMP").Range("B17:B18").Interior.ColorIndex = 35

Sheets("EAG_NGEAG_COMP").Range("C17").Value = "EAG"
Sheets("EAG_NGEAG_COMP").Range("C18").Value = "All Days in May 2014"
Sheets("EAG_NGEAG_COMP").Range("C17").Font.Bold = True
Sheets("EAG_NGEAG_COMP").Range("C17:C18").Interior.ColorIndex = 16

Sheets("EAG_NGEAG_COMP").Range("A19").Value = "Min TPS"
Sheets("EAG_NGEAG_COMP").Range("A20").Value = "Max TPS"
Sheets("EAG_NGEAG_COMP").Range("A21").Value = "Average TPS"
Sheets("EAG_NGEAG_COMP").Range("A19:A21").Interior.ColorIndex = 41

Sheets("EAG_NGEAG_COMP").Range("C19").Value = "0.0"
Sheets("EAG_NGEAG_COMP").Range("C20").Value = "26.0"
Sheets("EAG_NGEAG_COMP").Range("C21").Value = ".49"

Sheets("EAG_NGEAG_COMP").Range("B19").Value = Sheets("Result_TPS").Range("M20").Value
Sheets("EAG_NGEAG_COMP").Range("B19").Value = Sheets("Result_TPS").Range("M20").Value
Sheets("EAG_NGEAG_COMP").Range("B20").Value = Sheets("Result_TPS").Range("N20").Value
avg = (Sheets("Result_TPS").Range("O20").Value + Sheets("Result_TPS").Range("O21").Value) / 2
Sheets("EAG_NGEAG_COMP").Range("B21").Value = avg
Sheets("EAG_NGEAG_COMP").Range("B21").NumberFormat = "0.00%"
Sheets("EAG_NGEAG_COMP").Range("B19:C21").Interior.ColorIndex = 42

Worksheets("EAG_NGEAG_COMP").Columns("A:C").AutoFit

Dim MyChart As Chart
Dim MyRange As Range
Dim LastRow As Long
        
Set MyRange = Range("A4:B4" & ",A9:B9" & ",A5:B6")
Set MyChart = ActiveSheet.Shapes.AddChart(xlColumnClustered).Chart
MyChart.SetSourceData Source:=MyRange
MyChart.SeriesCollection(1).Name = Range("K26").Value
MyChart.Legend.Delete
MyChart.ClearToMatchStyle
MyChart.SeriesCollection([1]).Points([1]).Interior.Color = RGB(255, 165, 0)
MyChart.SeriesCollection([1]).Points([2]).Interior.Color = RGB(154, 205, 50)
MyChart.SeriesCollection([1]).Points([3]).Interior.Color = RGB(34, 139, 34)
MyChart.SeriesCollection([1]).Points([4]).Interior.Color = RGB(0, 191, 255)

With MyChart.Parent
         .Height = 230 ' resize
         .Width = 305  ' resize
         .Top = 10    ' reposition
         .Left = 171   ' reposition
End With


End Sub

