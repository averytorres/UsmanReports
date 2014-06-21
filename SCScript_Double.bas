
Sub MakeSCPivot_Double()

Dim pt As PivotTable
Dim cacheOfpt As PivotCache  'source data for pt
Dim pf As PivotField
Dim pi As PivotItem
Dim WSD As Worksheet
Dim PRange As Range
Dim wsSheet1 As Worksheet

'try to rename the workbook with data into 'Data', if it already exists use it
On Error Resume Next
Set wsSheet1 = Sheets("Allen")
On Error GoTo 0
If Not wsSheet1 Is Nothing Then
    Sheets("Allen").Select ' does exist
Else
    ActiveSheet.Name = "Allen" 'does not exist
End If


'get rage of used cells in data workbook
Set WSD = Worksheets("Allen")
FinalRow = WSD.Cells(Application.Rows.Count, 1).End(xlUp).Row
FinalCol = WSD.Cells(1, Application.Columns.Count). _
    End(xlToLeft).Column
Set PRange = WSD.Cells(1, 1).Resize(FinalRow, FinalCol)

'try to create a workbok named 'Result' if it already exists use it
Dim wsSheet As Worksheet
On Error Resume Next
Set wsSheet = Sheets("Allen_Result")
On Error GoTo 0
If Not wsSheet Is Nothing Then
    Sheets("Allen_Result").Select ' does exist
Else
    Sheets.Add.Name = "Allen_Result" 'does not exist
End If

'delete old pivot table if it exists
On Error Resume Next
ActiveSheet.PivotTables("MyPT").TableRange2.Clear 'delete old pt

'set cache of pt
Sheets("Allen").Select
Set cacheOfpt = ActiveWorkbook.PivotCaches.Create(xlDatabase, SourceData:=PRange)

'creat pt
Sheets("Allen_Result").Select
Set pt = ActiveSheet.PivotTables.Add(cacheOfpt, Range("a1"), "MyPT")

'put fields in
With pt
    'add fields
    .PivotFields(" Time").Orientation = xlRowField
    .PivotFields(" Node").Orientation = xlPageField
    .PivotFields(" Status").Orientation = xlColumnField
    .PivotFields(" Count").Orientation = xlDataField
End With

'set correct nodes to be displayed
Set pf = pt.PivotFields(" Node")
With pf
    For Each pi In pf.PivotItems
        If pi.Name Like " AES_*" Then
        pi.Visible = True
        Else
        pi.Visible = False
        End If
    Next pi
End With

'rename columns
Sheets("Allen_Result").Range("A3").Value = "Total Transactions"
Sheets("Allen_Result").Range("A4").Value = "Intervals"
Sheets("Allen_Result").Range("B4").Value = "Success"
Sheets("Allen_Result").Range("C4").Value = "Error"
Sheets("Allen_Result").Range("D4").Value = "Total Transactions"
Sheets("Allen_Result").Select

'find the last cell in column to get toal and rate values for result table
Allen_TT = Range("E65536").End(xlUp).Value
Allen_ST = Range("B65536").End(xlUp).Value
Allen_SR = Allen_ST / Allen_TT
Allen_ET = Range("C65536").End(xlUp).Value
Allen_ER = Allen_ET / Allen_TT

'get pivot table size
Set WSD = Worksheets("Allen_Result")
FinalRow = WSD.Cells(Application.Rows.Count, 1).End(xlUp).Row
FinalCol = WSD.Cells(1, Application.Columns.Count). _
    End(xlToLeft).Column
Set PRange = WSD.Cells(1, 1).Resize(FinalRow, FinalCol)

'create the end result worksheet
Dim wsSheetResult As Worksheet
Set wsSheetResult = Sheets("Result")
On Error GoTo 0
If Not wsSheetResult Is Nothing Then
    Sheets("Result").Select ' does exist
Else
    Sheets.Add.Name = "Result" 'does not exist
End If
Sheets("Result").Select

'add text
Sheets("Result").Range("A2").Value = "Allen Datacenter"
Sheets("Result").Range("A2").Font.Bold = True
Sheets("Result").Range("A2:G2").Interior.ColorIndex = 15
Sheets("Result").Range("B17").Value = "Centeral Time"
Sheets("Result").Range("B17").Font.Bold = True
Sheets("Result").Range("B17:C17").Interior.ColorIndex = 15

'create line chart
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=PRange
ActiveChart.ChartType = xlLine
ActiveChart.SetElement (msoElementLegendNone)
With ActiveChart.Parent
         .Height = 200 ' resize
         .Width = 335  ' resize
         .Top = 30    ' reposition
         .Left = 0   ' reposition
End With

'------------------------------------------------------------------------------------------

'try to rename the workbook with data into 'Data', if it already exists use it
On Error Resume Next
Set wsSheet1 = Sheets("Bothell")
On Error GoTo 0
If Not wsSheet1 Is Nothing Then
    Sheets("Bothell").Select ' does exist
Else
    ActiveSheet.Name = "Bothell" 'does not exist
End If

'get rage of used cells in data workbook
Set WSD = Worksheets("Bothell")
FinalRow = WSD.Cells(Application.Rows.Count, 1).End(xlUp).Row
FinalCol = WSD.Cells(1, Application.Columns.Count). _
    End(xlToLeft).Column
Set PRange = WSD.Cells(1, 1).Resize(FinalRow, FinalCol)

'try to create a workbok named 'Result' if it already exists use it
Dim wsSheet2 As Worksheet
On Error Resume Next
Set wsSheet2 = Sheets("Bothell_Result")
On Error GoTo 0
If Not wsSheet2 Is Nothing Then
    Sheets("Bothell_Result").Select ' does exist
Else
    Sheets.Add.Name = "Bothell_Result" 'does not exist
End If

'delete old pivot table if it exists
On Error Resume Next
ActiveSheet.PivotTables("MyPT").TableRange2.Clear 'delete old pt

'set cache of pt
Sheets("Bothell").Select
Set cacheOfpt = ActiveWorkbook.PivotCaches.Create(xlDatabase, SourceData:=PRange)

'creat pt
Sheets("Bothell_Result").Select
Set pt = ActiveSheet.PivotTables.Add(cacheOfpt, Range("a1"), "MyPT")

'put fields in
With pt
    'add fields
    .PivotFields(" Time").Orientation = xlRowField
    .PivotFields(" Node").Orientation = xlPageField
    .PivotFields(" Status").Orientation = xlColumnField
    .PivotFields(" Count").Orientation = xlDataField
End With

'set correct nodes to be displayed
Set pf = pt.PivotFields(" Node")
With pf
    For Each pi In pf.PivotItems
        If pi.Name Like " AES_*" Then
        pi.Visible = True
        Else
        pi.Visible = False
        End If
    Next pi
End With

'rename columns
Sheets("Bothell_Result").Range("A3").Value = "Total Transactions"
Sheets("Bothell_Result").Range("A4").Value = "Intervals"
Sheets("Bothell_Result").Range("B4").Value = "Success"
Sheets("Bothell_Result").Range("C4").Value = "Error"
Sheets("Bothell_Result").Range("D4").Value = "Total Transactions"
Sheets("Bothell_Result").Select

'find the last cell in column to get toal and rate values for result table
Bothell_TT = Range("E65536").End(xlUp).Value
Bothell_ST = Range("B65536").End(xlUp).Value
Bothell_SR = Allen_ST / Allen_TT
Bothell_ET = Range("C65536").End(xlUp).Value
Bothell_ER = Allen_ET / Allen_TT

'get pivot table size
Set WSD = Worksheets("Bothell_Result")
FinalRow = WSD.Cells(Application.Rows.Count, 1).End(xlUp).Row
FinalCol = WSD.Cells(1, Application.Columns.Count). _
    End(xlToLeft).Column
Set PRange = WSD.Cells(1, 1).Resize(FinalRow, FinalCol)

Set wsSheetResult = Sheets("Result")
On Error GoTo 0
If Not wsSheetResult Is Nothing Then
    Sheets("Result").Select ' does exist
Else
    Sheets.Add.Name = "Result" 'does not exist
End If
Sheets("Result").Select

'add text
Sheets("Result").Range("A19").Value = "Bothell Datacenter"
Sheets("Result").Range("A19").Font.Bold = True
Sheets("Result").Range("A19:G19").Interior.ColorIndex = 37
Sheets("Result").Range("B34").Value = "Pacific Time"
Sheets("Result").Range("B34").Font.Bold = True
Sheets("Result").Range("B34:C34").Interior.ColorIndex = 37

'create line chart
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=PRange
ActiveChart.ChartType = xlLine
ActiveChart.SetElement (msoElementLegendNone)
With ActiveChart.Parent
         .Height = 200 ' resize
         .Width = 335  ' resize
         .Top = 285    ' reposition
         .Left = 0   ' reposition
End With



'-----------------------------------
'Draw chart to right

Columns("I").ColumnWidth = 20
Sheets("Result").Range("J8").Value = "Allen"
Sheets("Result").Range("J8").Font.Bold = True
Sheets("Result").Range("J8:J13").Interior.ColorIndex = 15

Sheets("Result").Range("K8").Value = "Bothell"
Sheets("Result").Range("K8").Font.Bold = True
Sheets("Result").Range("K8:K13").Interior.ColorIndex = 37

Sheets("Result").Range("L8").Value = "Total"
Sheets("Result").Range("L8").Font.Bold = True
Sheets("Result").Range("L8:L13").Interior.ColorIndex = 16

Sheets("Result").Range("I9").Value = "Total Transactions"
Sheets("Result").Range("I9").Font.Bold = True
Sheets("Result").Range("I9").Interior.ColorIndex = 17

Sheets("Result").Range("I10").Value = "Success Transactions"
Sheets("Result").Range("I10").Font.Bold = True
Sheets("Result").Range("I10").Interior.ColorIndex = 46

Sheets("Result").Range("I11").Value = "Success Rate"
Sheets("Result").Range("I11").Font.Bold = True
Sheets("Result").Range("I11").Interior.ColorIndex = 45

Sheets("Result").Range("I12").Value = "Error Transactions"
Sheets("Result").Range("I12").Font.Bold = True
Sheets("Result").Range("I12").Interior.ColorIndex = 9

Sheets("Result").Range("I13").Value = "Error Rate"
Sheets("Result").Range("I13").Font.Bold = True
Sheets("Result").Range("I13").Interior.ColorIndex = 3

'add values to result chart

'Allen
Sheets("Result").Range("J9").Value = Allen_TT
Sheets("Result").Range("J10").Value = Allen_ST
Sheets("Result").Range("J11").Value = Allen_SR
Sheets("Result").Range("J11").NumberFormat = "0.00%"
Sheets("Result").Range("J12").Value = Allen_ET
Sheets("Result").Range("J13").Value = Allen_ER
Sheets("Result").Range("J13").NumberFormat = "0.00%"

'Bothell
Sheets("Result").Range("K9").Value = Bothell_TT
Sheets("Result").Range("K10").Value = Bothell_ST
Sheets("Result").Range("K11").Value = Bothell_SR
Sheets("Result").Range("K11").NumberFormat = "0.00%"
Sheets("Result").Range("K12").Value = Bothell_ET
Sheets("Result").Range("K13").Value = Bothell_ER
Sheets("Result").Range("K13").NumberFormat = "0.00%"

'totals
Sheets("Result").Range("L9").Value = Allen_TT + Bothell_TT
Sheets("Result").Range("L10").Value = Allen_ST + Bothell_ST
Sheets("Result").Range("L11").Value = Application.Max(Allen_SR, Bothell_SR)
Sheets("Result").Range("L11").NumberFormat = "0.00%"
Sheets("Result").Range("L12").Value = Allen_ET + Bothell_ET
Sheets("Result").Range("L13").Value = Application.Max(Allen_ER, Bothell_ER)
Sheets("Result").Range("L13").NumberFormat = "0.00%"

'delete unneeded worksheets
Application.DisplayAlerts = False
Worksheets("Bothell_Result").Delete
Worksheets("Allen_Result").Delete
Application.DisplayAlerts = True

End Sub

