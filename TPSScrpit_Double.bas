

Sub MakeTPSPivot_Double()

Dim pt As PivotTable
Dim cacheOfpt As PivotCache  'source data for pt
Dim pf As PivotField
Dim pi As PivotItem
Dim WSD As Worksheet
Dim PRange As Range
Dim wsSheet1 As Worksheet

'try to rename current workbook 'Data' if it is not already
On Error Resume Next
Set wsSheet1 = Sheets("Allen")
On Error GoTo 0
If Not wsSheet1 Is Nothing Then
    Sheets("Allen").Select ' does exist
Else
    ActiveSheet.Name = "Allen" 'does not exist
End If

'get data rage size for later
Set WSD = Worksheets("Allen")
FinalRow = WSD.Cells(Application.Rows.Count, 1).End(xlUp).Row
FinalCol = WSD.Cells(1, Application.Columns.Count). _
    End(xlToLeft).Column
Set PRange = WSD.Cells(1, 1).Resize(FinalRow, FinalCol)

'rename the EDR Types
Dim c As Range
Application.ScreenUpdating = False
For Each c In Range("H1", Range("H" & Rows.Count).End(xlUp))
  If Left(c, 35) Like " ServiceComponentNnsSMPPServiceSMS_" Then c = Mid(c, 35, Len(c))
  
Next c
Application.ScreenUpdating = True

Dim d As Range
Application.ScreenUpdating = False
For Each d In Range("H1", Range("H" & Rows.Count).End(xlUp))
  If Left(d, 1) = "_" Then d = Mid(d, 2, Len(d) - 1)
Next d
Application.ScreenUpdating = True

'try to create workbook named 'Result' if it does not already exist
Dim wsSheet2 As Worksheet
On Error Resume Next
Set wsSheet2 = Sheets("Result")
On Error GoTo 0
If Not wsSheet2 Is Nothing Then
    Sheets("Result").Select ' does exist
Else
    Sheets.Add.Name = "Result" 'does not exist
End If

'try to create workbook named 'Result' if it does not already exist
Dim wsSheet As Worksheet
On Error Resume Next
Set wsSheet = Sheets("Allen_Result")
On Error GoTo 0
If Not wsSheet Is Nothing Then
    Sheets("Allen_Result").Select ' does exist
Else
    Sheets.Add.Name = "Allen_Result" 'does not exist
End If

'try to delete previous pivot table if one exists
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
    .PivotFields(" Avg TPS").Orientation = xlDataField
    .PivotFields(" Max TPS").Orientation = xlDataField
    .PivotFields(" Min TPS").Orientation = xlDataField
    .PivotFields(" EDR Type").Orientation = xlColumnField
End With

'display only nodes that are needed
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
Sheets("Allen_Result").Range("A4").Value = "Intervals"
Sheets("Allen_Result").Range("B4").Value = "Average"
Sheets("Allen_Result").Range("C4").Value = "Max TPS"
Sheets("Allen_Result").Range("D4").Value = "Min TPS"

'set max, average and min settings
With ActiveSheet.PivotTables("MyPT").PivotFields("Average")
    .Caption = "Average TPS"
    .Function = xlAverage
End With
With ActiveSheet.PivotTables("MyPT").PivotFields("Max TPS")
    .Caption = "Max TPS"
    .Function = xlMax
End With
With ActiveSheet.PivotTables("MyPT").PivotFields("Min TPS")
    .Caption = "Min TPS"
    .Function = xlMin
End With

'format number columns
Columns(2).NumberFormat = "0.00"
Columns(3).NumberFormat = "0.00"

'find the last cell in column to get toal and rate values for result table
Allen_Bind_Avg = Range("B65536").End(xlUp).Value
Allen_Bind_Max = Range("E65536").End(xlUp).Value
Allen_Bind_Min = Range("H65536").End(xlUp).Value

Allen_Submit_Avg = Range("D65536").End(xlUp).Value
Allen_Submit_Max = Range("G65536").End(xlUp).Value
Allen_Submit_Min = Range("J65536").End(xlUp).Value

Allen_Deliver_Avg = Range("C65536").End(xlUp).Value
Allen_Deliver_Max = Range("F65536").End(xlUp).Value
Allen_Deliver_Min = Range("I65536").End(xlUp).Value

'get range data of pivot table
Set WSD = Worksheets("Allen_Result")
FinalRow = WSD.Cells(Application.Rows.Count, 1).End(xlUp).Row
FinalCol = WSD.Cells(1, Application.Columns.Count). _
    End(xlToLeft).Column
Set PRange = WSD.Cells(1, 1).Resize(FinalRow, FinalCol)

'add text
Sheets("Result").Range("A2").Value = "Allen Datacenter"
Sheets("Result").Range("A2").Font.Bold = True
Sheets("Result").Range("A2:G2").Interior.ColorIndex = 15
Sheets("Result").Range("B27").Value = "Centeral Time"
Sheets("Result").Range("B27").Font.Bold = True
Sheets("Result").Range("B27:C27").Interior.ColorIndex = 15

Sheets("Result").Select
'create line chart
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=PRange
ActiveChart.ChartType = xlLineMarkers
With ActiveChart.Parent
         .Height = 350 ' resize
         .Width = 500  ' resize
         .Top = 30    ' reposition
         .Left = 0   ' reposition
End With


'---------------------------------------------
'create bothell data

'try to rename current workbook 'Data' if it is not already
On Error Resume Next
Set wsSheet1 = Sheets("Bothell")
On Error GoTo 0
If Not wsSheet1 Is Nothing Then
    Sheets("Bothell").Select ' does exist
Else
    ActiveSheet.Name = "Bothell" 'does not exist
End If

'get data rage size for later
Set WSD = Worksheets("Bothell")
FinalRow = WSD.Cells(Application.Rows.Count, 1).End(xlUp).Row
FinalCol = WSD.Cells(1, Application.Columns.Count). _
    End(xlToLeft).Column
Set PRange = WSD.Cells(1, 1).Resize(FinalRow, FinalCol)

'rename the EDR Types
Dim a As Range
Application.ScreenUpdating = False
For Each a In Range("H1", Range("H" & Rows.Count).End(xlUp))
  If Left(a, 35) Like " ServiceComponentNnsSMPPServiceSMS_" Then a = Mid(a, 35, Len(a))
  
Next a
Application.ScreenUpdating = True

Dim q As Range
Application.ScreenUpdating = False
For Each q In Range("H1", Range("H" & Rows.Count).End(xlUp))
  If Left(q, 1) = "_" Then q = Mid(q, 2, Len(q) - 1)
Next q
Application.ScreenUpdating = True

'try to create workbook named 'Result' if it does not already exist
Dim wsSheet4 As Worksheet
On Error Resume Next
Set wsSheet4 = Sheets("Result")
On Error GoTo 0
If Not wsSheet4 Is Nothing Then
    Sheets("Result").Select ' does exist
Else
    Sheets.Add.Name = "Result" 'does not exist
End If

'try to create workbook named 'Result' if it does not already exist
Dim wsSheet5 As Worksheet
On Error Resume Next
Set wsSheet5 = Sheets("Bothell_Result")
On Error GoTo 0
If Not wsSheet5 Is Nothing Then
    Sheets("Bothell_Result").Select ' does exist
Else
    Sheets.Add.Name = "Bothell_Result" 'does not exist
End If

'try to delete previous pivot table if one exists
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
    .PivotFields(" Avg TPS").Orientation = xlDataField
    .PivotFields(" Max TPS").Orientation = xlDataField
    .PivotFields(" Min TPS").Orientation = xlDataField
    .PivotFields(" EDR Type").Orientation = xlColumnField
End With

'display only nodes that are needed
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
Sheets("Bothell_Result").Range("A4").Value = "Intervals"
Sheets("Bothell_Result").Range("B4").Value = "Average"
Sheets("Bothell_Result").Range("C4").Value = "Max TPS"
Sheets("Bothell_Result").Range("D4").Value = "Min TPS"

'set max, average and min settings
With ActiveSheet.PivotTables("MyPT").PivotFields("Average")
    .Caption = "Average TPS"
    .Function = xlAverage
End With
With ActiveSheet.PivotTables("MyPT").PivotFields("Max TPS")
    .Caption = "Max TPS"
    .Function = xlMax
End With
With ActiveSheet.PivotTables("MyPT").PivotFields("Min TPS")
    .Caption = "Min TPS"
    .Function = xlMin
End With

'format number columns
Columns(2).NumberFormat = "0.00"
Columns(3).NumberFormat = "0.00"

'find the last cell in column to get toal and rate values for result table
Bothell_Bind_Avg = Range("B65536").End(xlUp).Value
Bothell_Bind_Max = Range("E65536").End(xlUp).Value
Bothell_Bind_Min = Range("H65536").End(xlUp).Value

Bothell_Submit_Avg = Range("D65536").End(xlUp).Value
Bothell_Submit_Max = Range("G65536").End(xlUp).Value
Bothell_Submit_Min = Range("J65536").End(xlUp).Value

Bothell_Deliver_Avg = Range("C65536").End(xlUp).Value
Bothell_Deliver_Max = Range("F65536").End(xlUp).Value
Bothell_Deliver_Min = Range("I65536").End(xlUp).Value

'get range data of pivot table
Set WSD = Worksheets("Bothell_Result")
FinalRow = WSD.Cells(Application.Rows.Count, 1).End(xlUp).Row
FinalCol = WSD.Cells(1, Application.Columns.Count). _
    End(xlToLeft).Column
Set PRange = WSD.Cells(1, 1).Resize(FinalRow, FinalCol)

'add text
Sheets("Result").Range("A29").Value = "Bothell Datacenter"
Sheets("Result").Range("A29").Font.Bold = True
Sheets("Result").Range("A29:G29").Interior.ColorIndex = 15
Sheets("Result").Range("B54").Value = "Pacific Time"
Sheets("Result").Range("B54").Font.Bold = True
Sheets("Result").Range("B54:C54").Interior.ColorIndex = 15

Sheets("Result").Select
'create line chart
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=PRange
ActiveChart.ChartType = xlLineMarkers
With ActiveChart.Parent
         .Height = 350 ' resize
         .Width = 500  ' resize
         .Top = 435    ' reposition
         .Left = 0   ' reposition
End With

'-------------------------------------------------------------------------
'draw result table
Sheets("Result").Range("M18").Value = "Min"
Sheets("Result").Range("M18").Font.Bold = True
Sheets("Result").Range("M18:M21").Interior.ColorIndex = 15

Sheets("Result").Range("N18").Value = "Max"
Sheets("Result").Range("N18").Font.Bold = True
Sheets("Result").Range("N18:N21").Interior.ColorIndex = 37

Sheets("Result").Range("O18").Value = "Avg"
Sheets("Result").Range("O18").Font.Bold = True
Sheets("Result").Range("O18:O21").Interior.ColorIndex = 16

Sheets("Result").Range("L19").Value = "Bind"
Sheets("Result").Range("L19").Font.Bold = True
Sheets("Result").Range("L19").Interior.ColorIndex = 33

Sheets("Result").Range("L20").Value = "Submit"
Sheets("Result").Range("L20").Font.Bold = True
Sheets("Result").Range("L20").Interior.ColorIndex = 33

Sheets("Result").Range("L21").Value = "Deliver"
Sheets("Result").Range("L21").Font.Bold = True
Sheets("Result").Range("L21").Interior.ColorIndex = 33

Sheets("Result").Range("M19").Value = Allen_Bind_Min
Sheets("Result").Range("N19").Value = Allen_Bind_Max
Sheets("Result").Range("O19").Value = Allen_Bind_Avg
Sheets("Result").Range("O19").NumberFormat = "0.00%"

Sheets("Result").Range("M20").Value = Allen_Submit_Min
Sheets("Result").Range("N20").Value = Allen_Submit_Max
Sheets("Result").Range("O20").Value = Allen_Submit_Avg
Sheets("Result").Range("O20").NumberFormat = "0.00%"

Sheets("Result").Range("M21").Value = Allen_Deliver_Min
Sheets("Result").Range("N21").Value = Allen_Deliver_Max
Sheets("Result").Range("O21").Value = Allen_Deliver_Avg
Sheets("Result").Range("O21").NumberFormat = "0.00%"
'---

Sheets("Result").Range("M45").Value = "Min"
Sheets("Result").Range("M45").Font.Bold = True
Sheets("Result").Range("M45:M48").Interior.ColorIndex = 15

Sheets("Result").Range("N45").Value = "Max"
Sheets("Result").Range("N45").Font.Bold = True
Sheets("Result").Range("N45:N48").Interior.ColorIndex = 37

Sheets("Result").Range("O45").Value = "Avg"
Sheets("Result").Range("O45").Font.Bold = True
Sheets("Result").Range("O45:O48").Interior.ColorIndex = 16

Sheets("Result").Range("L46").Value = "Bind"
Sheets("Result").Range("L46").Font.Bold = True
Sheets("Result").Range("L46").Interior.ColorIndex = 33

Sheets("Result").Range("L47").Value = "Submit"
Sheets("Result").Range("L47").Font.Bold = True
Sheets("Result").Range("L47").Interior.ColorIndex = 33

Sheets("Result").Range("L48").Value = "Deliver"
Sheets("Result").Range("L48").Font.Bold = True
Sheets("Result").Range("L48").Interior.ColorIndex = 33

Sheets("Result").Range("M46").Value = Bothell_Bind_Min
Sheets("Result").Range("N46").Value = Bothell_Bind_Max
Sheets("Result").Range("O46").Value = Bothell_Bind_Avg
Sheets("Result").Range("O46").NumberFormat = "0.00%"

Sheets("Result").Range("M47").Value = Bothell_Submit_Min
Sheets("Result").Range("N47").Value = Bothell_Submit_Max
Sheets("Result").Range("O47").Value = Bothell_Submit_Avg
Sheets("Result").Range("O47").NumberFormat = "0.00%"

Sheets("Result").Range("M48").Value = Bothell_Deliver_Min
Sheets("Result").Range("N48").Value = Bothell_Deliver_Max
Sheets("Result").Range("O48").Value = Bothell_Deliver_Avg
Sheets("Result").Range("O48").NumberFormat = "0.00%"

'delete unneeded worksheets
Application.DisplayAlerts = False
Worksheets("Bothell_Result").Delete
Worksheets("Allen_Result").Delete
Application.DisplayAlerts = True


End Sub

