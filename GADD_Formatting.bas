'Macro Formatting For Kpis
'By: Evan Robertson
'Last Edit August 19th 2017
'Notes:
'Need to Make paid to picked based off of Wait for merge not based of of picked column to better represent pick time of merged orders

'

Sub Formatting()
 '
 ' Formatting Macro
 '

     'Variable declaration
     Application.ScreenUpdating = False

     Dim ordercount As Integer
     Dim i As Long
     Dim CustomerCount As Integer
     Dim TransportCount As Integer
     Dim CollectCount As Integer
     Dim C2Count As Integer
     Dim TotalPicks As Integer
     Dim UniqueDeliveries As Integer
     
     'Formatting
     Sheets("Data").Select
     Range("B:B,E:E,G:O,Q:V,X:Z,AB:AZ").Select
     Selection.EntireColumn.Hidden = True
     ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
     ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add Key:=Range( _
         "C1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
         xlSortNormal
     ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add Key:=Range( _
         "D1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
         xlSortNormal
     With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
         .Header = xlYes
         .MatchCase = False
         .Orientation = xlTopToBottom
         .SortMethod = xlPinYin
         .Apply
     End With
     Range("BA1").FormulaR1C1 = "Paid to Picked"
     Range("BB1").FormulaR1C1 = "Picked to Checked"
     
     
     'Calculation
         For i = 2 To Rows.Count
            If Cells(i, 4).Value <> "" Then
                Cells(i, 53).FormulaR1C1 = "=RC[-30]-RC[-37]"
                If (Cells(i, 4).Value = "Customer 2 ") Or (Cells(i, 4).Value = "Customer 1 ") Or (Cells(i, 4).Value = "Transport 1") Then
                    Cells(i, 54).FormulaR1C1 = "=RC[-27]-RC[-31]"
                End If
        
    
                If Cells(i, 54).Value < 0 Then
                    Cells(i, 54).ClearContents
                End If
                If Cells(i, 53).Value < 0 Then
                    Cells(i, 53).ClearContents
                End If

                If Cells(i, 4).Value = "Customer 1 " Then
                    CustomerCount = CustomerCount + 1
        
                ElseIf Cells(i, 4).Value = "Transport 1" Then
                    TransportCount = TransportCount + 1
                    If Cells(i, 1).Value <> Cells(i - 1, 1) Then
                        UniqueDeliveries = UniqueDeliveries + 1
                    End If
                
                ElseIf Cells(i, 4).Value = "Customer 2 " Then
                    C2Count = C2Count + 1
                    
                ElseIf Cells(i, 4).Value = "Collect 1  " Or Cells(i, 4).Value = "Collect 2  " Then
                    CollectCount = CollectCount + 1
                End If
            End If
        Next i
    'Total Rows
    TotalPicks = CollectCount + C2Count + TransportCount + CustomerCount
   'Headings
    'FSHO Metrics
    Cells(2, 57).Value = "Total FS Picks"
    Cells(3, 57).Value = "Paid to picked"
    Cells(4, 57).Value = "Picked to Checked"
    Cells(5, 57).Value = "Total Wait Time"
    Cells(6, 57).Value = "Average Order Size"
    
    'C&C Metrics
    Cells(8, 57).Value = "Total Picks"
    Cells(9, 57).Value = "Paid to Picked"
    Cells(10, 57).Value = "Picked To Checked C2"
    Cells(11, 57).Value = "Average Order Size"
    
    'Home Delivery Metrics
    Cells(13, 57).Value = "Paid to Picked"
    Cells(14, 57).Value = "Picked to Checked"
    Cells(15, 57).Value = "Total Wait Time"
    Cells(16, 57).Value = "Average Order Size"
    Cells(17, 57).Value = "Number of Deliveries(Paid)"
    
'Values
    'FSHO Values
    Cells(2, 58).Value = CustomerCount
    Cells(3, 58).Value = Application.WorksheetFunction.Average(Range(Cells(CollectCount + C2Count + 2, 53), Cells(CollectCount + C2Count + 1 + CustomerCount, 53)))
    Cells(4, 58).Value = Application.WorksheetFunction.Average(Range(Cells(CollectCount + C2Count + 2, 54), Cells(CollectCount + C2Count + 1 + CustomerCount, 54)))
    Cells(5, 58).Value = Cells(4, 58).Value + Cells(3, 58).Value
    Cells(6, 58).Value = Application.WorksheetFunction.Average(Range(Cells(CollectCount + C2Count + 2, 6), Cells(CollectCount + C2Count + 1 + CustomerCount, 6)))
    
    'CNC Values
    Cells(8, 58).Value = CollectCount + C2Count
    Cells(9, 58).Value = Application.WorksheetFunction.Average(Range(Cells(2, 53), Cells(CollectCount + C2Count + 1, 53)))
    Cells(10, 58).Value = Application.WorksheetFunction.Average(Range(Cells(CollectCount + 2, 54), Cells(CollectCount + C2Count + 1, 54)))
    Cells(11, 58).Value = Application.WorksheetFunction.Average(Range(Cells(2, 6), Cells(CollectCount + C2Count + 1, 6)))
    
    'Delivery Values
    Cells(13, 58).Value = Application.WorksheetFunction.Average(Range(Cells(CollectCount + C2Count + CustomerCount + 2, 53), Cells(CollectCount + C2Count + 1 + CustomerCount + TransportCount, 53)))
    Cells(14, 58).Value = Application.WorksheetFunction.Average(Range(Cells(CollectCount + C2Count + CustomerCount + 2, 54), Cells(CollectCount + C2Count + 1 + CustomerCount + TransportCount, 54)))
    Cells(15, 58).Value = Cells(13, 58).Value + Cells(14, 58).Value
    Cells(16, 58).Value = Application.WorksheetFunction.Average(Range(Cells(CollectCount + C2Count + CustomerCount + 2, 6), Cells(CollectCount + C2Count + 1 + CustomerCount + TransportCount, 6)))
    Cells(17, 58).Value = UniqueDeliveries
    
    'Formatting
    Columns("BE:BE").EntireColumn.AutoFit
    Range("BF3:BF5,BF9,BF10,BF13:BF15").Select
    Selection.NumberFormat = "hh:mm:ss;@"
    Range("BF6,BF8,BF11,BF16:BF17").Select
    Selection.NumberFormat = "0"
    Range("BF2:BF6,BF8:BF11,BF13:BF17").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("BF17").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("BF12,BF7").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("BF2:BF17").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
        Range("BF2:BF17").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
 End Sub

