Sub Formatting()
 '
 ' Formatting Macro
 '

     'Variable declaration
     Dim ordercount As Integer
     Dim i As Long
     Dim CustomerCount As Integer
     Dim TransportCount As Integer

     'Formatting
     Sheets("Data").Select
     Columns("B:B").Select
     Selection.EntireColumn.Hidden = True
     Range("D:D,E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L").Select
     Range("L1").Activate
     Selection.EntireColumn.Hidden = True
     Range("N:N,R:R").Select
     Range("R1").Activate
     Selection.EntireColumn.Hidden = True
     Columns("T:T").Select
     Selection.EntireColumn.Hidden = True
     ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
     ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add Key:=Range( _
         "C1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
         xlSortNormal
     With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
         .Header = xlYes
         .MatchCase = False
         .Orientation = xlTopToBottom
         .SortMethod = xlPinYin
         .Apply
     End With
     Range("X1").FormulaR1C1 = "Paid to Picked"
     Range("Y1").FormulaR1C1 = "Picked to Checked"
     
     
     'Calculation
     
    For i = 1 To Rows.Count
    
        If Cells(i, 4).Value = "Customer 1 " Or Cells(i, 4).Value = "Transport 1" Then
            Cells(i, 25).FormulaR1C1 = "=RC[+2]-RC[-2]"
            Cells(i, 24).FormulaR1C1 = "=RC[-1]-RC[-8]"
        
    
            If Cells(i, 25).Value < 0 Then
                Cells(i, 25).ClearContents
            End If
            If Cells(i, 24).Value < 0 Then
                Cells(i, 24).ClearContents
            End If
        End If

        If Cells(i, 4).Value = "Customer 1 " Then
            CustomerCount = CustomerCount + 1
        
        ElseIf Cells(i, 4).Value = "Transport 1" Then
            TransportCount = TransportCount + 1
        End If
    

    Next i

    'Headings
    
    Cells(4 + TransportCount + CustomerCount, 25).Value = "Total Time Average"
    Cells(5 + TransportCount + CustomerCount, 25).Value = "Handout"
    Cells(6 + TransportCount + CustomerCount, 25).Value = "Home Delivery"
    Cells(7 + TransportCount + CustomerCount, 25).Value = "Paid to Picked"
    Cells(8 + TransportCount + CustomerCount, 25).Value = "Handout"
    Cells(9 + TransportCount + CustomerCount, 25).Value = "Home Delivery"
    Cells(10 + TransportCount + CustomerCount, 25).Value = "Picked to Checked"
    Cells(11 + TransportCount + CustomerCount, 25).Value = "Handout"
    Cells(12 + TransportCount + CustomerCount, 25).Value = "Home Delivery"


    
    'Values
    
    Cells(7 + TransportCount + CustomerCount, 26).Value = Application.WorksheetFunction.Average(Range(Cells(2, 24), Cells(TransportCount + CustomerCount + 1, 24)))
    Cells(8 + TransportCount + CustomerCount, 26).Value = Application.WorksheetFunction.Average(Range(Cells(2, 24), Cells(CustomerCount + 1, 24)))
    Cells(9 + TransportCount + CustomerCount, 26).Value = Application.WorksheetFunction.Average(Range(Cells(CustomerCount + 2, 24), Cells(CustomerCount + TransportCount + 1, 24)))
    Cells(10 + TransportCount + CustomerCount, 26).Value = Application.WorksheetFunction.Average(Range(Cells(2, 25), Cells(TransportCount + CustomerCount + 1, 25)))
    Cells(11 + TransportCount + CustomerCount, 26).Value = Application.WorksheetFunction.Average(Range(Cells(2, 25), Cells(CustomerCount + 1, 25)))
    Cells(12 + TransportCount + CustomerCount, 26).Value = Application.WorksheetFunction.Average(Range(Cells(CustomerCount + 2, 25), Cells(CustomerCount + TransportCount + 1, 25)))
    Cells(4 + TransportCount + CustomerCount, 26).Value = Application.WorksheetFunction.Sum(Cells(7 + TransportCount + CustomerCount, 26), Cells(10 + TransportCount + CustomerCount, 26))
    Cells(5 + TransportCount + CustomerCount, 26).Value = Application.WorksheetFunction.Sum(Cells(8 + TransportCount + CustomerCount, 26), Cells(11 + TransportCount + CustomerCount, 26))
    Cells(6 + TransportCount + CustomerCount, 26).Value = Application.WorksheetFunction.Sum(Cells(9 + TransportCount + CustomerCount, 26), Cells(12 + TransportCount + CustomerCount, 26))
    
    Range(Cells(4 + TransportCount + CustomerCount, 25), Cells(12 + TransportCount + CustomerCount, 26)).Select
    Selection.NumberFormat = "hh:mm:ss;@"
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
