
'Calculate countIndex, searchValue, columnIndex
Sub weekly_aggregates()
    'Declarations
    Application.ScreenUpdating = False
    'preparation
    Dim columnIndex As Integer
    Dim i As Long
    Dim countIndex As Integer
    Dim searchValue As Integer
    Dim counter As Integer
        'Algorithm
    Sheets("Data").Select
    
    'Loop for Every week
    For searchValue = 1 To 52
            'Finds where row begins
        For i = 1 To Columns.Count
            If Cells(2, i).Value = searchValue Then
                columnIndex = i
            End If
        Next i
        'Splits based on week number as week 1 only contains 2 days
        If searchValue = 1 Then
            For counter = 0 To 1
                If counter = 0 Then
                    For countIndex = 6 To 40
                        If Not ((countIndex = 12) Or (countIndex = 19) Or (countIndex = 23) Or (countIndex = 29) Or (countIndex = 34)) Then
                            Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value = Cells(countIndex, columnIndex + counter).Value
                        End If
                        If Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value = 0 Then
                            Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).ClearContents
                        End If
                    Next countIndex
                ElseIf counter < 7 Then
                    For countIndex = 6 To 40
                        If Not ((countIndex = 12) Or (countIndex = 19) Or (countIndex = 23) Or (countIndex = 29) Or (countIndex = 34)) Then
                            Calculate countIndex, searchValue, columnIndex, counter
                        End If
                    'Averages countIndex, searchValue
                    If counter = 1 Then
                        Averages countIndex, searchValue, 2
                    End If
                    If Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value = 0 Then
                        Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).ClearContents
                    End If
                    Next countIndex
                Else
                End If
            Next counter
        'For any value besides one
        Else
            For counter = 0 To 6
            If counter = 0 Then
                For countIndex = 6 To 40
                    If Not ((countIndex = 12) Or (countIndex = 19) Or (countIndex = 23) Or (countIndex = 29) Or (countIndex = 34)) Then
                        Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value = Cells(countIndex, columnIndex + counter).Value
                    End If
                    If Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value = 0 Then
                        Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).ClearContents
                    End If
                Next countIndex
            ElseIf counter < 7 Then
                For countIndex = 6 To 40
                    If Not ((countIndex = 12) Or (countIndex = 19) Or (countIndex = 23) Or (countIndex = 29) Or (countIndex = 34)) Then
                        Calculate countIndex, searchValue, columnIndex, counter
                    End If
                'Averages countIndex, searchValue
                If counter = 6 Then
                    Averages countIndex, searchValue, 7
                End If
                If Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value = 0 Then
                    Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).ClearContents
                End If
                Next countIndex
            Else
            End If
        Next counter
        End If
    Next searchValue
    Sheets("Weekly Aggregates").Select
End Sub
'The following functions are based off cell references in data sheet

Sub Calculate(countIndex As Integer, searchValue As Integer, columnIndex As Integer, counter As Integer)
    'Used when counter > 0 so when it is not the first value of the week
    Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value = Cells(countIndex, columnIndex + counter).Value + Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value
End Sub

Sub Averages(countIndex, searchValue, days As Integer)
    'Average Out waittimes, and percentages
    If ((countIndex > 24 And countIndex < 29) Or (countIndex > 30 And countIndex < 34) Or (countIndex > 34 And countIndex < 39)) Then
        Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value = Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value / days
    End If
End Sub
