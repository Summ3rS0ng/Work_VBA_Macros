Attribute VB_Name = "Module4"

'Calculate countIndex, searchValue, columnIndex
Sub weekly_aggregates()
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
        'Splits based on week number as week 1 only contains 3 days
        If searchValue = 1 Then
            For counter = 0 To 2
            If counter = 0 Then
                For countIndex = 6 To 40
                    If Not ((countIndex = 12) Or (countIndex = 24) Or (countIndex = 28) Or (countIndex = 39) Or ((countIndex > 14) And (countIndex < 20))) Then
                        Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value = Cells(countIndex, columnIndex + counter).Value
                    End If
                Next countIndex
            ElseIf counter < 7 Then
                For countIndex = 6 To 40
                    If Not ((countIndex = 12) Or (countIndex = 24) Or (countIndex = 28) Or (countIndex = 39) Or ((countIndex > 14) And (countIndex < 20))) Then
                        Calculate countIndex, searchValue, columnIndex, counter
                    End If
                'Averages countIndex, searchValue
                If counter = 2 Then
                    Averages countIndex, searchValue, 3
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
                    If Not ((countIndex = 12) Or (countIndex = 24) Or (countIndex = 28) Or (countIndex = 39) Or ((countIndex > 14) And (countIndex < 20))) Then
                        Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value = Cells(countIndex, columnIndex + counter).Value
                    End If
                Next countIndex
            ElseIf counter < 7 Then
                For countIndex = 6 To 40
                    If Not ((countIndex = 12) Or (countIndex = 24) Or (countIndex = 28) Or (countIndex = 39) Or ((countIndex > 14) And (countIndex < 20))) Then
                        Calculate countIndex, searchValue, columnIndex, counter
                    End If
                'Averages countIndex, searchValue
                If counter = 6 Then
                    Averages countIndex, searchValue, 7
                End If
                Next countIndex
            Else
            End If
        Next counter
        End If

        
    Next searchValue
    Sheets("Weekly Aggregates").Select
End Sub


Sub Calculate(countIndex As Integer, searchValue As Integer, columnIndex As Integer, counter As Integer)
    'Used when counter > 0 so when it is not the first value of the week
    Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value = Cells(countIndex, columnIndex + counter).Value + Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value
End Sub

Sub Averages(countIndex, searchValue, days As Integer)
    'Average Out waittimes, and percentages
    If ((countIndex = 11) Or ((countIndex > 29) And (countIndex < 39))) Then
        Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value = Sheets("Weekly Aggregates").Cells(countIndex - 2, searchValue + 1).Value / days
    End If
End Sub
