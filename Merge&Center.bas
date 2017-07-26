Attribute VB_Name = "Module1"
Sub Merge_Center()
Attribute Merge_Center.VB_ProcData.VB_Invoke_Func = " \n14"

'Macro to Merge and Center for Weeks
'

'Variable Declaration
    Dim Range As Integer
    Dim i As Long
    Dim First_Column As Integer
    Dim Last_Column As Integer

    
    


    First_Column = First_Column + 4
    Last_Column = Last_Column + 10
    
'Loop to process Merging
    Worksheets("Data").Activate
    
    For i = 1 To 52
    
        'Cells(Row, First_Column).Value = i
        ActiveSheet.Range(Cells(2, First_Column), Cells(2, Last_Column)).Select
        Selection.Merge
        First_Column = First_Column + 7
        Last_Column = Last_Column + 7
        
    Next i
    
    

End Sub
