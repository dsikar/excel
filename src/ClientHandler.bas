Attribute VB_Name = "Module1"
Sub Open_ExistingWorkbook()
    
    Dim wks2 As Worksheet
    ' Open Workbook (excel document)
    Workbooks.Open "C:\Users\dsikar\Documents\excel\clients2.xlsx"
    ' Activate Worksheet
    Workbooks(2).Worksheets("Sheet1").Activate
    ' Set Worksheet document to active sheet
    Set wks2 = Workbooks(2).ActiveSheet
    ' Print a value to immediate (debug) screen
    Debug.Print wks2.Range("A1").Value
    ' Same without active sheet
    Set wks2 = Workbooks(2).Worksheets("Sheet2")
    Debug.Print wks2.Range("A1").Value
    
    Dim wks3 As Worksheet
    ' Open Workbook (excel document)
    Workbooks.Open "C:\Users\dsikar\Documents\excel\clients3.xlsx"
    ' Activate Worksheet
    Workbooks(3).Worksheets("Sheet1").Activate
    ' Set Worksheet document to active sheet
    Set wks3 = Workbooks(3).ActiveSheet
    ' Print a value to immediate (debug) screen
    'Debug.Print wks3.Range("A1").Value
    ' Same without active sheet
    'Set wks3 = Workbooks(2).Worksheets("Sheet2")
    'Debug.Print wks2.Range("A1").Value
       
    
    ' Iterate
    Dim wks2rng As Range, wks2cell As Range
    Set wks2rng = wks2.Range("A1:A4")
    
    For Each wks2cell In wks2rng
        Debug.Print wks2cell
    Next wks2cell
    
        ' Iterate with counters
    Dim i As Integer
    For Counter = 1 To 4
        Set curcell = Workbooks(2).Worksheets("Sheet1").Cells(Counter, 2)
        If wks2.Cells(Counter, 2) < wks3.Cells(Counter, 2) Then
            ' do things
        End If
        ' If Abs(curCell.Value) < 0.01 Then curCell.Value = 0
        ' Debug.Print curcell
        ' i = LookUp(1, wks3)
        'wks2.Rows(Counter).Copy _
        'wks2.Rows(Counter + 4)
    Next Counter
    
    ' Copy rows 1 to 4 from wks2 sheet to wks3 sheet beginning at row 12
    wks2.Range("A1:B4").Copy wks3.Range("D6:E9")
    
    ' Workbooks.Open "C:\Users\dsikar\Documents\excel\clients3.xlsx"
    Debug.Print Workbooks.Count
    ' Index starts at 1 - Need to close in reverse order
    Workbooks(3).Save
    Workbooks(3).Close
    Workbooks(2).Close
End Sub

Function LookUp(CustName As String, wks3 As Worksheet) As Integer

    ' TODO add open and close functions to decrease overhead
    ' Open worksheet
    For Counter = 1 To 4
        Set curcell = wks3.Cells(Counter, 2)
        ' If Abs(curCell.Value) < 0.01 Then curCell.Value = 0
        If curcell > CustName Then
            Debug.Print curcell
        End If
    Next Counter
    
    LookUp = 1

End Function



