Sub Open_ExistingWorkbook()
    
    Dim wks1 As Worksheet
    Dim iCycle As Integer
    Dim strService As String
    Dim strPosition As String
    Dim lRITTC As Long
    
    iCycle = 266
    strService = "Voice"
    strPosition = "Payer EUR"
    lRITTC = 0
    
    
    ' Open Workbook (excel document)
    ' Workbooks.Open "C:\Users\dsikar\Documents\excel\clients2.xlsx"
    ' Activate Worksheet
    'Workbooks(1).Worksheets("Sheet1").Activate
    ' Set Worksheet document to active sheet
    Set wks1 = Workbooks(1).ActiveSheet
    ' Print a value to immediate (debug) screen
    Debug.Print wks1.Cells(15694, 2).Value
    Debug.Print wks1.Cells(15694, 16).Value
    Debug.Print wks1.Cells(15694, 9).Value
    Debug.Print wks1.Cells(15694, 4).Value
    ' Same without active sheet
    'Set wks2 = Workbooks(2).Worksheets("Sheet2")
    'Debug.Print wks2.Range("A1").Value
    Debug.Print Workbooks.Count
  '  For Each wks2cell In wks2rng
  '      Debug.Print wks2cell
  '  Next wks2cell
  
  ' Go through every row, if cycle = 266 and Service = voice and Position = Payer EUR, sum RI TTC
End Sub
