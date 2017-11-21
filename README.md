# excel

Excel job to implement vlookup style functionality with if statements, to compare two or more sheets in the same or different workbooks and add rows as applicable to a third workbook.

NB When running macros, path of least resistance (bypass saving worksheet with  
macro) is to import and export macros as required. For this end, open a workbook  
then key ALT + F11 to open VBA console, then File > Import and navigate to macro  
be it a module (.bas), class (.cls) or form (.frm).

## Work units

Once a workbook is open, and macro loaded:

* Open workbooks with handles

Open a work
```
' Declare handle
Dim wks2 As Worksheet
' Open Workbook (excel document), note with current scheme this will be the
' second workbook open hence (2) reference found.
Workbooks.Open "path to .xlsx"
' Set Worksheet document to active sheet
Set wks2 = Workbooks(2).ActiveSheet
' Print a value to immediate (debug) screen
Debug.Print wks2.Range("A1").Value
' Same without active sheet
Set wks2 = Workbooks(2).Worksheets("Sheet2")
Debug.Print wks2.Range("A1").Value

```
* Iterate through rows
Note we are using handles
```
Dim wks2rng As Range, wks2cell As Range
Set wks2rng = wks2.Range("A1:A4")
For Each wks2cell In wks2rng
    Debug.Print wks2cell
Next wks2cell

' Iterate with a for loop
Dim i As Integer
For Counter = 1 To 4
    Set curcell = Workbooks(2).Worksheets("Sheet1").Cells(Counter, 2)
    ' Alternatively
    Set curcell = wks2.Cells(Counter, 2)
    ' etc
Next Counter
```
* Compare column values
```
' Comparing rules
If wks2.Cells(Counter, 2) < LookUp(Counter, wks3) Then
    ' do things
End If

' using a function

Function LookUp(CustName As String, wks As Worksheet) As Integer

    For Counter = 1 To 4
        ' do things
    Next Counter

    LookUp = 1
End Function
```
* Write rows to target workbooks/sheet
```
' Copy rows 1 to 4 from wks2 sheet to wks3 sheet beginning at row 12
wks2.Rows("1:4").Copy wks3.Rows("12")
' Copy cell ranges
wks2.Range("A1:B4").Copy wks3.Range("D6:E9")
```

* Close all workbooks
```
' Save and close in reverse order to avoid index out of bounds error
Workbooks(3).Save
Workbooks(3).Close
Workbooks(2).Close
```            
## Implementation

To be implemented using modules and/or classes and/or forms.
