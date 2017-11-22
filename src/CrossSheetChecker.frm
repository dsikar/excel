VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CrossSheetChecker 
   Caption         =   "Cross Sheet Checker"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16935
   OleObjectBlob   =   "CrossSheetChecker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CrossSheetChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Call openDialog
End Sub

Private Sub openDialog()
    Dim fd As Office.FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

   With fd

      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Please select the file."

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "Excel 2010", "*.xlsx"
      .Filters.Add "All Files", "*.*"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
        lblFile1.Caption = .SelectedItems(1) 'replace txtFileName with your textbox

      End If
   End With
End Sub
