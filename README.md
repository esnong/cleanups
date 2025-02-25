Sub FolderFileDataCleaning()
    Dim FolderPath As String
    Dim FileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim DebitRow As Long
    Dim YearValue As Integer, MonthValue As Integer
    Dim DateValue As String
    Dim FolderPicker As FileDialog
    Dim rng As Range
    Dim fso As Object
    Dim FoundValidSheet As Boolean
    
    ' Ask user to select a folder
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With FolderPicker
        .Title = "Select Folder Containing Excel Files"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub ' Exit if user cancels
        FolderPath = .SelectedItems(1) & "\"
    End With
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Process all Excel files in the selected folder
    FileName = Dir(FolderPath & "*.xlsx") ' only .xlsx format
    Do While FileName <> ""
        ' Open the workbook safely
        On Error Resume Next
        Set wb = Workbooks.Open(FolderPath & FileName)
        On Error GoTo 0
        
        ' Check if workbook opened successfully
        If wb Is Nothing Then
            MsgBox "Failed to open " & FileName, vbExclamation
            GoTo NextFile
        End If
        
        FoundValidSheet = False
        Set ws = Nothing ' Reset ws for each file
        
        ' Loop through all sheets to find a valid one
        For Each ws In wb.Sheets
            On Error Resume Next ' Prevents errors if cells are empty
            YearValue = ws.Range("E13").Value
            MonthValue = ws.Range("E15").Value
            On Error GoTo 0
            
            ' Ensure valid month and year
            If YearValue >= 1900 And YearValue <= 2100 And MonthValue >= 1 And MonthValue <= 12 Then
                FoundValidSheet = True
                Exit For
            End If
        Next ws

        ' If no valid sheet is found, skip the file
        If Not FoundValidSheet Or ws Is Nothing Then
            MsgBox "No valid data found in " & FileName & ". Skipping file.", vbExclamation
            wb.Close False
            FileName = Dir
            GoTo NextFile
        End If
        
        ' Format the date as MM/1/YYYY
        DateValue = MonthValue & "/1/" & YearValue
        
        ' Remove unnecessary columns (Keep only B and E)
        Application.DisplayAlerts = False ' Suppress delete confirmation
        ws.Columns("A").Delete
        ws.Columns("B:C").Delete
        ws.Columns("C:Z").Delete
        Application.DisplayAlerts = True
        
        ' Remove rows 1 to 34
        ws.Rows("1:34").Delete Shift:=xlUp
        
        ' Add headers in the first row
        ws.Cells(1, 1).Value = "Cost Elements"
        ws.Cells(1, 2).Value = "Act. Costs"
        ws.Cells(1, 3).Value = "Date"

        ' Fill down the date column in column C
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ws.Range("C2:C" & LastRow).Value = DateValue
        
        ' Copy contents from column E to column B
        
        ' Find the row containing "* Debit" in Column A
        Set rng = ws.Range("A:A").Find("* Debit", LookAt:=xlPart)
        If Not rng Is Nothing Then
            DebitRow = rng.Row
            LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            ws.Rows(DebitRow & ":" & LastRow).Delete Shift:=xlUp
        End If
        
        ' Save and close workbook
        wb.Save
        wb.Close False
        
NextFile:
        ' Get next file
        FileName = Dir
    Loop
    
    ' Cleanup
    MsgBox "Processing complete!", vbInformation
End Sub


