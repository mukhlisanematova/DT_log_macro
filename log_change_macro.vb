git@github.com:mukhlisanematova/DT_log_macro.git

Sub CreateLogSheet()
    Dim wb As Workbook
    Dim wsLog As Worksheet
    Dim sheetName As String
    
    sheetName = "Log"   ' Name your log sheet however you like
    Set wb = ThisWorkbook
    
    ' Check if the sheet already exists
    On Error Resume Next
    Set wsLog = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    ' If not found, create it
    If wsLog Is Nothing Then
        Set wsLog = wb.Worksheets.Add( _
            After:=wb.Worksheets(wb.Worksheets.Count))
        wsLog.Name = sheetName
        
        ' Create headers in row 1, for instance:
        wsLog.Range("A1").Value = "Timestamp"
        wsLog.Range("B1").Value = "Cell"
        wsLog.Range("C1").Value = "Sheet"
        wsLog.Range("D1").Value = "Previous Value"
        wsLog.Range("E1").Value = "New Value"
        wsLog.Range("F1").Value = "Formula"
        wsLog.Range("G1").Value = "User"
        wsLog.Range("H1").Value = "Notes"
        
        ' Optional: You could format the headers (bold, column widths, etc.)
        ' For example:
        wsLog.Rows(1).Font.Bold = True
        wsLog.Columns("A:H").AutoFit
    End If
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo CleanExit
    Application.EnableEvents = False

    Dim j As Range
    Dim wb As Workbook
    Dim wsLog As Worksheet
    Dim NextRow As Long
    
    ' Variables to capture old and new values
    Dim newValue As Variant
    Dim oldValue As Variant
    Dim cellFormula As String
    Dim notes As String
    
    Call CreateLogSheet
    
    ' Set references
    Set wb = ThisWorkbook
    Set wsLog = wb.Worksheets("Log")
    
    ' Loop through each cell in the changed range
    For Each j In Target.Cells
        
        ' We first grab the new value from the changed cell
        newValue = j.Value
        
        If j.HasFormula Then
            cellFormula = j.Formula
        Else
            cellFormula = ""   ' or put "No formula" if you prefer
        End If
        
        Application.Undo
        oldValue = j.Value
        
        ' Undo again to put the new value back
        Application.Undo
        
        ' Move to next available row in Sheet2
        NextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
        
        ' Log the information
        Dim cellAdd As String
        cellAdd = j.Address
        
        wsLog.Cells(NextRow, 1).Value = Now()                 ' Timestamp
        wsLog.Cells(NextRow, 2).Value = j.Address             ' Cell
        wsLog.Cells(NextRow, 3).Value = j.Parent.Name         ' Sheet
        wsLog.Cells(NextRow, 4).Value = oldValue              ' Previous Value
        wsLog.Cells(NextRow, 5).Value = newValue              ' New Value
        wsLog.Cells(NextRow, 6).Value = cellFormula           ' Formula
        wsLog.Cells(NextRow, 7).Value = Environ("UserName")   ' User
        wsLog.Cells(NextRow, 8).Value = notes                 ' Notes
            
    Next j

CleanExit:
    Application.EnableEvents = True

End Sub

