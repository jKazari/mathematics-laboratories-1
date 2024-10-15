Sub CreateTableAndFill()
	Dim numRows As Integer
	Dim numCols As Integer
	Dim currentRow As Integer
	Dim currentCol As Integer
	Dim userInput As Variant

	numRows = InputBox("Enter the number of rows:", "Table Dimensions")
	numCols = InputBox("Enter the number of columns:", "Table Dimensions")

    If IsNumeric(numRows) And IsNumeric(numCols) And numRows > 0 And numCols > 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Table"
        Set ws = ActiveSheet
      
        For currentRow = 1 To numRows
            For currentCol = 1 To numCols
                userInput = InputBox("Enter value for cell [" & currentRow & ", " & currentCol & "]:", "Enter Value")
              
                If userInput <> "" Then
                    ws.Cells(currentRow, currentCol).Value = userInput
                Else
                    Exit Sub
                End If
            Next currentCol
        Next currentRow
    Else
        MsgBox "Please enter valid numeric values for rows and columns.", vbExclamation
    End If
End Sub