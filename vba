Sub FilterAndSplitDataWithWorkbookTitle()
    Dim wsSource As Worksheet
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim lastRow As Long, lastRowK As Long
    Dim uniqueValues As Collection
    Dim cell As Range
    Dim filterValue As String
    Dim sheetName As String
    Dim totalSum As Double
    Dim workbookTitle As String

    ' Set the source worksheet
    Set wsSource = ThisWorkbook.Sheets("Data") ' Data pasted on this sheet

    ' Get the value of B2 from Sheet1 for the new workbook title
    workbookTitle = ThisWorkbook.Sheets("Sheet1").Range("B2").Value & " - Broken out by transaction type"

    ' Determine the last row in Column B
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row

    ' Create a new workbook
    Set wbNew = Application.Workbooks.Add

    ' Copy unfiltered data to the first sheet in the new workbook
    wsSource.UsedRange.Copy Destination:=wbNew.Sheets(1).Range("A1")
    wbNew.Sheets(1).Name = "Unfiltered Data"

    ' Create a collection to store unique values in Column B
    Set uniqueValues = New Collection
    On Error Resume Next
    For Each cell In wsSource.Range("B2:B" & lastRow) ' Assuming data starts from Row 2
        uniqueValues.Add cell.Value, CStr(cell.Value)
    Next cell
    On Error GoTo 0

    ' Loop through unique values and create new sheets in the new workbook
    Application.ScreenUpdating = False
    For Each filterValue In uniqueValues
        ' Add a new worksheet to the new workbook
        Set wsNew = wbNew.Sheets.Add

        ' Check if the value starts with "Composer" and truncate to 16 characters if so, otherwise use the full value
        If Left(filterValue, 7) = "Composer" Then
            sheetName = Left(filterValue, 16) ' Use only the first 16 characters if it starts with "Composer"
        Else
            sheetName = filterValue ' Use the full name
        End If

        ' Sanitize sheet name to remove invalid characters
        sheetName = Application.Substitute(sheetName, "/", "_")
        sheetName = Application.Substitute(sheetName, "\", "_")
        sheetName = Application.Substitute(sheetName, ":", "_")
        sheetName = Application.Substitute(sheetName, "?", "_")
        sheetName = Application.Substitute(sheetName, "*", "_")
        sheetName = Application.Substitute(sheetName, "[", "_")
        sheetName = Application.Substitute(sheetName, "]", "_")

        ' Assign the sheet name
        wsNew.Name = sheetName

        ' Copy the header row to the new sheet
        wsSource.Rows(1).Copy Destination:=wsNew.Rows(1)

        ' Apply filter and copy filtered data
        wsSource.Rows(1).AutoFilter Field:=2, Criteria1:=filterValue
        wsSource.Rows(2 & ":" & lastRow).SpecialCells(xlCellTypeVisible).Copy Destination:=wsNew.Rows(2)
        wsSource.AutoFilterMode = False

        ' Add Total for Column K
        lastRowK = wsNew.Cells(wsNew.Rows.Count, "K").End(xlUp).Row
        If lastRowK >= 2 Then ' Ensure there's data in Column K
            totalSum = Application.WorksheetFunction.Sum(wsNew.Range("K2:K" & lastRowK))
            wsNew.Cells(lastRowK + 2, "J").Value = "TOTAL:"
            wsNew.Cells(lastRowK + 2, "K").Value = totalSum
        End If
    Next filterValue

    ' Add Total for Unfiltered Data
    Set wsNew = wbNew.Sheets("Unfiltered Data")
    lastRowK = wsNew.Cells(wsNew.Rows.Count, "K").End(xlUp).Row
    If lastRowK >= 2 Then ' Ensure there's data in Column K
        totalSum = Application.WorksheetFunction.Sum(wsNew.Range("K2:K" & lastRowK))
        wsNew.Cells(lastRowK + 2, "J").Value = "TOTAL:"
        wsNew.Cells(lastRowK + 2, "K").Value = totalSum
    End If

    Application.ScreenUpdating = True

    ' Set the title of the new workbook
    wbNew.SaveAs Application.DefaultFilePath & "\" & workbookTitle & ".xlsx"

    ' Notify user and activate the new workbook
    MsgBox "Data successfully split into a new workbook titled '" & workbookTitle & "'.", vbInformation
    wbNew.Activate
End Sub
