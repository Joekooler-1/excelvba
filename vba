Sub FilterAndSplitDataWithWorkbookTitle()
    Dim wsSource As Worksheet
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim lastRow As Long, lastRowK As Long
    Dim uniqueValues As Collection
    Dim cell As Range
    Dim filterValue As Variant ' Use Variant to handle any data type
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

    ' Add the unfiltered data as the first sheet
    Set wsNew = wbNew.Sheets.Add
    wsNew.Name = "Unfiltered Data"
    wsSource.UsedRange.Copy Destination:=wsNew.Range("A1")

    ' Delete all default sheets after adding "Unfiltered Data"
    Application.DisplayAlerts = False
    Do While wbNew.Sheets.Count > 1
        wbNew.Sheets(2).Delete
    Loop
    Application.DisplayAlerts = True

    ' Create a collection to store unique values in Column B
    Set uniqueValues = New Collection
    On Error Resume Next ' Avoid errors from duplicate keys
    For Each cell In wsSource.Range("B2:B" & lastRow) ' Assuming data starts from Row 2
        If cell.Value <> "" Then uniqueValues.Add cell.Value, CStr(cell.Value)
    Next cell
    On Error GoTo 0

    ' Loop through unique values and create new sheets in the new workbook
    Application.ScreenUpdating = False
    For Each filterValue In uniqueValues
        ' Add a new worksheet to the new workbook
        Set wsNew = wbNew.Sheets.Add(After:=wbNew.Sheets(wbNew.Sheets.Count))

        ' Truncate and sanitize the filter value to create a valid sheet name
        sheetName = Left(filterValue, 16)
        sheetName = Application.Substitute(sheetName, "/", "_")
        sheetName = Application.Substitute(sheetName, "\", "_")
        sheetName = Application.Substitute(sheetName, ":", "_")
        sheetName = Application.Substitute(sheetName, "?", "_")
        sheetName = Application.Substitute(sheetName, "*", "_")
        sheetName = Application.Substitute(sheetName, "[", "_")
        sheetName = Application.Substitute(sheetName, "]", "_")
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
            wsNew.Cells(lastRowK + 2, "K").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)" ' Set to Accounting format
        End If
    Next filterValue

    Application.ScreenUpdating = True

    ' Set the title of the new workbook
    wbNew.SaveAs Application.DefaultFilePath & "\" & workbookTitle & ".xlsx"

    ' Notify user and activate the new workbook
    MsgBox "Data successfully split into a new workbook titled '" & workbookTitle & "'.", vbInformation
    wbNew.Activate
End Sub
