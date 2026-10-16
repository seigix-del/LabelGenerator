Attribute VB_Name = "modUtils"
Option Compare Database
Option Explicit

' ============================================================================
' modUtils - text normalisation, file dialogs and Excel reading.
' ============================================================================

Private Const msoFileDialogFilePicker As Long = 3
Private Const msoFileDialogSaveAs As Long = 2

' ---------------------------------------------------------------------------
' Text normalisation
' ---------------------------------------------------------------------------

' Cleans a cell value for storage: removes invisible direction marks,
' non-breaking spaces, tatweel and line breaks, collapses runs of spaces.
' Returns "" for Null / Empty / error values.
Public Function NormText(ByVal v As Variant) As String
    Dim s As String
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then Exit Function
    s = CStr(v)

    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW$(160), " ")    ' no-break space
    s = Replace(s, ChrW$(8206), "")    ' LRM
    s = Replace(s, ChrW$(8207), "")    ' RLM
    s = Replace(s, ChrW$(8234), "")    ' LRE
    s = Replace(s, ChrW$(8235), "")    ' RLE
    s = Replace(s, ChrW$(8236), "")    ' PDF
    s = Replace(s, ChrW$(8237), "")    ' LRO
    s = Replace(s, ChrW$(8238), "")    ' RLO
    s = Replace(s, ChrW$(1600), "")    ' Arabic tatweel
    s = Replace(s, ChrW$(65279), "")   ' BOM / zero-width no-break space

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormText = Trim$(s)
End Function

' Matching key for headers and event names: NormText + unify Arabic letter
' variants so that e.g. "akhir" written with or without madda still matches.
Public Function NormKey(ByVal v As Variant) As String
    Dim s As String
    s = NormText(v)
    s = Replace(s, ChrW$(&H622), ChrW$(&H627))   ' alef with madda      -> alef
    s = Replace(s, ChrW$(&H623), ChrW$(&H627))   ' alef with hamza above -> alef
    s = Replace(s, ChrW$(&H625), ChrW$(&H627))   ' alef with hamza below -> alef
    s = Replace(s, ChrW$(&H649), ChrW$(&H64A))   ' alef maqsura          -> yeh
    s = Replace(s, ChrW$(&H629), ChrW$(&H647))   ' taa marbuta           -> heh
    NormKey = LCase$(s)
End Function

' Waybill / tracking numbers: upper case, no spaces at all.
Public Function NormWaybill(ByVal v As Variant) As String
    NormWaybill = UCase$(Replace(NormText(v), " ", ""))
End Function

' Consignee names end with a customer code, e.g. "<school name> - M174562".
' Returns the trailing code if it looks like letters followed by digits.
Public Function ParseCustomerCode(ByVal consignee As String) As String
    Dim p As Long, tail As String
    p = InStrRev(consignee, "-")
    If p = 0 Then Exit Function
    tail = Trim$(Mid$(consignee, p + 1))
    If Len(tail) >= 2 And Len(tail) <= 30 And InStr(tail, " ") = 0 _
       And tail Like "[A-Za-z]*#" Then
        ParseCustomerCode = UCase$(tail)
    End If
End Function

' Shipment contents look like "<distribution point> \\ G9-L1".
' Returns the part after the last backslash (the product / grade code).
Public Function ParseProductCode(ByVal contents As String) As String
    Dim p As Long
    p = InStrRev(contents, "\")
    If p = 0 Then Exit Function
    ParseProductCode = Left$(Trim$(Mid$(contents, p + 1)), 50)
End Function

' ---------------------------------------------------------------------------
' Files
' ---------------------------------------------------------------------------

Public Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(path) > 0) And (Len(Dir$(path, vbNormal)) > 0)
End Function

Public Function PickExcelFile(Optional ByVal title As String = "Select Excel file") As String
    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel workbooks", "*.xlsx;*.xlsm;*.xls"
        If .Show = True Then PickExcelFile = .SelectedItems(1)
    End With
End Function

Public Function PickSavePath(ByVal defaultName As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .Title = "Save as"
        .InitialFileName = defaultName
        If .Show = True Then PickSavePath = .SelectedItems(1)
    End With
    If Len(PickSavePath) > 0 And LCase$(Right$(PickSavePath, 5)) <> ".xlsx" Then
        PickSavePath = PickSavePath & ".xlsx"
    End If
End Function

' ---------------------------------------------------------------------------
' Excel reading
' ---------------------------------------------------------------------------

' Opens the workbook read-only through Excel automation (late bound, no
' reference needed) and returns the used range of the first worksheet that
' contains a waybill header. data is a 1-based 2-D array (rows, cols);
' headerRow is the row inside data that holds the headers.
' Returns False with errMsg set when no usable sheet is found.
Public Function LoadExcelData(ByVal filePath As String, ByRef data As Variant, _
                              ByRef headerRow As Long, ByRef sheetName As String, _
                              ByRef errMsg As String) As Boolean
    On Error GoTo EH

    Dim xl As Object, wb As Object, ws As Object
    Dim v As Variant, r As Long

    Set xl = CreateObject("Excel.Application")
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    Set wb = xl.Workbooks.Open(filePath, 0, True)   ' UpdateLinks:=0, ReadOnly:=True

    For Each ws In wb.Worksheets
        v = ws.UsedRange.Value
        If IsArray(v) Then
            r = FindHeaderRow(v)
            If r > 0 Then
                data = v
                headerRow = r
                sheetName = ws.Name
                LoadExcelData = True
                Exit For
            End If
        End If
    Next ws

    If Not LoadExcelData Then
        errMsg = "No worksheet in this file contains a waybill / tracking-number column." & vbCrLf & _
                 "Expected a header such as '" & HeaderAliases(COL_WAYBILL)(0) & "' or '" & _
                 HeaderAliases(COL_WAYBILL)(1) & "'."
    End If

ExitHere:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set ws = Nothing: Set wb = Nothing: Set xl = Nothing
    Exit Function
EH:
    errMsg = "Could not read the Excel file (" & Err.Number & "): " & Err.Description
    LoadExcelData = False
    Resume ExitHere
End Function

' First row (within HEADER_SEARCH_ROWS) that contains a waybill header.
Public Function FindHeaderRow(ByRef data As Variant) As Long
    Dim r As Long, c As Long, lastRow As Long
    lastRow = UBound(data, 1)
    If lastRow > HEADER_SEARCH_ROWS Then lastRow = HEADER_SEARCH_ROWS
    For r = LBound(data, 1) To lastRow
        For c = LBound(data, 2) To UBound(data, 2)
            If MatchesAlias(data(r, c), COL_WAYBILL) Then
                FindHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
End Function

' Column index of logicalCol in the header row, or 0 when absent.
Public Function FindColumn(ByRef data As Variant, ByVal headerRow As Long, _
                           ByVal logicalCol As String) As Long
    Dim c As Long
    For c = LBound(data, 2) To UBound(data, 2)
        If MatchesAlias(data(headerRow, c), logicalCol) Then
            FindColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function MatchesAlias(ByVal cellValue As Variant, ByVal logicalCol As String) As Boolean
    Dim key As String, a As Variant
    key = NormKey(cellValue)
    If Len(key) = 0 Then Exit Function
    For Each a In HeaderAliases(logicalCol)
        If key = NormKey(a) Then
            MatchesAlias = True
            Exit Function
        End If
    Next a
End Function

' Safe cell read (returns "" when the column is missing).
Public Function CellText(ByRef data As Variant, ByVal r As Long, ByVal c As Long) As String
    If c > 0 Then CellText = NormText(data(r, c))
End Function

Public Function SqlText(ByVal s As String) As String
    SqlText = "'" & Replace(s, "'", "''") & "'"
End Function
