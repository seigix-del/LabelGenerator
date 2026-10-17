Attribute VB_Name = "modImport"
Option Compare Database
Option Explicit

' ============================================================================
' modImport - uploads an Excel file and reconciles it with the database.
'
' Two file layouts are supported (detected automatically from the headers):
'
'   MASTER  (inventory list)   bolisat al-shahn | raqm al-marja |
'                              wasf muhtawayat al-shuhna | raqm al-marja al-thanawi
'           -> new waybills are added to tblOrderDetails as IN warehouse.
'
'   STATUS  (carrier update)   raqm al-tatabbu | madinat al-wijha |
'                              akhir hadath | ism al-mursal ilayh
'           -> every waybill gets its latest event. Only the IN events in
'              modConfig.SeedEvents (ready for pickup / received from
'              customer) mean IN warehouse; any other event means OUT.
'              Status changes are written to tblStatusHistory, waybills that
'              are not in the database go to tblUnmatched.
'
' Pipeline: Excel -> staging table (VBA) -> saved action queries (SQL).
' Everything after the batch row is created runs in one transaction, so a
' failed import leaves the database unchanged.
' ============================================================================

Public Type ImportResult
    BatchID As Long
    ImportType As String
    FileName As String
    SheetName As String
    RowsRead As Long
    RowsStaged As Long
    RowsBlank As Long
    RowsDuplicate As Long
    RowsInserted As Long
    RowsExisting As Long
    RowsChanged As Long
    RowsUnchanged As Long
    RowsUnmatched As Long
    FileInCount As Long
    FileOutCount As Long
    NewEventTypes As Long
End Type

' Entry point used by frmImport. Returns True on success and fills summary
' with a human readable report.
Public Function ImportExcelFile(ByVal filePath As String, ByVal requestedType As String, _
                                ByRef summary As String) As Boolean
    On Error GoTo EH

    Dim data As Variant, headerRow As Long, sheetName As String, errMsg As String
    Dim res As ImportResult, importType As String
    Dim wrk As DAO.Workspace, db As DAO.Database, inTrans As Boolean
    Dim errNum As Long, errDesc As String

    If Not ValidateImportFile(filePath, errMsg) Then
        ShowWarning errMsg
        Exit Function
    End If

    DoCmd.Hourglass True
    SysCmd acSysCmdSetStatus, "Reading Excel file..."

    If Not LoadExcelData(filePath, data, headerRow, sheetName, errMsg) Then
        ShowWarning errMsg
        GoTo ExitHere
    End If

    importType = DetectImportType(data, headerRow, requestedType, errMsg)
    If Len(importType) = 0 Then
        ShowWarning errMsg
        GoTo ExitHere
    End If

    res.ImportType = importType
    res.FileName = Mid$(filePath, InStrRev(filePath, "\") + 1)
    res.SheetName = sheetName

    ' Large files touch many pages inside one transaction.
    DBEngine.SetOption dbMaxLocksPerFile, 500000

    Set wrk = DBEngine.Workspaces(0)
    Set db = CurrentDb
    wrk.BeginTrans
    inTrans = True

    res.BatchID = CreateBatch(db, importType, filePath)

    SysCmd acSysCmdSetStatus, "Staging rows..."
    If importType = IMPORT_MASTER Then
        StageMaster db, data, headerRow, res
        SysCmd acSysCmdSetStatus, "Updating inventory..."
        ProcessMaster db, res
    Else
        StageStatus db, data, headerRow, res
        SysCmd acSysCmdSetStatus, "Applying status events..."
        ProcessStatus db, res
    End If

    FinishBatch db, res

    wrk.CommitTrans
    inTrans = False

    summary = BuildSummary(res)
    ImportExcelFile = True

ExitHere:
    SysCmd acSysCmdClearStatus
    DoCmd.Hourglass False
    Exit Function
EH:
    errNum = Err.Number: errDesc = Err.Description
    Resume RollbackHere
RollbackHere:
    On Error Resume Next
    If inTrans Then wrk.Rollback
    inTrans = False
    On Error GoTo 0
    LogError "modImport.ImportExcelFile", errNum, errDesc
    summary = "Import failed - no changes were saved." & vbCrLf & "Error " & errNum & ": " & errDesc
    ImportExcelFile = False
    GoTo ExitHere
End Function

' ---------------------------------------------------------------------------
' Validation & detection
' ---------------------------------------------------------------------------

Public Function ValidateImportFile(ByVal filePath As String, ByRef errMsg As String) As Boolean
    Dim ext As String
    filePath = Trim$(filePath)
    If Len(filePath) = 0 Then
        errMsg = "Please choose an Excel file first."
    ElseIf Not FileExists(filePath) Then
        errMsg = "File not found:" & vbCrLf & filePath
    Else
        ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))
        If ext <> "xlsx" And ext <> "xlsm" And ext <> "xls" Then
            errMsg = "Only Excel files (.xlsx, .xlsm, .xls) can be imported."
        ElseIf Left$(Mid$(filePath, InStrRev(filePath, "\") + 1), 2) = "~$" Then
            errMsg = "That is an Excel lock file (~$...). Choose the real workbook."
        Else
            ValidateImportFile = True
        End If
    End If
End Function

Private Function DetectImportType(ByRef data As Variant, ByVal headerRow As Long, _
                                  ByVal requestedType As String, ByRef errMsg As String) As String
    Dim hasEvent As Boolean, hasMasterCols As Boolean
    hasEvent = FindColumn(data, headerRow, COL_LAST_EVENT) > 0
    hasMasterCols = FindColumn(data, headerRow, COL_REFERENCE) > 0 _
                 Or FindColumn(data, headerRow, COL_CONTENTS) > 0

    Select Case requestedType
        Case IMPORT_STATUS
            If hasEvent Then
                DetectImportType = IMPORT_STATUS
            Else
                errMsg = "This is not a status file: the column '" & _
                         HeaderAliases(COL_LAST_EVENT)(0) & "' was not found."
            End If
        Case IMPORT_MASTER
            DetectImportType = IMPORT_MASTER
        Case Else
            If hasEvent Then
                DetectImportType = IMPORT_STATUS
            ElseIf hasMasterCols Then
                DetectImportType = IMPORT_MASTER
            Else
                errMsg = "Could not tell what kind of file this is." & vbCrLf & _
                         "A status file needs '" & HeaderAliases(COL_LAST_EVENT)(0) & "'; " & _
                         "an inventory file needs '" & HeaderAliases(COL_REFERENCE)(0) & "' or '" & _
                         HeaderAliases(COL_CONTENTS)(0) & "'."
            End If
    End Select
End Function

' ---------------------------------------------------------------------------
' Batch bookkeeping
' ---------------------------------------------------------------------------

Private Function CreateBatch(ByVal db As DAO.Database, ByVal importType As String, _
                             ByVal filePath As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("tblImportBatches", dbOpenDynaset)
    rs.AddNew
    rs!ImportType = importType
    rs!FileName = Left$(filePath, 255)
    rs!ImportedBy = Left$(Environ$("USERNAME"), 100)
    rs.Update
    rs.Bookmark = rs.LastModified
    CreateBatch = rs!BatchID
    rs.Close
End Function

Private Sub FinishBatch(ByVal db As DAO.Database, ByRef res As ImportResult)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM tblImportBatches WHERE BatchID=" & res.BatchID, dbOpenDynaset)
    rs.Edit
    rs!RowsRead = res.RowsRead
    rs!RowsInserted = res.RowsInserted
    rs!RowsUpdated = res.RowsChanged
    rs!RowsUnmatched = res.RowsUnmatched
    rs!RowsSkipped = res.RowsBlank + res.RowsDuplicate + res.RowsExisting
    rs!FileInCount = res.FileInCount
    rs!FileOutCount = res.FileOutCount
    rs.Update
    rs.Close
End Sub

' ---------------------------------------------------------------------------
' MASTER file
' ---------------------------------------------------------------------------

Private Sub StageMaster(ByVal db As DAO.Database, ByRef data As Variant, _
                        ByVal headerRow As Long, ByRef res As ImportResult)
    Dim cWay As Long, cRef As Long, cCont As Long, cSec As Long
    Dim rows As Object, r As Long, wb As String, k As Variant
    Dim rs As DAO.Recordset, cont As String

    cWay = FindColumn(data, headerRow, COL_WAYBILL)
    cRef = FindColumn(data, headerRow, COL_REFERENCE)
    cCont = FindColumn(data, headerRow, COL_CONTENTS)
    cSec = FindColumn(data, headerRow, COL_SECONDARY_REF)

    ' Waybill -> last row index that contains it (later rows win).
    Set rows = CreateObject("Scripting.Dictionary")
    For r = headerRow + 1 To UBound(data, 1)
        wb = NormWaybill(data(r, cWay))
        If Len(wb) = 0 Then
            If RowHasData(data, r) Then res.RowsBlank = res.RowsBlank + 1
        Else
            res.RowsRead = res.RowsRead + 1
            If rows.Exists(wb) Then res.RowsDuplicate = res.RowsDuplicate + 1
            rows(wb) = r
        End If
    Next r

    db.QueryDefs("qryImp_ClearStageMaster").Execute dbFailOnError

    Set rs = db.OpenRecordset("tblStage_Master", dbOpenDynaset, dbAppendOnly)
    For Each k In rows.Keys
        r = rows(k)
        cont = CellText(data, r, cCont)
        rs.AddNew
        rs!BatchID = res.BatchID
        rs!Waybill = Left$(k, 50)
        ' Rows without a reference number become their own order.
        rs!ReferenceNo = Left$(IIf(Len(CellText(data, r, cRef)) > 0, CellText(data, r, cRef), k), 50)
        rs!Contents = NullIfEmpty(Left$(cont, 255))
        rs!ProductCode = NullIfEmpty(ParseProductCode(cont))
        rs!SecondaryRef = NullIfEmpty(Left$(CellText(data, r, cSec), 100))
        rs.Update
    Next k
    rs.Close
    res.RowsStaged = rows.Count
End Sub

Private Sub ProcessMaster(ByVal db As DAO.Database, ByRef res As ImportResult)
    RunAction db, "qryImp_AppendProducts"
    RunAction db, "qryImp_AppendOrders"
    RunAction db, "qryImp_UpdateOrdersSecondaryRef"
    res.RowsInserted = RunAction(db, "qryImp_AppendDetails")
    res.RowsExisting = res.RowsStaged - res.RowsInserted
End Sub

' ---------------------------------------------------------------------------
' STATUS file
' ---------------------------------------------------------------------------

Private Sub StageStatus(ByVal db As DAO.Database, ByRef data As Variant, _
                        ByVal headerRow As Long, ByRef res As ImportResult)
    Dim cWay As Long, cEvt As Long, cCity As Long, cCons As Long
    Dim rows As Object, r As Long, wb As String, k As Variant
    Dim rs As DAO.Recordset, evt As String, consignee As String

    cWay = FindColumn(data, headerRow, COL_WAYBILL)
    cEvt = FindColumn(data, headerRow, COL_LAST_EVENT)
    cCity = FindColumn(data, headerRow, COL_CITY)
    cCons = FindColumn(data, headerRow, COL_CONSIGNEE)

    Set rows = CreateObject("Scripting.Dictionary")
    For r = headerRow + 1 To UBound(data, 1)
        wb = NormWaybill(data(r, cWay))
        evt = CellText(data, r, cEvt)
        If Len(wb) = 0 Or Len(evt) = 0 Then
            ' No waybill or no event: nothing we can decide for this row.
            If RowHasData(data, r) Then res.RowsBlank = res.RowsBlank + 1
        Else
            res.RowsRead = res.RowsRead + 1
            If rows.Exists(wb) Then res.RowsDuplicate = res.RowsDuplicate + 1
            rows(wb) = r
        End If
    Next r

    db.QueryDefs("qryImp_ClearStageStatus").Execute dbFailOnError

    Set rs = db.OpenRecordset("tblStage_Status", dbOpenDynaset, dbAppendOnly)
    For Each k In rows.Keys
        r = rows(k)
        evt = CellText(data, r, cEvt)
        consignee = CellText(data, r, cCons)
        rs.AddNew
        rs!BatchID = res.BatchID
        rs!Waybill = Left$(k, 50)
        rs!LastEvent = Left$(evt, 150)
        rs!LastEventKey = Left$(NormKey(evt), 150)
        rs!CityName = NullIfEmpty(Left$(CellText(data, r, cCity), 100))
        rs!ConsigneeName = NullIfEmpty(Left$(consignee, 255))
        rs!CustomerCode = NullIfEmpty(ParseCustomerCode(consignee))
        rs.Update
    Next k
    rs.Close
    res.RowsStaged = rows.Count
End Sub

Private Sub ProcessStatus(ByVal db As DAO.Database, ByRef res As ImportResult)
    Dim rs As DAO.Recordset

    res.NewEventTypes = RunAction(db, "qryImp_AppendEventTypes")
    EnforceInRule db
    RunAction db, "qryAct_RecalcInWarehouse"            ' heal any stale flags
    RunAction db, "qryImp_AppendCities"
    RunAction db, "qryImp_AppendCustomers"
    res.RowsUnmatched = RunAction(db, "qryImp_AppendUnmatched")
    RunAction db, "qryImp_AppendHistory"                ' must run before the update
    res.RowsChanged = RunAction(db, "qryImp_UpdateDetailsStatus")
    RunAction db, "qryImp_UpdateOrdersCustomer"
    RunAction db, "qryImp_UpdateOrdersCity"

    res.RowsUnchanged = res.RowsStaged - res.RowsUnmatched - res.RowsChanged

    ' IN / OUT split of the waybills in this file that exist in the database.
    Set rs = db.OpenRecordset("qryRpt_LastStatusFileSplit", dbOpenSnapshot)
    If Not rs.EOF Then
        res.FileInCount = Nz(rs!InWarehouse, 0)
        res.FileOutCount = Nz(rs!OutOfWarehouse, 0)
    End If
    rs.Close
End Sub

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

Private Function RunAction(ByVal db As DAO.Database, ByVal queryName As String) As Long
    Dim qdf As DAO.QueryDef
    Set qdf = db.QueryDefs(queryName)
    qdf.Execute dbFailOnError
    RunAction = qdf.RecordsAffected
    qdf.Close
End Function

Private Function NullIfEmpty(ByVal s As String) As Variant
    If Len(s) = 0 Then NullIfEmpty = Null Else NullIfEmpty = s
End Function

Private Function RowHasData(ByRef data As Variant, ByVal r As Long) As Boolean
    Dim c As Long
    For c = LBound(data, 2) To UBound(data, 2)
        If Len(NormText(data(r, c))) > 0 Then
            RowHasData = True
            Exit Function
        End If
    Next c
End Function

Private Function BuildSummary(ByRef res As ImportResult) As String
    Dim s As String
    s = "File: " & res.FileName & "  (sheet: " & res.SheetName & ")" & vbCrLf & _
        "Type: " & IIf(res.ImportType = IMPORT_MASTER, "Inventory list (master)", "Status update") & _
        "   Batch #" & res.BatchID & vbCrLf & vbCrLf & _
        "Rows with a waybill:      " & res.RowsRead & vbCrLf & _
        "Duplicate waybills:       " & res.RowsDuplicate & " (last row kept)" & vbCrLf & _
        "Rows skipped (incomplete): " & res.RowsBlank & vbCrLf

    If res.ImportType = IMPORT_MASTER Then
        s = s & "New waybills added (IN):  " & res.RowsInserted & vbCrLf & _
                "Already in database:      " & res.RowsExisting & vbCrLf
    Else
        s = s & "Status changed:           " & res.RowsChanged & vbCrLf & _
                "No change:                " & res.RowsUnchanged & vbCrLf & _
                "Not found in database:    " & res.RowsUnmatched & "  (see Unmatched)" & vbCrLf & _
                "New event types created:  " & res.NewEventTypes & IIf(res.NewEventTypes > 0, "  (counted as OUT)", "") & vbCrLf & vbCrLf & _
                "This file -> IN warehouse:  " & res.FileInCount & vbCrLf & _
                "This file -> OUT:           " & res.FileOutCount & vbCrLf
    End If

    s = s & vbCrLf & "Inventory now -> IN: " & Kpi_InWarehouse() & "   OUT: " & Kpi_OutOfWarehouse() & _
            "   Total: " & Kpi_Total()
    BuildSummary = s
End Function

' Sets tblEventTypes.IsInWarehouse from the fixed rule: True only for the
' IN events in SeedEvents(), False for every other event.
Public Sub EnforceInRule(ByVal db As DAO.Database)
    Dim e As Variant, inKeys As String
    For Each e In SeedEvents()
        If e(1) Then inKeys = inKeys & IIf(Len(inKeys) > 0, ",", "") & SqlText(NormKey(e(0)))
    Next e
    db.Execute "UPDATE tblEventTypes SET IsInWarehouse = True WHERE EventKey IN (" & inKeys & ") AND IsInWarehouse = False", dbFailOnError
    db.Execute "UPDATE tblEventTypes SET IsInWarehouse = False WHERE EventKey NOT IN (" & inKeys & ") AND IsInWarehouse = True", dbFailOnError
End Sub

' Applies the IN/OUT rule to every waybill that already has an event.
Public Function RecalcInWarehouse() As Long
    On Error GoTo EH
    Dim db As DAO.Database
    Set db = CurrentDb
    EnforceInRule db
    RecalcInWarehouse = RunAction(db, "qryAct_RecalcInWarehouse")
ExitHere:
    Exit Function
EH:
    LogError "modImport.RecalcInWarehouse", Err.Number, Err.Description
    Resume ExitHere
End Function
