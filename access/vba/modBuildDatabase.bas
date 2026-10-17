Attribute VB_Name = "modBuildDatabase"
Option Compare Database
Option Explicit

' ============================================================================
' modBuildDatabase - creates the whole application inside the current .accdb:
' tables + keys, seed data, saved queries, forms and startup settings.
'
' Usage (once, in a new blank database):
'     1. Import every .bas file from access\vba into the VBA editor.
'     2. Immediate window (Ctrl+G):  BuildInventoryDatabase
'
' Re-running is safe for queries and forms (they are recreated). Tables are
' only created when missing, so existing data is never dropped unless you
' call BuildInventoryDatabase True.
'
' The SQL in this module is mirrored in access\sql\*.sql for reading.
' ============================================================================

Private Const TW As Long = 567          ' twips per centimetre

Private Const CLR_HEADER As Long = 6299648   ' dark blue  RGB(0,32,96)
Private Const CLR_WHITE As Long = 16777215
Private Const CLR_IN As Long = 5287936       ' green      RGB(0,176,80)
Private Const CLR_OUT As Long = 255          ' red        RGB(255,0,0)
Private Const CLR_PANEL As Long = 15921906   ' light grey RGB(242,242,242)

Public Sub BuildInventoryDatabase(Optional ByVal dropExistingTables As Boolean = False)
    On Error GoTo EH

    DoCmd.Hourglass True
    DoCmd.SetWarnings False

    If dropExistingTables Then DropTables
    CreateTables
    UpgradeSchema                    ' add columns introduced after the first release
    DBEngine.Idle dbRefreshCache     ' make ADO-created tables visible to DAO
    SeedData
    CreateQueries
    RecalcInWarehouse                ' apply the fixed IN/OUT rule to existing data
    CreateForms
    SetStartupOptions

    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    Application.RefreshDatabaseWindow
    MsgBox "Inventory database built successfully." & vbCrLf & _
           "Close and reopen the file to start on the main menu.", vbInformation, APP_TITLE
    Exit Sub
EH:
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    LogError "modBuildDatabase.BuildInventoryDatabase", Err.Number, Err.Description
End Sub

' ===========================================================================
' 1. TABLES
' DDL runs through ADO (CurrentProject.Connection) because Jet/ACE only
' accepts DEFAULT and ON DELETE CASCADE in ANSI-92 mode.
' ===========================================================================

Private Sub CreateTables()
    ' --- lookup / master tables --------------------------------------------
    RunDDL "tblCustomers", "CREATE TABLE tblCustomers (CustomerID COUNTER CONSTRAINT PK_Customers PRIMARY KEY, CustomerName TEXT(255) NOT NULL, CustomerCode TEXT(50), CreatedOn DATETIME DEFAULT Now())"
    RunDDL "", "CREATE UNIQUE INDEX UX_Customers_Name ON tblCustomers (CustomerName)"
    RunDDL "", "CREATE INDEX IX_Customers_Code ON tblCustomers (CustomerCode)"

    RunDDL "tblCities", "CREATE TABLE tblCities (CityID COUNTER CONSTRAINT PK_Cities PRIMARY KEY, CityName TEXT(100) NOT NULL)"
    RunDDL "", "CREATE UNIQUE INDEX UX_Cities_Name ON tblCities (CityName)"

    RunDDL "tblProducts", "CREATE TABLE tblProducts (ProductID COUNTER CONSTRAINT PK_Products PRIMARY KEY, ProductDescription TEXT(255) NOT NULL, ProductCode TEXT(50))"
    RunDDL "", "CREATE UNIQUE INDEX UX_Products_Desc ON tblProducts (ProductDescription)"

    RunDDL "tblEventTypes", "CREATE TABLE tblEventTypes (EventID COUNTER CONSTRAINT PK_EventTypes PRIMARY KEY, EventName TEXT(150) NOT NULL, EventKey TEXT(150) NOT NULL, IsInWarehouse YESNO DEFAULT 0, SortOrder LONG DEFAULT 100, Notes TEXT(255))"
    RunDDL "", "CREATE UNIQUE INDEX UX_EventTypes_Key ON tblEventTypes (EventKey)"

    RunDDL "tblImportBatches", "CREATE TABLE tblImportBatches (BatchID COUNTER CONSTRAINT PK_ImportBatches PRIMARY KEY, ImportType TEXT(20) NOT NULL, FileName TEXT(255), ImportedOn DATETIME DEFAULT Now(), ImportedBy TEXT(100), RowsRead LONG DEFAULT 0, RowsInserted LONG DEFAULT 0, RowsUpdated LONG DEFAULT 0, RowsUnmatched LONG DEFAULT 0, RowsSkipped LONG DEFAULT 0, FileInCount LONG DEFAULT 0, FileOutCount LONG DEFAULT 0)"

    ' --- transactional tables ---------------------------------------------
    RunDDL "tblOrders", "CREATE TABLE tblOrders (OrderID COUNTER CONSTRAINT PK_Orders PRIMARY KEY, ReferenceNo TEXT(50) NOT NULL, SecondaryRef TEXT(100), CustomerID LONG, CityID LONG, CreatedBatchID LONG, CreatedOn DATETIME DEFAULT Now(), " & _
        "CONSTRAINT FK_Orders_Customers FOREIGN KEY (CustomerID) REFERENCES tblCustomers (CustomerID), " & _
        "CONSTRAINT FK_Orders_Cities FOREIGN KEY (CityID) REFERENCES tblCities (CityID), " & _
        "CONSTRAINT FK_Orders_Batches FOREIGN KEY (CreatedBatchID) REFERENCES tblImportBatches (BatchID))"
    RunDDL "", "CREATE UNIQUE INDEX UX_Orders_Ref ON tblOrders (ReferenceNo)"

    RunDDL "tblOrderDetails", "CREATE TABLE tblOrderDetails (DetailID COUNTER CONSTRAINT PK_OrderDetails PRIMARY KEY, Waybill TEXT(50) NOT NULL, OrderID LONG NOT NULL, ProductID LONG, Quantity LONG DEFAULT 1, GradeCode TEXT(10), LossQty LONG DEFAULT 0, CurrentEventID LONG, InWarehouse YESNO DEFAULT -1, LastUpdated DATETIME DEFAULT Now(), LastBatchID LONG, CreatedOn DATETIME DEFAULT Now(), " & _
        "CONSTRAINT FK_Details_Orders FOREIGN KEY (OrderID) REFERENCES tblOrders (OrderID) ON DELETE CASCADE, " & _
        "CONSTRAINT FK_Details_Products FOREIGN KEY (ProductID) REFERENCES tblProducts (ProductID), " & _
        "CONSTRAINT FK_Details_Events FOREIGN KEY (CurrentEventID) REFERENCES tblEventTypes (EventID), " & _
        "CONSTRAINT FK_Details_Batches FOREIGN KEY (LastBatchID) REFERENCES tblImportBatches (BatchID))"
    RunDDL "", "CREATE UNIQUE INDEX UX_Details_Waybill ON tblOrderDetails (Waybill)"
    RunDDL "", "CREATE INDEX IX_Details_InWarehouse ON tblOrderDetails (InWarehouse)"

    RunDDL "tblStatusHistory", "CREATE TABLE tblStatusHistory (HistoryID COUNTER CONSTRAINT PK_StatusHistory PRIMARY KEY, DetailID LONG NOT NULL, EventID LONG NOT NULL, PrevEventID LONG, WasInWarehouse YESNO, IsInWarehouse YESNO, BatchID LONG, RecordedOn DATETIME DEFAULT Now(), " & _
        "CONSTRAINT FK_History_Details FOREIGN KEY (DetailID) REFERENCES tblOrderDetails (DetailID) ON DELETE CASCADE, " & _
        "CONSTRAINT FK_History_Events FOREIGN KEY (EventID) REFERENCES tblEventTypes (EventID), " & _
        "CONSTRAINT FK_History_PrevEvents FOREIGN KEY (PrevEventID) REFERENCES tblEventTypes (EventID), " & _
        "CONSTRAINT FK_History_Batches FOREIGN KEY (BatchID) REFERENCES tblImportBatches (BatchID))"

    RunDDL "tblUnmatched", "CREATE TABLE tblUnmatched (UnmatchedID COUNTER CONSTRAINT PK_Unmatched PRIMARY KEY, BatchID LONG NOT NULL, Waybill TEXT(50), LastEvent TEXT(150), CityName TEXT(100), ConsigneeName TEXT(255), RecordedOn DATETIME DEFAULT Now(), " & _
        "CONSTRAINT FK_Unmatched_Batches FOREIGN KEY (BatchID) REFERENCES tblImportBatches (BatchID) ON DELETE CASCADE)"
    RunDDL "", "CREATE INDEX IX_Unmatched_Waybill ON tblUnmatched (Waybill)"

    ' --- staging (raw rows of the last uploaded file) ----------------------
    RunDDL "tblStage_Master", "CREATE TABLE tblStage_Master (StageID COUNTER CONSTRAINT PK_StageMaster PRIMARY KEY, BatchID LONG, Waybill TEXT(50) NOT NULL, ReferenceNo TEXT(50), Contents TEXT(255), ProductCode TEXT(50), SecondaryRef TEXT(100), GradeCode TEXT(10), LossQty LONG DEFAULT 0)"
    RunDDL "", "CREATE UNIQUE INDEX UX_StageMaster_Waybill ON tblStage_Master (Waybill)"

    RunDDL "tblStage_Status", "CREATE TABLE tblStage_Status (StageID COUNTER CONSTRAINT PK_StageStatus PRIMARY KEY, BatchID LONG, Waybill TEXT(50) NOT NULL, LastEvent TEXT(150), LastEventKey TEXT(150), CityName TEXT(100), ConsigneeName TEXT(255), CustomerCode TEXT(50))"
    RunDDL "", "CREATE UNIQUE INDEX UX_StageStatus_Waybill ON tblStage_Status (Waybill)"

    ' --- infrastructure ----------------------------------------------------
    RunDDL "tblErrorLog", "CREATE TABLE tblErrorLog (ErrorID COUNTER CONSTRAINT PK_ErrorLog PRIMARY KEY, ErrNumber LONG, ErrDescription MEMO, ErrSource TEXT(255), UserName TEXT(100), LoggedOn DATETIME DEFAULT Now())"
End Sub

' Runs a DDL statement. When tableName is given, the statement is skipped
' if that table already exists; index statements that follow a skipped
' table are skipped as well.
Private Sub RunDDL(ByVal tableName As String, ByVal sql As String)
    Static skipping As Boolean
    If Len(tableName) > 0 Then
        skipping = TableExists(tableName)
        If skipping Then Debug.Print "exists: "; tableName
    End If
    If skipping Then Exit Sub
    CurrentProject.Connection.Execute sql
End Sub

' Databases built before grades/losses existed get the new columns here.
Private Sub UpgradeSchema()
    EnsureColumn "tblOrderDetails", "GradeCode", "TEXT(10)"
    EnsureColumn "tblOrderDetails", "LossQty", "LONG DEFAULT 0"
    EnsureColumn "tblStage_Master", "GradeCode", "TEXT(10)"
    EnsureColumn "tblStage_Master", "LossQty", "LONG DEFAULT 0"
    CurrentProject.Connection.Execute "UPDATE tblOrderDetails SET LossQty = 0 WHERE LossQty Is Null"
    On Error Resume Next             ' index may already exist
    CurrentProject.Connection.Execute "CREATE INDEX IX_Details_Grade ON tblOrderDetails (GradeCode)"
End Sub

Private Sub EnsureColumn(ByVal tableName As String, ByVal columnName As String, ByVal columnType As String)
    Dim fld As DAO.Field
    CurrentDb.TableDefs.Refresh
    For Each fld In CurrentDb.TableDefs(tableName).Fields
        If fld.Name = columnName Then Exit Sub
    Next fld
    CurrentProject.Connection.Execute "ALTER TABLE " & tableName & " ADD COLUMN " & columnName & " " & columnType
    Debug.Print "added column: "; tableName; "."; columnName
End Sub

Private Function TableExists(ByVal tableName As String) As Boolean
    Dim td As DAO.TableDef
    CurrentDb.TableDefs.Refresh
    For Each td In CurrentDb.TableDefs
        If td.Name = tableName Then
            TableExists = True
            Exit Function
        End If
    Next td
End Function

Private Sub DropTables()
    ' Children first so foreign keys never block the drop.
    Dim t As Variant
    For Each t In Array("tblStatusHistory", "tblUnmatched", "tblOrderDetails", "tblOrders", _
                        "tblStage_Master", "tblStage_Status", "tblImportBatches", "tblEventTypes", _
                        "tblProducts", "tblCities", "tblCustomers", "tblErrorLog")
        If TableExists(CStr(t)) Then CurrentDb.Execute "DROP TABLE " & t, dbFailOnError
    Next t
End Sub

' ===========================================================================
' 2. SEED DATA
' ===========================================================================

Private Sub SeedData()
    Dim db As DAO.Database, rs As DAO.Recordset, e As Variant, key As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblEventTypes", dbOpenDynaset)
    For Each e In SeedEvents()
        key = NormKey(e(0))
        rs.FindFirst "EventKey=" & SqlText(key)
        If rs.NoMatch Then
            rs.AddNew
            rs!EventName = e(0)
            rs!EventKey = key
            rs!IsInWarehouse = e(1)
            rs!SortOrder = e(2)
            rs!Notes = IIf(e(1), "Counts as IN warehouse", "Counts as OUT of warehouse")
            rs.Update
        End If
    Next e
    rs.Close
End Sub

' ===========================================================================
' 3. SAVED QUERIES  (DAO / ANSI-89 syntax)
' ===========================================================================

Private Sub CreateQueries()
    ' ---- import: master file ---------------------------------------------
    SaveQuery "qryImp_ClearStageMaster", "DELETE FROM tblStage_Master"
    SaveQuery "qryImp_AppendProducts", "INSERT INTO tblProducts (ProductDescription, ProductCode) SELECT s.Contents, First(s.ProductCode) AS FirstCode FROM tblStage_Master AS s LEFT JOIN tblProducts AS p ON s.Contents = p.ProductDescription WHERE p.ProductID Is Null AND s.Contents Is Not Null GROUP BY s.Contents"
    SaveQuery "qryImp_AppendOrders", "INSERT INTO tblOrders (ReferenceNo, SecondaryRef, CreatedBatchID) SELECT s.ReferenceNo, First(s.SecondaryRef) AS FirstSecRef, First(s.BatchID) AS FirstBatch FROM tblStage_Master AS s LEFT JOIN tblOrders AS o ON s.ReferenceNo = o.ReferenceNo WHERE o.OrderID Is Null AND s.ReferenceNo Is Not Null GROUP BY s.ReferenceNo"
    SaveQuery "qryImp_UpdateOrdersSecondaryRef", "UPDATE tblOrders AS o INNER JOIN tblStage_Master AS s ON o.ReferenceNo = s.ReferenceNo SET o.SecondaryRef = s.SecondaryRef WHERE s.SecondaryRef Is Not Null"
    SaveQuery "qryImp_AppendDetails", "INSERT INTO tblOrderDetails (Waybill, OrderID, ProductID, Quantity, GradeCode, LossQty, InWarehouse, LastBatchID) SELECT s.Waybill, o.OrderID, p.ProductID, 1, s.GradeCode, IIf(s.LossQty Is Null, 0, s.LossQty), True, s.BatchID FROM ((tblStage_Master AS s INNER JOIN tblOrders AS o ON s.ReferenceNo = o.ReferenceNo) LEFT JOIN tblProducts AS p ON s.Contents = p.ProductDescription) LEFT JOIN tblOrderDetails AS d ON s.Waybill = d.Waybill WHERE d.DetailID Is Null"

    SaveQuery "qryImp_UpdateDetailsGrade", "UPDATE tblOrderDetails AS d INNER JOIN tblStage_Master AS s ON d.Waybill = s.Waybill SET d.GradeCode = s.GradeCode, d.LossQty = IIf(s.LossQty Is Null, 0, s.LossQty)"

    ' ---- import: status file ---------------------------------------------
    SaveQuery "qryImp_ClearStageStatus", "DELETE FROM tblStage_Status"
    SaveQuery "qryImp_AppendEventTypes", "INSERT INTO tblEventTypes (EventName, EventKey, IsInWarehouse, SortOrder, Notes) SELECT First(s.LastEvent) AS FirstName, s.LastEventKey, False, 100, 'Auto-added on import - review' FROM tblStage_Status AS s LEFT JOIN tblEventTypes AS e ON s.LastEventKey = e.EventKey WHERE e.EventID Is Null AND s.LastEventKey Is Not Null GROUP BY s.LastEventKey"
    SaveQuery "qryImp_AppendCities", "INSERT INTO tblCities (CityName) SELECT s.CityName FROM tblStage_Status AS s LEFT JOIN tblCities AS c ON s.CityName = c.CityName WHERE c.CityID Is Null AND s.CityName Is Not Null GROUP BY s.CityName"
    SaveQuery "qryImp_AppendCustomers", "INSERT INTO tblCustomers (CustomerName, CustomerCode) SELECT s.ConsigneeName, First(s.CustomerCode) AS FirstCode FROM tblStage_Status AS s LEFT JOIN tblCustomers AS c ON s.ConsigneeName = c.CustomerName WHERE c.CustomerID Is Null AND s.ConsigneeName Is Not Null GROUP BY s.ConsigneeName"
    SaveQuery "qryImp_AppendUnmatched", "INSERT INTO tblUnmatched (BatchID, Waybill, LastEvent, CityName, ConsigneeName) SELECT s.BatchID, s.Waybill, s.LastEvent, s.CityName, s.ConsigneeName FROM tblStage_Status AS s LEFT JOIN tblOrderDetails AS d ON s.Waybill = d.Waybill WHERE d.DetailID Is Null"
    SaveQuery "qryImp_AppendHistory", "INSERT INTO tblStatusHistory (DetailID, EventID, PrevEventID, WasInWarehouse, IsInWarehouse, BatchID) SELECT d.DetailID, e.EventID, d.CurrentEventID, d.InWarehouse, e.IsInWarehouse, s.BatchID FROM (tblStage_Status AS s INNER JOIN tblOrderDetails AS d ON s.Waybill = d.Waybill) INNER JOIN tblEventTypes AS e ON s.LastEventKey = e.EventKey WHERE d.CurrentEventID Is Null OR d.CurrentEventID <> e.EventID"
    SaveQuery "qryImp_UpdateDetailsStatus", "UPDATE (tblOrderDetails AS d INNER JOIN tblStage_Status AS s ON d.Waybill = s.Waybill) INNER JOIN tblEventTypes AS e ON s.LastEventKey = e.EventKey SET d.CurrentEventID = e.EventID, d.InWarehouse = e.IsInWarehouse, d.LastUpdated = Now(), d.LastBatchID = s.BatchID WHERE d.CurrentEventID Is Null OR d.CurrentEventID <> e.EventID"
    SaveQuery "qryImp_UpdateOrdersCustomer", "UPDATE ((tblOrders AS o INNER JOIN tblOrderDetails AS d ON o.OrderID = d.OrderID) INNER JOIN tblStage_Status AS s ON d.Waybill = s.Waybill) INNER JOIN tblCustomers AS c ON s.ConsigneeName = c.CustomerName SET o.CustomerID = c.CustomerID"
    SaveQuery "qryImp_UpdateOrdersCity", "UPDATE ((tblOrders AS o INNER JOIN tblOrderDetails AS d ON o.OrderID = d.OrderID) INNER JOIN tblStage_Status AS s ON d.Waybill = s.Waybill) INNER JOIN tblCities AS c ON s.CityName = c.CityName SET o.CityID = c.CityID"

    ' ---- maintenance -----------------------------------------------------
    SaveQuery "qryAct_RecalcInWarehouse", "UPDATE tblOrderDetails AS d INNER JOIN tblEventTypes AS e ON d.CurrentEventID = e.EventID SET d.InWarehouse = e.IsInWarehouse WHERE d.InWarehouse <> e.IsInWarehouse"
    SaveQuery "qryAct_PurgeOldUnmatched", "DELETE FROM tblUnmatched WHERE RecordedOn < DateAdd('d', -90, Date())"

    ' ---- data entry / lists ----------------------------------------------
    SaveQuery "qryShipmentsList", "SELECT d.DetailID, d.Waybill, o.ReferenceNo, o.SecondaryRef, p.ProductDescription, p.ProductCode, d.GradeCode, d.LossQty, d.Quantity, e.EventName, d.InWarehouse, IIf(d.InWarehouse, 'IN', 'OUT') AS Location, ci.CityName, cu.CustomerName, cu.CustomerCode, d.LastUpdated " & _
        "FROM ((((tblOrderDetails AS d INNER JOIN tblOrders AS o ON d.OrderID = o.OrderID) LEFT JOIN tblProducts AS p ON d.ProductID = p.ProductID) LEFT JOIN tblEventTypes AS e ON d.CurrentEventID = e.EventID) LEFT JOIN tblCities AS ci ON o.CityID = ci.CityID) LEFT JOIN tblCustomers AS cu ON o.CustomerID = cu.CustomerID " & _
        "ORDER BY d.Waybill"
    SaveQuery "qryShipmentHistory", "SELECT h.HistoryID, h.DetailID, d.Waybill, pe.EventName AS FromEvent, ne.EventName AS ToEvent, IIf(h.WasInWarehouse, 'IN', 'OUT') AS FromLocation, IIf(h.IsInWarehouse, 'IN', 'OUT') AS ToLocation, h.RecordedOn, b.FileName " & _
        "FROM (((tblStatusHistory AS h INNER JOIN tblOrderDetails AS d ON h.DetailID = d.DetailID) INNER JOIN tblEventTypes AS ne ON h.EventID = ne.EventID) LEFT JOIN tblEventTypes AS pe ON h.PrevEventID = pe.EventID) LEFT JOIN tblImportBatches AS b ON h.BatchID = b.BatchID " & _
        "ORDER BY h.RecordedOn DESC, h.HistoryID DESC"
    SaveQuery "qryEventTypesList", "SELECT EventID, EventName, EventKey, IsInWarehouse, SortOrder, Notes FROM tblEventTypes ORDER BY SortOrder, EventName"
    SaveQuery "qryImportBatchesList", "SELECT BatchID, ImportType, ImportedOn, ImportedBy, FileName, RowsRead, RowsInserted, RowsUpdated, RowsUnmatched, RowsSkipped, FileInCount, FileOutCount FROM tblImportBatches ORDER BY BatchID DESC"

    ' ---- reporting -------------------------------------------------------
    SaveQuery "qryRpt_InventorySummary", "SELECT Sum(IIf(d.InWarehouse, d.Quantity, 0)) AS InWarehouse, Sum(IIf(d.InWarehouse, 0, d.Quantity)) AS OutOfWarehouse, Sum(d.Quantity) AS TotalQty, Count(d.DetailID) AS Waybills FROM tblOrderDetails AS d"
    SaveQuery "qryRpt_ByEvent", "SELECT IIf(e.EventName Is Null, '(no status yet)', e.EventName) AS EventLabel, IIf(d.InWarehouse, 'IN', 'OUT') AS Location, Sum(d.Quantity) AS Qty FROM tblOrderDetails AS d LEFT JOIN tblEventTypes AS e ON d.CurrentEventID = e.EventID GROUP BY IIf(e.EventName Is Null, '(no status yet)', e.EventName), IIf(d.InWarehouse, 'IN', 'OUT'), e.SortOrder ORDER BY e.SortOrder, Sum(d.Quantity) DESC"
    SaveQuery "qryRpt_ByProduct", "SELECT IIf(p.ProductDescription Is Null, '(unknown)', p.ProductDescription) AS Product, Sum(IIf(d.InWarehouse, d.Quantity, 0)) AS InQty, Sum(IIf(d.InWarehouse, 0, d.Quantity)) AS OutQty, Sum(d.Quantity) AS TotalQty FROM tblOrderDetails AS d LEFT JOIN tblProducts AS p ON d.ProductID = p.ProductID GROUP BY IIf(p.ProductDescription Is Null, '(unknown)', p.ProductDescription) ORDER BY Sum(IIf(d.InWarehouse, d.Quantity, 0)) DESC"
    SaveQuery "qryRpt_ByCity", "SELECT IIf(c.CityName Is Null, '(unknown)', c.CityName) AS City, Sum(IIf(d.InWarehouse, d.Quantity, 0)) AS InQty, Sum(IIf(d.InWarehouse, 0, d.Quantity)) AS OutQty FROM (tblOrderDetails AS d INNER JOIN tblOrders AS o ON d.OrderID = o.OrderID) LEFT JOIN tblCities AS c ON o.CityID = c.CityID GROUP BY IIf(c.CityName Is Null, '(unknown)', c.CityName) ORDER BY Sum(IIf(d.InWarehouse, d.Quantity, 0)) DESC"
    ' Grades: a plain G<n> row counts 1 towards that grade's quantity; a G<n>-L<k>
    ' row is a loss row - it adds k to Loss and nothing to the quantity.
    SaveQuery "qryRpt_ByGrade", "SELECT IIf(d.GradeCode Is Null, '(no grade)', d.GradeCode) AS Grade, Sum(IIf(d.LossQty > 0, 0, d.Quantity)) AS Qty, Sum(d.LossQty) AS Loss, Sum(IIf(d.LossQty > 0, 0, IIf(d.InWarehouse, d.Quantity, 0))) AS InQty, Sum(IIf(d.LossQty > 0, 0, IIf(d.InWarehouse, 0, d.Quantity))) AS OutQty " & _
        "FROM tblOrderDetails AS d GROUP BY IIf(d.GradeCode Is Null, '(no grade)', d.GradeCode), IIf(d.GradeCode Is Null, 999, Val(Mid(d.GradeCode, 2))) ORDER BY IIf(d.GradeCode Is Null, 999, Val(Mid(d.GradeCode, 2)))"
    SaveQuery "qryRpt_ByConsigneeGrade", "SELECT IIf(cu.CustomerName Is Null, '(no consignee yet)', cu.CustomerName) AS Consignee, cu.CustomerCode, IIf(d.GradeCode Is Null, '(no grade)', d.GradeCode) AS Grade, Sum(IIf(d.LossQty > 0, 0, d.Quantity)) AS Qty, Sum(d.LossQty) AS Loss, Sum(IIf(d.LossQty > 0, 0, IIf(d.InWarehouse, d.Quantity, 0))) AS InQty, Sum(IIf(d.LossQty > 0, 0, IIf(d.InWarehouse, 0, d.Quantity))) AS OutQty " & _
        "FROM (tblOrderDetails AS d INNER JOIN tblOrders AS o ON d.OrderID = o.OrderID) LEFT JOIN tblCustomers AS cu ON o.CustomerID = cu.CustomerID " & _
        "GROUP BY IIf(cu.CustomerName Is Null, '(no consignee yet)', cu.CustomerName), cu.CustomerCode, IIf(d.GradeCode Is Null, '(no grade)', d.GradeCode), IIf(d.GradeCode Is Null, 999, Val(Mid(d.GradeCode, 2))) " & _
        "ORDER BY IIf(cu.CustomerName Is Null, '(no consignee yet)', cu.CustomerName), IIf(d.GradeCode Is Null, 999, Val(Mid(d.GradeCode, 2)))"
    SaveQuery "qryRpt_ConsigneeGradeQty", "TRANSFORM Sum(IIf(d.LossQty > 0, 0, d.Quantity)) AS Qty SELECT IIf(cu.CustomerName Is Null, '(no consignee yet)', cu.CustomerName) AS Consignee, Sum(IIf(d.LossQty > 0, 0, d.Quantity)) AS TotalQty " & _
        "FROM (tblOrderDetails AS d INNER JOIN tblOrders AS o ON d.OrderID = o.OrderID) LEFT JOIN tblCustomers AS cu ON o.CustomerID = cu.CustomerID " & _
        "GROUP BY IIf(cu.CustomerName Is Null, '(no consignee yet)', cu.CustomerName) PIVOT d.GradeCode IN ('G1','G2','G3','G4','G5','G6','G7','G8','G9','G10','G11','G12','G13')"
    SaveQuery "qryRpt_ConsigneeGradeLoss", "TRANSFORM Sum(d.LossQty) AS Loss SELECT IIf(cu.CustomerName Is Null, '(no consignee yet)', cu.CustomerName) AS Consignee, Sum(d.LossQty) AS TotalLoss " & _
        "FROM (tblOrderDetails AS d INNER JOIN tblOrders AS o ON d.OrderID = o.OrderID) LEFT JOIN tblCustomers AS cu ON o.CustomerID = cu.CustomerID " & _
        "GROUP BY IIf(cu.CustomerName Is Null, '(no consignee yet)', cu.CustomerName) PIVOT d.GradeCode IN ('G1','G2','G3','G4','G5','G6','G7','G8','G9','G10','G11','G12','G13')"
    SaveQuery "qryRpt_InWarehouse", "SELECT Waybill, ReferenceNo, SecondaryRef, ProductDescription, EventName, CityName, CustomerName, LastUpdated FROM qryShipmentsList WHERE InWarehouse = True"
    SaveQuery "qryRpt_OutOfWarehouse", "SELECT Waybill, ReferenceNo, SecondaryRef, ProductDescription, EventName, CityName, CustomerName, LastUpdated FROM qryShipmentsList WHERE InWarehouse = False"
    SaveQuery "qryRpt_LastStatusBatch", "SELECT TOP 1 * FROM tblImportBatches WHERE ImportType = 'Status' ORDER BY BatchID DESC"
    SaveQuery "qryRpt_LastStatusFileSplit", "SELECT Sum(IIf(e.IsInWarehouse, 1, 0)) AS InWarehouse, Sum(IIf(e.IsInWarehouse, 0, 1)) AS OutOfWarehouse FROM (tblStage_Status AS s INNER JOIN tblOrderDetails AS d ON s.Waybill = d.Waybill) INNER JOIN tblEventTypes AS e ON s.LastEventKey = e.EventKey"
    SaveQuery "qryRpt_LastStatusChanges", "SELECT h.* FROM qryShipmentHistory AS h INNER JOIN tblStatusHistory AS sh ON h.HistoryID = sh.HistoryID WHERE sh.BatchID = (SELECT Max(BatchID) FROM tblImportBatches WHERE ImportType = 'Status')"
    SaveQuery "qryRpt_UnmatchedAll", "SELECT u.UnmatchedID, u.BatchID, b.ImportedOn, u.Waybill, u.LastEvent, u.CityName, u.ConsigneeName FROM tblUnmatched AS u INNER JOIN tblImportBatches AS b ON u.BatchID = b.BatchID ORDER BY u.BatchID DESC, u.Waybill"
    SaveQuery "qryRpt_UnmatchedLastBatch", "SELECT Waybill, LastEvent, CityName, ConsigneeName FROM tblUnmatched WHERE BatchID = (SELECT Max(BatchID) FROM tblImportBatches WHERE ImportType = 'Status') ORDER BY Waybill"
    SaveQuery "qryRpt_NeverUpdated", "SELECT Waybill, ReferenceNo, ProductDescription, LastUpdated FROM qryShipmentsList WHERE EventName Is Null"
End Sub

Private Sub SaveQuery(ByVal queryName As String, ByVal sql As String)
    Dim db As DAO.Database
    Set db = CurrentDb
    On Error Resume Next
    db.QueryDefs.Delete queryName
    On Error GoTo 0
    db.CreateQueryDef queryName, sql
End Sub

' ===========================================================================
' 4. FORMS
' ===========================================================================

Private Sub CreateForms()
    Dim i As Long
    For i = Forms.Count - 1 To 0 Step -1
        DoCmd.Close acForm, Forms(i).Name, acSaveNo
    Next i

    BuildMainMenu
    BuildImportForm
    BuildShipmentsForm
    BuildHistoryForm
    BuildEventTypesForm
    BuildListForm FRM_UNMATCHED, "Unmatched waybills (in status file, not in database)", "qryRpt_UnmatchedAll", _
        Array("BatchID", "ImportedOn", "Waybill", "LastEvent", "CityName", "ConsigneeName"), _
        Array("Batch", "Imported", "Waybill", "Last event", "City", HeaderAliases(COL_CONSIGNEE)(0)), _
        Array(1.5, 3.2, 4.5, 4, 3, 7.5), "=Exp_Unmatched()", "Export last batch"
    BuildListForm FRM_BATCHES, "Import history", "qryImportBatchesList", _
        Array("BatchID", "ImportType", "ImportedOn", "RowsRead", "RowsInserted", "RowsUpdated", "RowsUnmatched", "RowsSkipped", "FileInCount", "FileOutCount", "FileName"), _
        Array("Batch", "Type", "Imported", "Read", "New", "Changed", "Unmatched", "Skipped", "File IN", "File OUT", "File"), _
        Array(1.4, 1.8, 3.2, 1.4, 1.4, 1.6, 1.8, 1.6, 1.6, 1.6, 7)
    BuildListForm FRM_GRADES, "Quantity and loss by consignee and grade (G1-G13)", "qryRpt_ByConsigneeGrade", _
        Array("Consignee", "CustomerCode", "Grade", "Qty", "Loss", "InQty", "OutQty"), _
        Array(HeaderAliases(COL_CONSIGNEE)(0), "Code", "Grade", "Qty", "Loss", "Qty IN", "Qty OUT"), _
        Array(9, 2.4, 2, 1.8, 1.8, 1.8, 1.8), "=Exp_ConsigneeGrades()", "Export to Excel"
End Sub

' ---- frmMainMenu (dashboard) -------------------------------------------
Private Sub BuildMainMenu()
    Dim frm As Form, y As Long, i As Long, ctl As Control
    Set frm = NewForm("", 0)
    frm.Caption = APP_TITLE & " - Dashboard"
    frm.Width = 26 * TW
    frm.OnActivate = "=Dash_Refresh()"
    frm.Section(acDetail).Height = 15 * TW
    frm.Section(acDetail).BackColor = CLR_WHITE

    TitleBar frm, acDetail, APP_TITLE, 26

    ' KPI tiles: caption, control source, colour
    Dim kpis As Variant
    kpis = Array( _
        Array("IN warehouse", "=Kpi_InWarehouse()", CLR_IN), _
        Array("OUT of warehouse", "=Kpi_OutOfWarehouse()", CLR_OUT), _
        Array("Total waybills", "=Kpi_Total()", CLR_HEADER), _
        Array("% in warehouse", "=Kpi_InPercent()", CLR_HEADER))
    For i = 0 To UBound(kpis)
        KpiTile frm, "kpi" & i, kpis(i)(0), kpis(i)(1), kpis(i)(2), (0.5 + i * 6.3) * TW, 1.6 * TW
    Next i

    AddLabel frm, acDetail, "lblLastImportCap", "Last import:", 0.5 * TW, 4.4 * TW, 2.5 * TW
    AddTextBox frm, acDetail, "txtLastImport", "=Kpi_LastImport()", 3 * TW, 4.4 * TW, 6 * TW, , True
    AddLabel frm, acDetail, "lblLastChangedCap", "Last status file - changed:", 9.5 * TW, 4.4 * TW, 4.5 * TW
    AddTextBox frm, acDetail, "txtLastChanged", "=Kpi_LastStatusChanged()", 14 * TW, 4.4 * TW, 1.8 * TW, , True
    AddLabel frm, acDetail, "lblLastUnmatchedCap", "not found:", 16.2 * TW, 4.4 * TW, 2 * TW
    AddTextBox frm, acDetail, "txtLastUnmatched", "=Kpi_LastStatusUnmatched()", 18.2 * TW, 4.4 * TW, 1.8 * TW, , True

    ' Breakdown lists
    AddLabel frm, acDetail, "lblByEvent", "Quantity by last event", 0.5 * TW, 5.4 * TW, 8 * TW, True
    Set ctl = CreateControl(frm.Name, acListBox, acDetail, , , 0.5 * TW, 6 * TW, 12 * TW, 6 * TW)
    ctl.Name = "lstByEvent"
    ctl.RowSource = "qryRpt_ByEvent"
    ctl.ColumnCount = 3
    ctl.ColumnHeads = True
    ctl.ColumnWidths = CStr(7 * TW) & ";" & CStr(2 * TW) & ";" & CStr(2 * TW)

    AddLabel frm, acDetail, "lblByGrade", "Quantity and loss by grade", 13 * TW, 5.4 * TW, 8 * TW, True
    Set ctl = CreateControl(frm.Name, acListBox, acDetail, , , 13 * TW, 6 * TW, 12.5 * TW, 6 * TW)
    ctl.Name = "lstByGrade"
    ctl.RowSource = "qryRpt_ByGrade"
    ctl.ColumnCount = 5
    ctl.ColumnHeads = True
    ctl.ColumnWidths = CStr(3 * TW) & ";" & CStr(2.2 * TW) & ";" & CStr(2.2 * TW) & ";" & CStr(2.2 * TW) & ";" & CStr(2.2 * TW)

    ' Menu buttons (two rows)
    Dim btns As Variant
    btns = Array( _
        Array("btnImport", "Upload Excel", "=Nav_OpenImport()"), _
        Array("btnShipments", "All shipments", "=Nav_OpenShipments()"), _
        Array("btnIn", "IN warehouse", "=Nav_OpenInWarehouse()"), _
        Array("btnOut", "OUT of warehouse", "=Nav_OpenOutOfWarehouse()"), _
        Array("btnEvents", "Event rules", "=Nav_OpenEventTypes()"), _
        Array("btnUnmatched", "Unmatched", "=Nav_OpenUnmatched()"), _
        Array("btnGrades", "Consignee x grade", "=Nav_OpenConsigneeGrades()"), _
        Array("btnHistory", "Import history", "=Nav_OpenImportHistory()"), _
        Array("btnExpIn", "Export IN list", "=Exp_InWarehouse()"), _
        Array("btnExpOut", "Export OUT list", "=Exp_OutOfWarehouse()"), _
        Array("btnRefresh", "Refresh", "=Dash_Refresh()"), _
        Array("btnExit", "Exit", "=Nav_Exit()"))
    For i = 0 To UBound(btns)
        AddButton frm, acDetail, btns(i)(0), btns(i)(1), btns(i)(2), _
            (0.5 + (i Mod 6) * 4.2) * TW, IIf(i < 6, 12.5, 13.6) * TW, 4 * TW
    Next i

    DialogStyle frm
    SaveFormAs frm, FRM_MAIN
End Sub

' ---- frmImport ----------------------------------------------------------
Private Sub BuildImportForm()
    Dim frm As Form, ctl As Control
    Set frm = NewForm("", 0)
    frm.Caption = "Upload Excel file"
    frm.Width = 20 * TW
    frm.Section(acDetail).Height = 13 * TW

    TitleBar frm, acDetail, "Upload Excel file", 20

    AddLabel frm, acDetail, "lblFile", "Excel file:", 0.5 * TW, 1.6 * TW, 3 * TW
    AddTextBox frm, acDetail, "txtFilePath", "", 3.5 * TW, 1.6 * TW, 12.5 * TW
    AddButton frm, acDetail, "btnBrowse", "Browse...", "=Imp_Browse()", 16.3 * TW, 1.55 * TW, 3.2 * TW

    AddLabel frm, acDetail, "lblType", "File type:", 0.5 * TW, 2.6 * TW, 3 * TW
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , , 3.5 * TW, 2.6 * TW, 5 * TW, 0.6 * TW)
    ctl.Name = "cboImportType"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = IMPORT_AUTO & ";" & IMPORT_MASTER & ";" & IMPORT_STATUS
    ctl.DefaultValue = """" & IMPORT_AUTO & """"
    ctl.LimitToList = True
    AddLabel frm, acDetail, "lblTypeHelp", "Auto = detect from headers (status file has the 'last event' column).", 8.8 * TW, 2.6 * TW, 10.7 * TW

    AddButton frm, acDetail, "btnRun", "Import", "=Imp_Run()", 3.5 * TW, 3.6 * TW, 4 * TW
    AddButton frm, acDetail, "btnClose", "Close", "=Nav_CloseActive()", 15.5 * TW, 3.6 * TW, 4 * TW

    AddLabel frm, acDetail, "lblResult", "Result", 0.5 * TW, 4.6 * TW, 3 * TW, True
    AddTextBox frm, acDetail, "txtResult", "", 0.5 * TW, 5.2 * TW, 19 * TW, 7.3 * TW, True
    frm.Controls("txtResult").ScrollBars = 2
    frm.Controls("txtResult").FontName = "Consolas"

    DialogStyle frm
    SaveFormAs frm, FRM_IMPORT
End Sub

' ---- frmShipments (continuous, read-only, with search header) -----------
Private Sub BuildShipmentsForm()
    Dim frm As Form, ctl As Control
    Dim flds As Variant, caps As Variant, wid As Variant

    flds = Array("Waybill", "ReferenceNo", "SecondaryRef", "ProductDescription", "GradeCode", "LossQty", "EventName", "Location", "CityName", "CustomerName", "LastUpdated")
    caps = Array("Waybill", "Reference", "Secondary ref", "Contents", "Grade", "Loss", "Last event", "Loc.", "City", HeaderAliases(COL_CONSIGNEE)(0), "Updated")
    wid = Array(4, 2.4, 3.2, 5.5, 1.5, 1.2, 3.6, 1.3, 2.4, 6.5, 3)

    Set frm = BuildContinuous("qryShipmentsList", "Shipments", flds, caps, wid, 2.6)

    ' Search bar in the header (above the column captions)
    AddLabel frm, acHeader, "lblSearch", "Search:", 0.2 * TW, 0.2 * TW, 1.8 * TW
    AddTextBox frm, acHeader, "txtSearch", "", 2 * TW, 0.2 * TW, 6 * TW
    frm.Controls("txtSearch").AfterUpdate = "=Shp_ApplyFilter()"
    AddLabel frm, acHeader, "lblLoc", "Show:", 8.4 * TW, 0.2 * TW, 1.4 * TW
    Set ctl = CreateControl(frm.Name, acComboBox, acHeader, , , 9.8 * TW, 0.2 * TW, 4 * TW, 0.6 * TW)
    ctl.Name = "cboLocation"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = LOC_ALL & ";" & LOC_IN & ";" & LOC_OUT
    ctl.DefaultValue = """" & LOC_ALL & """"
    ctl.LimitToList = True
    ctl.AfterUpdate = "=Shp_ApplyFilter()"
    AddButton frm, acHeader, "btnSearch", "Search", "=Shp_ApplyFilter()", 14.2 * TW, 0.15 * TW, 2.5 * TW
    AddButton frm, acHeader, "btnClear", "Clear", "=Shp_ClearFilter()", 16.9 * TW, 0.15 * TW, 2.5 * TW
    AddButton frm, acHeader, "btnExpIn", "Export IN", "=Exp_InWarehouse()", 19.6 * TW, 0.15 * TW, 2.8 * TW
    AddButton frm, acHeader, "btnExpOut", "Export OUT", "=Exp_OutOfWarehouse()", 22.6 * TW, 0.15 * TW, 2.8 * TW
    AddButton frm, acHeader, "btnClose", "Close", "=Nav_CloseActive()", 25.6 * TW, 0.15 * TW, 2.5 * TW
    AddLabel frm, acHeader, "lblHint", "Double-click a row to see its status history. Green = IN, red = OUT.", 0.2 * TW, 1.1 * TW, 14 * TW

    ' Colour the location column: green IN / red OUT.
    Dim fc As Object
    Set fc = frm.Controls("txtLocation").FormatConditions.Add(acFieldValue, acEqual, """IN""")
    fc.ForeColor = CLR_IN: fc.FontBold = True
    Set fc = frm.Controls("txtLocation").FormatConditions.Add(acFieldValue, acEqual, """OUT""")
    fc.ForeColor = CLR_OUT: fc.FontBold = True

    Dim c As Variant
    For Each c In flds
        frm.Controls("txt" & c).OnDblClick = "=Nav_OpenHistoryForCurrent()"
    Next c

    frm.AllowEdits = False
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    SaveFormAs frm, FRM_SHIPMENTS
End Sub

' ---- frmShipmentHistory -----------------------------------------------
Private Sub BuildHistoryForm()
    BuildListForm FRM_HISTORY, "Status history", "qryShipmentHistory", _
        Array("Waybill", "FromEvent", "FromLocation", "ToEvent", "ToLocation", "RecordedOn", "FileName"), _
        Array("Waybill", "From event", "From", "To event", "To", "Recorded", "Source file"), _
        Array(4, 4, 1.4, 4, 1.4, 3.2, 8)
End Sub

' ---- frmEventTypes (event list; the IN/OUT flag is fixed and read-only) --
Private Sub BuildEventTypesForm()
    Dim frm As Form
    Set frm = BuildContinuous("qryEventTypesList", "Event rules - only 2 events mean IN warehouse", _
        Array("EventName", "IsInWarehouse", "SortOrder", "Notes"), _
        Array("Last event (as in status file)", "IN warehouse?", "Sort", "Notes"), _
        Array(8, 2.6, 1.6, 8), 1.6)

    AddLabel frm, acHeader, "lblHelp", "IN warehouse = ready for pickup or received from customer. Every other event is OUT (fixed rule).", 0.2 * TW, 0.2 * TW, 18 * TW
    AddButton frm, acHeader, "btnClose", "Close", "=Nav_CloseActive()", 18.5 * TW, 0.15 * TW, 2.5 * TW

    With frm.Controls("txtEventName")
        .Locked = False
        .TabStop = True
        .BackColor = CLR_WHITE
    End With
    With frm.Controls("txtSortOrder")
        .Locked = False
        .TabStop = True
        .BackColor = CLR_WHITE
    End With
    With frm.Controls("txtNotes")
        .Locked = False
        .TabStop = True
        .BackColor = CLR_WHITE
    End With
    frm.AllowEdits = True
    frm.AllowAdditions = True
    frm.AllowDeletions = False
    frm.Controls("txtIsInWarehouse").Locked = True    ' rule is fixed in code

    ' BeforeUpdate needs Cancel, so this one form gets a small class module.
    frm.BeforeUpdate = "[Event Procedure]"
    frm.HasModule = True
    frm.Module.AddFromString _
        "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf & _
        "    Dim nm As String, k As String" & vbCrLf & _
        "    nm = NormText(Me!EventName)" & vbCrLf & _
        "    If Len(nm) = 0 Then" & vbCrLf & _
        "        ShowWarning ""Event name is required.""" & vbCrLf & _
        "        Cancel = True: Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    k = NormKey(nm)" & vbCrLf & _
        "    If DCount(""*"", ""tblEventTypes"", ""EventKey="" & SqlText(k) & "" AND EventID<>"" & Nz(Me!EventID, 0)) > 0 Then" & vbCrLf & _
        "        ShowWarning ""This event already exists.""" & vbCrLf & _
        "        Cancel = True: Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    If Not IsNull(Me!SortOrder) And Not IsNumeric(Me!SortOrder) Then" & vbCrLf & _
        "        ShowWarning ""Sort must be a number.""" & vbCrLf & _
        "        Cancel = True: Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    Me!EventName = nm" & vbCrLf & _
        "    Me!EventKey = k" & vbCrLf & _
        "End Sub" & vbCrLf

    SaveFormAs frm, FRM_EVENTS
End Sub

' ---- generic read-only list form ---------------------------------------
Private Sub BuildListForm(ByVal formName As String, ByVal caption As String, ByVal recordSource As String, _
                          ByVal flds As Variant, ByVal caps As Variant, ByVal wid As Variant, _
                          Optional ByVal extraAction As String = "", Optional ByVal extraCaption As String = "")
    Dim frm As Form
    Set frm = BuildContinuous(recordSource, caption, flds, caps, wid, 1.2)
    AddButton frm, acHeader, "btnClose", "Close", "=Nav_CloseActive()", 0.2 * TW, 0.15 * TW, 2.5 * TW
    If Len(extraAction) > 0 Then
        AddButton frm, acHeader, "btnExtra", extraCaption, extraAction, 3 * TW, 0.15 * TW, 4 * TW
    End If
    frm.AllowEdits = False
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    SaveFormAs frm, formName
End Sub

' Continuous form: header holds column captions at headerTopCm, detail
' holds one text box (or check box for Yes/No fields) per field, footer
' shows a record count.
Private Function BuildContinuous(ByVal recordSource As String, ByVal caption As String, _
                                 ByVal flds As Variant, ByVal caps As Variant, ByVal wid As Variant, _
                                 ByVal headerTopCm As Double) As Form
    Dim frm As Form, i As Long, x As Long, ctl As Control, w As Long

    Set frm = NewForm(recordSource, 1)   ' 1 = continuous
    frm.Caption = caption

    frm.Section(acHeader).Height = (headerTopCm + 0.7) * TW
    frm.Section(acHeader).BackColor = CLR_PANEL
    frm.Section(acDetail).Height = 0.6 * TW
    frm.Section(acFooter).Height = 0.7 * TW

    x = 0.2 * TW
    For i = 0 To UBound(flds)
        w = wid(i) * TW
        With AddLabel(frm, acHeader, "lbl" & flds(i), caps(i), x, headerTopCm * TW, w, True)
            .TextAlign = 2
        End With
        If IsYesNoField(recordSource, flds(i)) Then
            Set ctl = CreateControl(frm.Name, acCheckBox, acDetail, , flds(i), x + w / 2 - 120, 0.1 * TW)
            ctl.Name = "txt" & flds(i)
        Else
            Set ctl = AddTextBox(frm, acDetail, "txt" & flds(i), flds(i), x, 0, w, 0.6 * TW, True)
        End If
        x = x + w + 0.05 * TW
    Next i
    If frm.Width < x + 0.2 * TW Then frm.Width = x + 0.2 * TW

    AddLabel frm, acFooter, "lblCount", "Records:", 0.2 * TW, 0.1 * TW, 2 * TW
    AddTextBox frm, acFooter, "txtCount", "=Count(*)", 2.2 * TW, 0.1 * TW, 2 * TW, , True

    Set BuildContinuous = frm
End Function

Private Function IsYesNoField(ByVal recordSource As String, ByVal fieldName As String) As Boolean
    On Error Resume Next
    IsYesNoField = (CurrentDb.QueryDefs(recordSource).Fields(fieldName).Type = dbBoolean)
End Function

' ---------------------------------------------------------------------------
' Form-building primitives
' ---------------------------------------------------------------------------

Private Function NewForm(ByVal recordSource As String, ByVal defaultView As Integer) As Form
    Dim frm As Form
    Set frm = CreateForm()
    If Len(recordSource) > 0 Then frm.RecordSource = recordSource
    frm.DefaultView = defaultView
    frm.ScrollBars = IIf(defaultView = 1, 2, 0)
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = (defaultView = 1)
    frm.AutoCenter = True
    frm.AllowLayoutView = False
    If defaultView = 1 Then
        DoCmd.RunCommand acCmdFormHdrFtr     ' add header + footer sections
    End If
    Set NewForm = frm
End Function

Private Sub DialogStyle(ByVal frm As Form)
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.ScrollBars = 0
    frm.DividingLines = False
    frm.AllowAdditions = False
    frm.AllowDeletions = False
End Sub

Private Sub SaveFormAs(ByVal frm As Form, ByVal finalName As String)
    Dim tmp As String
    tmp = frm.Name
    DoCmd.Close acForm, tmp, acSaveYes
    On Error Resume Next
    DoCmd.DeleteObject acForm, finalName
    On Error GoTo 0
    DoCmd.Rename finalName, acForm, tmp
    Debug.Print "form built: "; finalName
End Sub

Private Sub TitleBar(ByVal frm As Form, ByVal section As Long, ByVal text As String, ByVal widthCm As Double)
    With AddLabel(frm, section, "lblTitle", text, 0, 0, widthCm * TW)
        .Height = 1.2 * TW
        .BackStyle = 1
        .BackColor = CLR_HEADER
        .ForeColor = CLR_WHITE
        .FontSize = 16
        .FontBold = True
        .TextAlign = 2
    End With
End Sub

Private Sub KpiTile(ByVal frm As Form, ByVal baseName As String, ByVal caption As String, _
                    ByVal source As String, ByVal colour As Long, ByVal x As Long, ByVal y As Long)
    Dim box As Control
    Set box = CreateControl(frm.Name, acRectangle, acDetail, , , x, y, 6 * TW, 2.5 * TW)
    box.Name = "rect" & baseName
    box.BackStyle = 1
    box.BackColor = CLR_PANEL
    box.BorderColor = colour
    box.BorderWidth = 2
    With AddLabel(frm, acDetail, "lbl" & baseName, caption, x + 0.2 * TW, y + 0.15 * TW, 5.6 * TW)
        .TextAlign = 2
        .ForeColor = colour
    End With
    With AddTextBox(frm, acDetail, "txt" & baseName, source, x + 0.2 * TW, y + 0.8 * TW, 5.6 * TW, 1.4 * TW, True)
        .FontSize = 24
        .FontBold = True
        .TextAlign = 2
        .ForeColor = colour
        .BackStyle = 0
        .BorderStyle = 0
    End With
End Sub

Private Function AddLabel(ByVal frm As Form, ByVal section As Long, ByVal ctlName As String, _
                          ByVal caption As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, _
                          Optional ByVal bold As Boolean = False) As Control
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acLabel, section, , , x, y, w, 0.6 * TW)
    ctl.Name = ctlName
    ctl.Caption = caption
    ctl.FontBold = bold
    Set AddLabel = ctl
End Function

Private Function AddTextBox(ByVal frm As Form, ByVal section As Long, ByVal ctlName As String, _
                            ByVal source As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, _
                            Optional ByVal h As Long = 0, Optional ByVal readOnly As Boolean = False) As Control
    Dim ctl As Control
    If h = 0 Then h = 0.6 * TW
    Set ctl = CreateControl(frm.Name, acTextBox, section, , , x, y, w, h)
    ctl.Name = ctlName
    If Len(source) > 0 Then ctl.ControlSource = source
    If readOnly Then
        ctl.Locked = True
        ctl.TabStop = False
        ctl.BackColor = CLR_PANEL
    End If
    Set AddTextBox = ctl
End Function

Private Function AddButton(ByVal frm As Form, ByVal section As Long, ByVal ctlName As String, _
                           ByVal caption As String, ByVal onClick As String, _
                           ByVal x As Long, ByVal y As Long, ByVal w As Long) As Control
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acCommandButton, section, , , x, y, w, 0.8 * TW)
    ctl.Name = ctlName
    ctl.Caption = caption
    ctl.OnClick = onClick
    Set AddButton = ctl
End Function

' ===========================================================================
' 5. STARTUP
' ===========================================================================

Private Sub SetStartupOptions()
    SetDbProperty "StartUpForm", dbText, FRM_MAIN
    SetDbProperty "AppTitle", dbText, APP_TITLE
    SetDbProperty "StartUpShowDBWindow", dbBoolean, True   ' keep nav pane for admins
    Application.RefreshTitleBar
End Sub

Private Sub SetDbProperty(ByVal propName As String, ByVal propType As Integer, ByVal value As Variant)
    Dim db As DAO.Database
    Set db = CurrentDb
    On Error Resume Next
    db.Properties(propName) = value
    If Err.Number = 3270 Then          ' property not found
        Err.Clear
        db.Properties.Append db.CreateProperty(propName, propType, value)
    End If
End Sub
