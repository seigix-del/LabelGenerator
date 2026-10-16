-- Generated from access/vba/modBuildDatabase.bas by access/tools/extract_sql.py - do not edit.
-- Saved Access queries (DAO / ANSI-89). Name of each query is on the line above it.

-- qryImp_ClearStageMaster
DELETE FROM tblStage_Master;

-- qryImp_AppendProducts
INSERT INTO tblProducts (ProductDescription, ProductCode) SELECT s.Contents, First(s.ProductCode) AS FirstCode FROM tblStage_Master AS s LEFT JOIN tblProducts AS p ON s.Contents = p.ProductDescription WHERE p.ProductID Is Null AND s.Contents Is Not Null GROUP BY s.Contents;

-- qryImp_AppendOrders
INSERT INTO tblOrders (ReferenceNo, SecondaryRef, CreatedBatchID) SELECT s.ReferenceNo, First(s.SecondaryRef) AS FirstSecRef, First(s.BatchID) AS FirstBatch FROM tblStage_Master AS s LEFT JOIN tblOrders AS o ON s.ReferenceNo = o.ReferenceNo WHERE o.OrderID Is Null AND s.ReferenceNo Is Not Null GROUP BY s.ReferenceNo;

-- qryImp_UpdateOrdersSecondaryRef
UPDATE tblOrders AS o INNER JOIN tblStage_Master AS s ON o.ReferenceNo = s.ReferenceNo SET o.SecondaryRef = s.SecondaryRef WHERE s.SecondaryRef Is Not Null;

-- qryImp_AppendDetails
INSERT INTO tblOrderDetails (Waybill, OrderID, ProductID, Quantity, InWarehouse, LastBatchID) SELECT s.Waybill, o.OrderID, p.ProductID, 1, True, s.BatchID FROM ((tblStage_Master AS s INNER JOIN tblOrders AS o ON s.ReferenceNo = o.ReferenceNo) LEFT JOIN tblProducts AS p ON s.Contents = p.ProductDescription) LEFT JOIN tblOrderDetails AS d ON s.Waybill = d.Waybill WHERE d.DetailID Is Null;

-- qryImp_ClearStageStatus
DELETE FROM tblStage_Status;

-- qryImp_AppendEventTypes
INSERT INTO tblEventTypes (EventName, EventKey, IsInWarehouse, SortOrder, Notes) SELECT First(s.LastEvent) AS FirstName, s.LastEventKey, False, 100, 'Auto-added on import - review' FROM tblStage_Status AS s LEFT JOIN tblEventTypes AS e ON s.LastEventKey = e.EventKey WHERE e.EventID Is Null AND s.LastEventKey Is Not Null GROUP BY s.LastEventKey;

-- qryImp_AppendCities
INSERT INTO tblCities (CityName) SELECT s.CityName FROM tblStage_Status AS s LEFT JOIN tblCities AS c ON s.CityName = c.CityName WHERE c.CityID Is Null AND s.CityName Is Not Null GROUP BY s.CityName;

-- qryImp_AppendCustomers
INSERT INTO tblCustomers (CustomerName, CustomerCode) SELECT s.ConsigneeName, First(s.CustomerCode) AS FirstCode FROM tblStage_Status AS s LEFT JOIN tblCustomers AS c ON s.ConsigneeName = c.CustomerName WHERE c.CustomerID Is Null AND s.ConsigneeName Is Not Null GROUP BY s.ConsigneeName;

-- qryImp_AppendUnmatched
INSERT INTO tblUnmatched (BatchID, Waybill, LastEvent, CityName, ConsigneeName) SELECT s.BatchID, s.Waybill, s.LastEvent, s.CityName, s.ConsigneeName FROM tblStage_Status AS s LEFT JOIN tblOrderDetails AS d ON s.Waybill = d.Waybill WHERE d.DetailID Is Null;

-- qryImp_AppendHistory
INSERT INTO tblStatusHistory (DetailID, EventID, PrevEventID, WasInWarehouse, IsInWarehouse, BatchID) SELECT d.DetailID, e.EventID, d.CurrentEventID, d.InWarehouse, e.IsInWarehouse, s.BatchID FROM (tblStage_Status AS s INNER JOIN tblOrderDetails AS d ON s.Waybill = d.Waybill) INNER JOIN tblEventTypes AS e ON s.LastEventKey = e.EventKey WHERE d.CurrentEventID Is Null OR d.CurrentEventID <> e.EventID;

-- qryImp_UpdateDetailsStatus
UPDATE (tblOrderDetails AS d INNER JOIN tblStage_Status AS s ON d.Waybill = s.Waybill) INNER JOIN tblEventTypes AS e ON s.LastEventKey = e.EventKey SET d.CurrentEventID = e.EventID, d.InWarehouse = e.IsInWarehouse, d.LastUpdated = Now(), d.LastBatchID = s.BatchID WHERE d.CurrentEventID Is Null OR d.CurrentEventID <> e.EventID;

-- qryImp_UpdateOrdersCustomer
UPDATE ((tblOrders AS o INNER JOIN tblOrderDetails AS d ON o.OrderID = d.OrderID) INNER JOIN tblStage_Status AS s ON d.Waybill = s.Waybill) INNER JOIN tblCustomers AS c ON s.ConsigneeName = c.CustomerName SET o.CustomerID = c.CustomerID;

-- qryImp_UpdateOrdersCity
UPDATE ((tblOrders AS o INNER JOIN tblOrderDetails AS d ON o.OrderID = d.OrderID) INNER JOIN tblStage_Status AS s ON d.Waybill = s.Waybill) INNER JOIN tblCities AS c ON s.CityName = c.CityName SET o.CityID = c.CityID;

-- qryAct_RecalcInWarehouse
UPDATE tblOrderDetails AS d INNER JOIN tblEventTypes AS e ON d.CurrentEventID = e.EventID SET d.InWarehouse = e.IsInWarehouse WHERE d.InWarehouse <> e.IsInWarehouse;

-- qryAct_PurgeOldUnmatched
DELETE FROM tblUnmatched WHERE RecordedOn < DateAdd('d', -90, Date());

-- qryShipmentsList
SELECT d.DetailID, d.Waybill, o.ReferenceNo, o.SecondaryRef, p.ProductDescription, p.ProductCode, d.Quantity, e.EventName, d.InWarehouse, IIf(d.InWarehouse, 'IN', 'OUT') AS Location, ci.CityName, cu.CustomerName, cu.CustomerCode, d.LastUpdated FROM ((((tblOrderDetails AS d INNER JOIN tblOrders AS o ON d.OrderID = o.OrderID) LEFT JOIN tblProducts AS p ON d.ProductID = p.ProductID) LEFT JOIN tblEventTypes AS e ON d.CurrentEventID = e.EventID) LEFT JOIN tblCities AS ci ON o.CityID = ci.CityID) LEFT JOIN tblCustomers AS cu ON o.CustomerID = cu.CustomerID ORDER BY d.Waybill;

-- qryShipmentHistory
SELECT h.HistoryID, h.DetailID, d.Waybill, pe.EventName AS FromEvent, ne.EventName AS ToEvent, IIf(h.WasInWarehouse, 'IN', 'OUT') AS FromLocation, IIf(h.IsInWarehouse, 'IN', 'OUT') AS ToLocation, h.RecordedOn, b.FileName FROM (((tblStatusHistory AS h INNER JOIN tblOrderDetails AS d ON h.DetailID = d.DetailID) INNER JOIN tblEventTypes AS ne ON h.EventID = ne.EventID) LEFT JOIN tblEventTypes AS pe ON h.PrevEventID = pe.EventID) LEFT JOIN tblImportBatches AS b ON h.BatchID = b.BatchID ORDER BY h.RecordedOn DESC, h.HistoryID DESC;

-- qryEventTypesList
SELECT EventID, EventName, EventKey, IsInWarehouse, SortOrder, Notes FROM tblEventTypes ORDER BY SortOrder, EventName;

-- qryImportBatchesList
SELECT BatchID, ImportType, ImportedOn, ImportedBy, FileName, RowsRead, RowsInserted, RowsUpdated, RowsUnmatched, RowsSkipped, FileInCount, FileOutCount FROM tblImportBatches ORDER BY BatchID DESC;

-- qryRpt_InventorySummary
SELECT Sum(IIf(d.InWarehouse, d.Quantity, 0)) AS InWarehouse, Sum(IIf(d.InWarehouse, 0, d.Quantity)) AS OutOfWarehouse, Sum(d.Quantity) AS TotalQty, Count(d.DetailID) AS Waybills FROM tblOrderDetails AS d;

-- qryRpt_ByEvent
SELECT IIf(e.EventName Is Null, '(no status yet)', e.EventName) AS EventLabel, IIf(d.InWarehouse, 'IN', 'OUT') AS Location, Sum(d.Quantity) AS Qty FROM tblOrderDetails AS d LEFT JOIN tblEventTypes AS e ON d.CurrentEventID = e.EventID GROUP BY IIf(e.EventName Is Null, '(no status yet)', e.EventName), IIf(d.InWarehouse, 'IN', 'OUT'), e.SortOrder ORDER BY e.SortOrder, Sum(d.Quantity) DESC;

-- qryRpt_ByProduct
SELECT IIf(p.ProductDescription Is Null, '(unknown)', p.ProductDescription) AS Product, Sum(IIf(d.InWarehouse, d.Quantity, 0)) AS InQty, Sum(IIf(d.InWarehouse, 0, d.Quantity)) AS OutQty, Sum(d.Quantity) AS TotalQty FROM tblOrderDetails AS d LEFT JOIN tblProducts AS p ON d.ProductID = p.ProductID GROUP BY IIf(p.ProductDescription Is Null, '(unknown)', p.ProductDescription) ORDER BY Sum(IIf(d.InWarehouse, d.Quantity, 0)) DESC;

-- qryRpt_ByCity
SELECT IIf(c.CityName Is Null, '(unknown)', c.CityName) AS City, Sum(IIf(d.InWarehouse, d.Quantity, 0)) AS InQty, Sum(IIf(d.InWarehouse, 0, d.Quantity)) AS OutQty FROM (tblOrderDetails AS d INNER JOIN tblOrders AS o ON d.OrderID = o.OrderID) LEFT JOIN tblCities AS c ON o.CityID = c.CityID GROUP BY IIf(c.CityName Is Null, '(unknown)', c.CityName) ORDER BY Sum(IIf(d.InWarehouse, d.Quantity, 0)) DESC;

-- qryRpt_InWarehouse
SELECT Waybill, ReferenceNo, SecondaryRef, ProductDescription, EventName, CityName, CustomerName, LastUpdated FROM qryShipmentsList WHERE InWarehouse = True;

-- qryRpt_OutOfWarehouse
SELECT Waybill, ReferenceNo, SecondaryRef, ProductDescription, EventName, CityName, CustomerName, LastUpdated FROM qryShipmentsList WHERE InWarehouse = False;

-- qryRpt_LastStatusBatch
SELECT TOP 1 * FROM tblImportBatches WHERE ImportType = 'Status' ORDER BY BatchID DESC;

-- qryRpt_LastStatusFileSplit
SELECT Sum(IIf(e.IsInWarehouse, 1, 0)) AS InWarehouse, Sum(IIf(e.IsInWarehouse, 0, 1)) AS OutOfWarehouse FROM (tblStage_Status AS s INNER JOIN tblOrderDetails AS d ON s.Waybill = d.Waybill) INNER JOIN tblEventTypes AS e ON s.LastEventKey = e.EventKey;

-- qryRpt_LastStatusChanges
SELECT h.* FROM qryShipmentHistory AS h INNER JOIN tblStatusHistory AS sh ON h.HistoryID = sh.HistoryID WHERE sh.BatchID = (SELECT Max(BatchID) FROM tblImportBatches WHERE ImportType = 'Status');

-- qryRpt_UnmatchedAll
SELECT u.UnmatchedID, u.BatchID, b.ImportedOn, u.Waybill, u.LastEvent, u.CityName, u.ConsigneeName FROM tblUnmatched AS u INNER JOIN tblImportBatches AS b ON u.BatchID = b.BatchID ORDER BY u.BatchID DESC, u.Waybill;

-- qryRpt_UnmatchedLastBatch
SELECT Waybill, LastEvent, CityName, ConsigneeName FROM tblUnmatched WHERE BatchID = (SELECT Max(BatchID) FROM tblImportBatches WHERE ImportType = 'Status') ORDER BY Waybill;

-- qryRpt_NeverUpdated
SELECT Waybill, ReferenceNo, ProductDescription, LastUpdated FROM qryShipmentsList WHERE EventName Is Null;
