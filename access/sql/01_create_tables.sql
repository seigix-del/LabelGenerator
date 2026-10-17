-- Generated from access/vba/modBuildDatabase.bas by access/tools/extract_sql.py - do not edit.
-- Access DDL, ANSI-92 mode (run through ADO: CurrentProject.Connection.Execute).

CREATE TABLE tblCustomers (CustomerID COUNTER CONSTRAINT PK_Customers PRIMARY KEY, CustomerName TEXT(255) NOT NULL, CustomerCode TEXT(50), CreatedOn DATETIME DEFAULT Now());

CREATE UNIQUE INDEX UX_Customers_Name ON tblCustomers (CustomerName);

CREATE INDEX IX_Customers_Code ON tblCustomers (CustomerCode);

CREATE TABLE tblCities (CityID COUNTER CONSTRAINT PK_Cities PRIMARY KEY, CityName TEXT(100) NOT NULL);

CREATE UNIQUE INDEX UX_Cities_Name ON tblCities (CityName);

CREATE TABLE tblProducts (ProductID COUNTER CONSTRAINT PK_Products PRIMARY KEY, ProductDescription TEXT(255) NOT NULL, ProductCode TEXT(50));

CREATE UNIQUE INDEX UX_Products_Desc ON tblProducts (ProductDescription);

CREATE TABLE tblEventTypes (EventID COUNTER CONSTRAINT PK_EventTypes PRIMARY KEY, EventName TEXT(150) NOT NULL, EventKey TEXT(150) NOT NULL, IsInWarehouse YESNO DEFAULT 0, SortOrder LONG DEFAULT 100, Notes TEXT(255));

CREATE UNIQUE INDEX UX_EventTypes_Key ON tblEventTypes (EventKey);

CREATE TABLE tblImportBatches (BatchID COUNTER CONSTRAINT PK_ImportBatches PRIMARY KEY, ImportType TEXT(20) NOT NULL, FileName TEXT(255), ImportedOn DATETIME DEFAULT Now(), ImportedBy TEXT(100), RowsRead LONG DEFAULT 0, RowsInserted LONG DEFAULT 0, RowsUpdated LONG DEFAULT 0, RowsUnmatched LONG DEFAULT 0, RowsSkipped LONG DEFAULT 0, FileInCount LONG DEFAULT 0, FileOutCount LONG DEFAULT 0);

CREATE TABLE tblGrades (GradeCode TEXT(10) CONSTRAINT PK_Grades PRIMARY KEY, SortOrder LONG NOT NULL);

CREATE TABLE tblOrders (OrderID COUNTER CONSTRAINT PK_Orders PRIMARY KEY, ReferenceNo TEXT(50) NOT NULL, SecondaryRef TEXT(100), CustomerID LONG, CityID LONG, CreatedBatchID LONG, CreatedOn DATETIME DEFAULT Now(),
    CONSTRAINT FK_Orders_Customers FOREIGN KEY (CustomerID) REFERENCES tblCustomers (CustomerID),
    CONSTRAINT FK_Orders_Cities FOREIGN KEY (CityID) REFERENCES tblCities (CityID),
    CONSTRAINT FK_Orders_Batches FOREIGN KEY (CreatedBatchID) REFERENCES tblImportBatches (BatchID));

CREATE UNIQUE INDEX UX_Orders_Ref ON tblOrders (ReferenceNo);

CREATE TABLE tblOrderDetails (DetailID COUNTER CONSTRAINT PK_OrderDetails PRIMARY KEY, Waybill TEXT(50) NOT NULL, OrderID LONG NOT NULL, ProductID LONG, Quantity LONG DEFAULT 1, GradeCode TEXT(10), LossQty LONG DEFAULT 0, CurrentEventID LONG, InWarehouse YESNO DEFAULT -1, LastUpdated DATETIME DEFAULT Now(), LastBatchID LONG, CreatedOn DATETIME DEFAULT Now(),
    CONSTRAINT FK_Details_Orders FOREIGN KEY (OrderID) REFERENCES tblOrders (OrderID) ON DELETE CASCADE,
    CONSTRAINT FK_Details_Products FOREIGN KEY (ProductID) REFERENCES tblProducts (ProductID),
    CONSTRAINT FK_Details_Events FOREIGN KEY (CurrentEventID) REFERENCES tblEventTypes (EventID),
    CONSTRAINT FK_Details_Batches FOREIGN KEY (LastBatchID) REFERENCES tblImportBatches (BatchID));

CREATE UNIQUE INDEX UX_Details_Waybill ON tblOrderDetails (Waybill);

CREATE INDEX IX_Details_InWarehouse ON tblOrderDetails (InWarehouse);

CREATE TABLE tblStatusHistory (HistoryID COUNTER CONSTRAINT PK_StatusHistory PRIMARY KEY, DetailID LONG NOT NULL, EventID LONG NOT NULL, PrevEventID LONG, WasInWarehouse YESNO, IsInWarehouse YESNO, BatchID LONG, RecordedOn DATETIME DEFAULT Now(),
    CONSTRAINT FK_History_Details FOREIGN KEY (DetailID) REFERENCES tblOrderDetails (DetailID) ON DELETE CASCADE,
    CONSTRAINT FK_History_Events FOREIGN KEY (EventID) REFERENCES tblEventTypes (EventID),
    CONSTRAINT FK_History_PrevEvents FOREIGN KEY (PrevEventID) REFERENCES tblEventTypes (EventID),
    CONSTRAINT FK_History_Batches FOREIGN KEY (BatchID) REFERENCES tblImportBatches (BatchID));

CREATE TABLE tblUnmatched (UnmatchedID COUNTER CONSTRAINT PK_Unmatched PRIMARY KEY, BatchID LONG NOT NULL, Waybill TEXT(50), LastEvent TEXT(150), CityName TEXT(100), ConsigneeName TEXT(255), RecordedOn DATETIME DEFAULT Now(),
    CONSTRAINT FK_Unmatched_Batches FOREIGN KEY (BatchID) REFERENCES tblImportBatches (BatchID) ON DELETE CASCADE);

CREATE INDEX IX_Unmatched_Waybill ON tblUnmatched (Waybill);

CREATE TABLE tblStage_Master (StageID COUNTER CONSTRAINT PK_StageMaster PRIMARY KEY, BatchID LONG, Waybill TEXT(50) NOT NULL, ReferenceNo TEXT(50), Contents TEXT(255), ProductCode TEXT(50), SecondaryRef TEXT(100), GradeCode TEXT(10), LossQty LONG DEFAULT 0);

CREATE UNIQUE INDEX UX_StageMaster_Waybill ON tblStage_Master (Waybill);

CREATE TABLE tblStage_Status (StageID COUNTER CONSTRAINT PK_StageStatus PRIMARY KEY, BatchID LONG, Waybill TEXT(50) NOT NULL, LastEvent TEXT(150), LastEventKey TEXT(150), CityName TEXT(100), ConsigneeName TEXT(255), CustomerCode TEXT(50));

CREATE UNIQUE INDEX UX_StageStatus_Waybill ON tblStage_Status (Waybill);

CREATE TABLE tblErrorLog (ErrorID COUNTER CONSTRAINT PK_ErrorLog PRIMARY KEY, ErrNumber LONG, ErrDescription MEMO, ErrSource TEXT(255), UserName TEXT(100), LoggedOn DATETIME DEFAULT Now());
