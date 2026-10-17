// Inventory list: one row per waybill. When a waybill appears in several
// files the newest file wins (then the last row in that file).
let
    Source = fnExcelFiles(MasterFolder),
    Clean = Table.TransformColumns(Source, {
        {"Waybill", fnNormWaybill}, {"ReferenceNo", fnCleanText},
        {"Contents", fnCleanText}, {"SecondaryRef", fnCleanText}}),
    WithWaybill = Table.SelectRows(Clean, each [Waybill] <> null),
    Sorted = Table.Buffer(Table.Sort(WithWaybill,
        {{"FileDate", Order.Descending}, {"SourceFile", Order.Descending}, {"RowIndex", Order.Descending}})),
    Latest = Table.Distinct(Sorted, {"Waybill"}),
    WithRef = Table.AddColumn(Latest, "Reference", each if [ReferenceNo] = null then [Waybill] else [ReferenceNo], type text),
    // Grade/loss from the contents; fall back to the secondary reference (e.g. ASR-SEP30-G9).
    WithGrade = Table.AddColumn(WithRef, "GL", each
        let a = fnParseGradeLoss([Contents]) in if a <> null then a else fnParseGradeLoss([SecondaryRef])),
    Expanded = Table.ExpandRecordColumn(WithGrade, "GL", {"Grade", "Loss"})
in
    Table.SelectColumns(Expanded, {"Waybill", "Reference", "SecondaryRef", "Contents", "Grade", "Loss", "SourceFile", "FileDate"})
