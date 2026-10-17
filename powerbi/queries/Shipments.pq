// Fact table: one row per waybill in the inventory list, with its latest
// status. No status yet = IN warehouse (it is on the inventory list).
let
    Master = Table.RenameColumns(MasterClean, {{"SourceFile", "MasterFile"}, {"FileDate", "MasterFileDate"}}),
    Joined = Table.NestedJoin(Master, {"Waybill"}, StatusLatest, {"Waybill"}, "St", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "St",
        {"LastEvent", "IsInEvent", "CityName", "ConsigneeName", "CustomerCode", "SourceFile", "FileDate"},
        {"LastEvent", "IsInEvent", "City", "ConsigneeRaw", "CustomerCode", "StatusFile", "StatusFileDate"}),
    AddHas = Table.AddColumn(Expanded, "HasStatus", each [LastEvent] <> null, type logical),
    AddIn = Table.AddColumn(AddHas, "InWarehouse", each if [LastEvent] = null then true else [IsInEvent], type logical),
    AddLoc = Table.AddColumn(AddIn, "Location", each if [InWarehouse] then "IN" else "OUT", type text),
    AddEvent = Table.AddColumn(AddLoc, "LastEventLabel", each if [LastEvent] = null then "(no status yet)" else [LastEvent], type text),
    AddConsignee = Table.AddColumn(AddEvent, "Consignee", each if [ConsigneeRaw] = null then "(no consignee yet)" else [ConsigneeRaw], type text),
    // G row = 1 item of its grade; G<n>-L<k> row = k lost items (L side).
    AddLossQty = Table.AddColumn(AddConsignee, "LossQty", each if [Loss] = null then 0 else [Loss], Int64.Type),
    AddRowType = Table.AddColumn(AddLossQty, "RowType", each if [LossQty] > 0 then "L" else "G", type text),
    AddGQty = Table.AddColumn(AddRowType, "GQty", each if [LossQty] > 0 then 0 else 1, Int64.Type),
    Selected = Table.SelectColumns(AddGQty, {
        "Waybill", "Reference", "SecondaryRef", "Contents", "Grade", "RowType", "GQty", "LossQty",
        "LastEventLabel", "HasStatus", "InWarehouse", "Location", "City", "Consignee", "CustomerCode",
        "MasterFile", "MasterFileDate", "StatusFile", "StatusFileDate"})
in
    Table.TransformColumnTypes(Selected, {
        {"Waybill", type text}, {"Reference", type text}, {"SecondaryRef", type text}, {"Contents", type text},
        {"Grade", type text}, {"City", type text}, {"CustomerCode", type text},
        {"MasterFile", type text}, {"MasterFileDate", type datetime},
        {"StatusFile", type text}, {"StatusFileDate", type datetime}})
