// Every status row ever uploaded (one per waybill per file) - the event trail.
let
    Selected = Table.SelectColumns(StatusAll, {"Waybill", "LastEvent", "IsInEvent", "SourceFile", "FileDate"}),
    WithLoc = Table.AddColumn(Selected, "Location", each if [IsInEvent] then "IN" else "OUT", type text),
    Distinct = Table.Distinct(WithLoc, {"Waybill", "SourceFile"})
in
    Table.TransformColumnTypes(Table.RemoveColumns(Distinct, {"IsInEvent"}), {
        {"Waybill", type text}, {"LastEvent", type text}, {"SourceFile", type text}, {"FileDate", type datetime}})
