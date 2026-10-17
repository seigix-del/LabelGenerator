// Waybills in the latest status rows that are not in the inventory list.
let
    Joined = Table.NestedJoin(StatusLatest, {"Waybill"}, MasterClean, {"Waybill"}, "M", JoinKind.LeftAnti),
    Selected = Table.SelectColumns(Joined, {"Waybill", "LastEvent", "CityName", "ConsigneeName", "SourceFile", "FileDate"})
in
    Table.TransformColumnTypes(Selected, {
        {"Waybill", type text}, {"LastEvent", type text}, {"CityName", type text},
        {"ConsigneeName", type text}, {"SourceFile", type text}, {"FileDate", type datetime}})
