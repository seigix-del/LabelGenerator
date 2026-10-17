// Every usable row of every status file (waybill + last event), with the
// IN/OUT rule applied. Used for the latest status and for history.
let
    Source = fnExcelFiles(StatusFolder),
    Clean = Table.TransformColumns(Source, {
        {"Waybill", fnNormWaybill}, {"LastEvent", fnCleanText},
        {"CityName", fnCleanText}, {"ConsigneeName", fnCleanText}}),
    Usable = Table.SelectRows(Clean, each [Waybill] <> null and [LastEvent] <> null),
    InKeys = List.Buffer(List.Transform(cfgInEvents, fnNormKey)),
    WithKey = Table.AddColumn(Usable, "EventKey", each fnNormKey([LastEvent]), type text),
    WithIn = Table.AddColumn(WithKey, "IsInEvent", each List.Contains(InKeys, [EventKey]), type logical),
    WithCode = Table.AddColumn(WithIn, "CustomerCode", each fnCustomerCode([ConsigneeName]), type text)
in
    Table.SelectColumns(WithCode,
        {"Waybill", "LastEvent", "EventKey", "IsInEvent", "CityName", "ConsigneeName", "CustomerCode", "SourceFile", "FileDate", "RowIndex"})
