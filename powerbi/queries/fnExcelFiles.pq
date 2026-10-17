// Every Excel file in a folder (and its subfolders) combined, with the
// source file name and its modified date. Excel lock files (~$...) are skipped.
(folder as text) as table =>
let
    Files = Folder.Files(folder),
    Excel = Table.SelectRows(Files, each List.Contains({".xlsx", ".xlsm", ".xls"}, Text.Lower([Extension]))
        and not Text.StartsWith([Name], "~$")),
    WithRows = Table.AddColumn(Excel, "Rows", each fnReadWorkbook([Content])),
    Slim = Table.SelectColumns(WithRows, {"Name", "Date modified", "Rows"}),
    Expanded = Table.ExpandTableColumn(Slim, "Rows",
        {"Waybill", "ReferenceNo", "Contents", "SecondaryRef", "CityName", "LastEvent", "ConsigneeName", "RowIndex"})
in
    Table.RenameColumns(Expanded, {{"Name", "SourceFile"}, {"Date modified", "FileDate"}})
