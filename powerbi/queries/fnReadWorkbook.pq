// Reads one workbook: takes the first sheet whose first 15 rows contain a
// waybill header, promotes that row to headers and returns the 7 logical
// columns (missing ones as null) plus RowIndex.
(content as binary) as table =>
let
    LogicalCols = {"Waybill", "ReferenceNo", "Contents", "SecondaryRef", "CityName", "LastEvent", "ConsigneeName"},
    Empty = #table(LogicalCols & {"RowIndex"}, {}),
    Prepare = (raw as table) as nullable table =>
        let
            Top = Table.ToRows(Table.FirstN(raw, 15)),
            Flags = List.Transform(Top, (row) => List.AnyTrue(List.Transform(row, (c) => fnHeaderLogical(c) = "Waybill"))),
            Idx = List.PositionOf(Flags, true)
        in
            if Idx < 0 then null
            else
                let
                    Body = Table.PromoteHeaders(Table.Skip(raw, Idx), [PromoteAllScalars = true]),
                    Pairs = List.Transform(Table.ColumnNames(Body), (n) => {n, fnHeaderLogical(n)}),
                    Mapped = List.Select(Pairs, each _{1} <> null),
                    FirstPerLogical = List.Accumulate(Mapped, {},
                        (acc, p) => if List.Contains(List.Transform(acc, each _{1}), p{1}) then acc else acc & {p}),
                    Kept = Table.SelectColumns(Body, List.Transform(FirstPerLogical, each _{0})),
                    Renamed = Table.RenameColumns(Kept, List.Select(FirstPerLogical, each _{0} <> _{1})),
                    Complete = List.Accumulate(LogicalCols, Renamed,
                        (t, c) => if Table.HasColumns(t, c) then t else Table.AddColumn(t, c, each null)),
                    Indexed = Table.AddIndexColumn(Table.SelectColumns(Complete, LogicalCols), "RowIndex", 1, 1)
                in
                    Indexed,
    Sheets = Table.SelectRows(Excel.Workbook(content, false), each [Kind] = "Sheet"),
    Prepared = List.RemoveNulls(List.Transform(Sheets[Data], Prepare))
in
    List.First(Prepared, Empty)
