// Grade dimension (G1-G9, G11-G13) so every grade shows, in numeric order.
let
    Rows = List.Transform(cfgGrades, each [Grade = _, SortOrder = Number.From(Text.Middle(_, 1))])
in
    Table.FromRecords(Rows, type table [Grade = text, SortOrder = Int64.Type])
