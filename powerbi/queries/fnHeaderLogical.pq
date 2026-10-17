// Logical column name for an Excel header cell, or null.
(header as any) as nullable text =>
let
    Key = fnNormKey(header),
    Hits = if Key = null then {} else Table.SelectRows(cfgHeaderAliases, each [Key] = Key)[Logical]
in
    List.First(Hits, null)
