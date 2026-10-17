// Cleans a cell value: removes direction marks, tatweel, BOM; turns line
// breaks, tabs and no-break spaces into spaces; collapses runs of spaces.
// Returns null for empty values.
(value as any) as nullable text =>
let
    Raw = if value = null then null else Text.From(value),
    NoMarks = if Raw = null then null
        else Text.Remove(Raw, List.Transform({8206, 8207, 8234, 8235, 8236, 8237, 8238, 1600, 65279}, Character.FromNumber)),
    Spaced = if NoMarks = null then null
        else List.Accumulate({"#(cr)", "#(lf)", "#(tab)", Character.FromNumber(160)}, NoMarks, (acc, c) => Text.Replace(acc, c, " ")),
    Collapsed = if Spaced = null then null else Text.Combine(List.Select(Text.Split(Spaced, " "), each _ <> ""), " ")
in
    if Collapsed = null or Collapsed = "" then null else Collapsed
