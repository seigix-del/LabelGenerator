// Latest status per waybill: newest file wins, then the last row in it.
let
    Sorted = Table.Buffer(Table.Sort(StatusAll,
        {{"FileDate", Order.Descending}, {"SourceFile", Order.Descending}, {"RowIndex", Order.Descending}}))
in
    Table.Distinct(Sorted, {"Waybill"})
