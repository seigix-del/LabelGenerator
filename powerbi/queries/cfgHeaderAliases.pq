// Excel header text -> logical column. Matching uses fnNormKey, so extra
// spaces and alef / yeh / taa-marbuta variants do not matter.
let
    Aliases = {
        {"Waybill",       {"بوليصة الشحن", "رقم التتبع", "رقم البوليصة", "Waybill", "AWB", "Tracking Number", "Tracking No"}},
        {"ReferenceNo",   {"رقم المرجع", "Reference", "Reference No", "Ref"}},
        {"Contents",      {"وصف محتويات الشحنة", "Contents", "Description", "Shipment Contents"}},
        {"SecondaryRef",  {"رقم المرجع الثانوي", "Secondary Reference", "Secondary Ref"}},
        {"CityName",      {"مدينة الوجهة", "Destination City", "City"}},
        {"LastEvent",     {"آخر حدث", "Last Event", "Status"}},
        {"ConsigneeName", {"اسم المرسل إليه", "Consignee", "Consignee Name", "Receiver"}}
    },
    Rows = List.Combine(List.Transform(Aliases, (a) => List.Transform(a{1}, (h) => [Key = fnNormKey(h), Logical = a{0}])))
in
    Table.Buffer(Table.FromRecords(Rows))
