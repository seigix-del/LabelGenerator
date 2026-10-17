// Waybill / tracking number: upper case, no spaces.
(value as any) as nullable text =>
let
    Clean = fnCleanText(value)
in
    if Clean = null then null else Text.Upper(Text.Replace(Clean, " ", ""))
