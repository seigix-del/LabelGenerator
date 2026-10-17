// "<school name> - M174562" -> "M174562" (letters followed by digits at the end).
(value as nullable text) as nullable text =>
let
    Parts = if value = null then {} else Text.Split(value, "-"),
    Tail = if List.Count(Parts) < 2 then "" else Text.Trim(List.Last(Parts)),
    First = Text.Upper(Text.Start(Tail, 1)),
    Last = Text.End(Tail, 1),
    Ok = Text.Length(Tail) >= 2 and Text.Length(Tail) <= 30 and not Text.Contains(Tail, " ")
        and First >= "A" and First <= "Z" and Last >= "0" and Last <= "9"
in
    if Ok then Text.Upper(Tail) else null
