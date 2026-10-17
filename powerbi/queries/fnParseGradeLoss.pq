// Finds the grade code in a contents text such as
//   "نقطة توزيع المجاردة \\ G1"   -> [Grade = "G1", Loss = 0]
//   "... // G1-L4"  or  "... / G1-L4" -> [Grade = "G1", Loss = 4]
// G<n> = 1 item of that grade; G<n>-L<k> = a loss row of k items.
// The last code in the text wins. Returns null when there is no code or the
// grade is not in cfgGrades (e.g. G10).
(value as nullable text) as nullable record =>
let
    S = if value = null then "" else Text.Upper(value),
    Len = Text.Length(S),
    Chr = (i as number) as text => if i >= 0 and i < Len then Text.At(S, i) else "",
    IsDigit = (c as text) as logical => c >= "0" and c <= "9",
    IsAlnum = (c as text) as logical => (c >= "A" and c <= "Z") or IsDigit(c),
    CountWhile = (i as number, test as function) as number =>
        List.Count(List.FirstN(List.Numbers(i, List.Max({0, Len - i})), each test(Chr(_)))),
    MatchAt = (p as number) as nullable record =>
        let
            OkStart = p = 0 or not IsAlnum(Chr(p - 1)),
            DigitsG = CountWhile(p + 1, IsDigit),
            EndG = p + 1 + DigitsG,
            OkG = OkStart and DigitsG >= 1 and DigitsG <= 2 and not IsAlnum(Chr(EndG)),
            I1 = EndG + CountWhile(EndG, each _ = " "),
            I2 = I1 + 1 + CountWhile(I1 + 1, each _ = " "),
            HasL = Chr(I1) = "-" and Chr(I2) = "L",
            I3 = I2 + 1 + CountWhile(I2 + 1, each _ = " "),
            DigitsL = if HasL then CountWhile(I3, IsDigit) else 0,
            OkL = HasL and DigitsL >= 1 and DigitsL <= 5 and not IsAlnum(Chr(I3 + DigitsL))
        in
            if OkG then
                [Grade = "G" & Text.From(Number.From(Text.Middle(S, p + 1, DigitsG))),
                 Loss = if OkL then Number.From(Text.Middle(S, I3, DigitsL)) else 0]
            else null,
    Matches = List.RemoveNulls(List.Transform(Text.PositionOf(S, "G", Occurrence.All), MatchAt)),
    Last = List.Last(Matches, null)
in
    if Last <> null and List.Contains(cfgGrades, Last[Grade]) then Last else null
