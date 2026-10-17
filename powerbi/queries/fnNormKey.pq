// Matching key for headers and event names: fnCleanText + unify Arabic
// letter variants (alef forms, alef maqsura, taa marbuta) + lower case.
(value as any) as nullable text =>
let
    Clean = fnCleanText(value),
    Pairs = {{"آ", "ا"}, {"أ", "ا"}, {"إ", "ا"}, {"ى", "ي"}, {"ة", "ه"}}
in
    if Clean = null then null
    else Text.Lower(List.Accumulate(Pairs, Clean, (acc, p) => Text.Replace(acc, p{0}, p{1})))
