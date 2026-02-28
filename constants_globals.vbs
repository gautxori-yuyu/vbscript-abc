Const strQuoteNrPattern = "\d{9}(?:[\-_]\d+)?"
Dim strQuoteNrRevPattern
strQuoteNrRevPattern = "(" & strQuoteNrPattern & ")(?:[ \-_]*rev\.?[ \-_]*\d+\b)?"
Const strModelPattern = "(\d)\s?T?\s*E\s?(H[AGPX])\s?\-\s?(\d)\s?\-\s?[LGT]{2,3}"
Dim strFullModelPattern
strFullModelPattern = strModelPattern & "(?:\-\d\x\d+T?)+(?: (?:NACE|ATEX))*"
Dim strOpModelsPattern
strOpModelsPattern = "((?:(?:" & strModelPattern & ")[ ,y]*)+|X{3,})"

Const strCustomerPattern = "((?:.(?! \- ))+?.(?:\s*[\-_]\s*(?:.(?! \- ))+.)*?)"
Const strOther_ProjectPattern = "((?:.(?! \- ))+?.(?:\s*[\-_]\s*(?:.(?! \- ))+.)*?)"
Dim strFilename_QuoteCustomerModelPattern
strFilename_QuoteCustomerModelPattern = "^" & strQuoteNrPattern & _
"\s*\-\s*" & strCustomerPattern & "(?:\s*\-\s*" & strOther_ProjectPattern & ")?\s*\-\s*" & strOpModelsPattern & "$"


