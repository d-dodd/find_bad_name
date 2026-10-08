# find_bad_name
Reads xl/workbook.xml directly out of the .xlsm/.xlsx package (no Excel, no COM) and runs validity checks on every <definedName>. Optionally diffs against a repaired copy of the same workbook; the name present in the original but missing from the repaired file is the culprit.
