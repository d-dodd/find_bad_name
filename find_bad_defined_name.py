"""
find_bad_defined_name.py

Locate the defined name that Excel strips out when it reports:

    Removed Records: Named range from /xl/workbook.xml part (Workbook)

Reads xl/workbook.xml directly out of the .xlsm/.xlsx package (no Excel,
no COM) and runs validity checks on every <definedName>. Optionally diffs
against a repaired copy of the same workbook; the name present in the
original but missing from the repaired file is the culprit.

Usage:
    python find_bad_defined_name.py "C:\\path\\Corrupt.xlsm"
    python find_bad_defined_name.py "C:\\path\\Corrupt.xlsm" "C:\\path\\Repaired.xlsm"

IMPORTANT: run this against the ORIGINAL file, before you save the
repaired version over it. If Excel already overwrote it, grab the
pre-repair copy from OneDrive/SharePoint version history, a backup,
or Azure DevOps.
"""

import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# Excel name rules: first char must be a letter, underscore or backslash.
# Remaining chars may be letters, digits, period, underscore, backslash,
# question mark. Built-in names are prefixed with "_xlnm.".
VALID_NAME = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.\\?]*$")
A1_REF = re.compile(r"^[A-Za-z]{1,3}[0-9]{1,7}$")
R1C1_REF = re.compile(r"^[Rr](\[?-?[0-9]+\]?)?[Cc](\[?-?[0-9]+\]?)?$")
CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def read_part(path, part="xl/workbook.xml"):
    """Return the raw bytes of a part inside the OPC package."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if part not in names:
            raise KeyError("%s not found in %s" % (part, path))
        return z.read(part)


def parse_defined_names(raw):
    """
    Return (sheet_names, [dict, ...]) for each <definedName>.

    Falls back to a regex scan if the XML will not parse, since
    unparseable XML is itself a strong signal of where the damage is.
    """
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        print("  !! workbook.xml does not parse as XML: %s" % exc)
        print("  !! falling back to a raw regex scan")
        return None, regex_scan(raw)

    sheets = [s.get("name") for s in root.iter(NS + "sheet")]

    out = []
    for i, dn in enumerate(root.iter(NS + "definedName")):
        out.append(
            {
                "index": i,
                "name": dn.get("name"),
                "localSheetId": dn.get("localSheetId"),
                "hidden": dn.get("hidden"),
                "refersTo": (dn.text or ""),
            }
        )
    return sheets, out


def regex_scan(raw):
    """Crude recovery path for packages whose workbook.xml is malformed."""
    text = raw.decode("utf-8", errors="replace")
    out = []
    pattern = re.compile(r"<definedName\b([^>]*)>(.*?)</definedName>", re.S)
    for i, m in enumerate(pattern.finditer(text)):
        attrs = dict(re.findall(r'(\w+)\s*=\s*"([^"]*)"', m.group(1)))
        out.append(
            {
                "index": i,
                "name": attrs.get("name"),
                "localSheetId": attrs.get("localSheetId"),
                "hidden": attrs.get("hidden"),
                "refersTo": m.group(2),
            }
        )
    return out


def check(entry, sheet_count):
    """Return a list of reasons this defined name looks damaged."""
    problems = []
    name = entry["name"]
    refers = entry["refersTo"]

    if name is None:
        problems.append("no name attribute")
        return problems
    if name == "":
        problems.append("empty name")
        return problems

    if CONTROL_CHARS.search(name):
        problems.append("control characters in name")
    if any(ord(c) > 127 for c in name):
        problems.append("non-ASCII characters in name")
    if len(name) > 255:
        problems.append("name longer than 255 chars (%d)" % len(name))

    bare = name.split(".", 1)[1] if name.startswith("_xlnm.") else name
    if not name.startswith("_xlnm.") and not VALID_NAME.match(name):
        problems.append("name violates Excel naming rules")
    if A1_REF.match(bare) or R1C1_REF.match(bare):
        problems.append("name looks like a cell reference")

    if refers.strip() == "":
        problems.append("empty refersTo")
    else:
        if CONTROL_CHARS.search(refers):
            problems.append("control characters in refersTo")
        if refers.count("'") % 2:
            problems.append("unbalanced single quotes in refersTo")
        if refers.count("[") != refers.count("]"):
            problems.append("unbalanced brackets in refersTo")
        if len(refers) > 255:
            problems.append("refersTo longer than 255 chars (%d)" % len(refers))

    lsid = entry["localSheetId"]
    if lsid is not None:
        try:
            n = int(lsid)
        except ValueError:
            problems.append("localSheetId is not an integer: %r" % lsid)
        else:
            if sheet_count is not None and not (0 <= n < sheet_count):
                problems.append(
                    "localSheetId %d out of range (workbook has %d sheets)"
                    % (n, sheet_count)
                )

    return problems


def key(entry):
    return (entry["localSheetId"], entry["name"])


def describe(entry, sheets):
    scope = "workbook"
    lsid = entry["localSheetId"]
    if lsid is not None:
        try:
            n = int(lsid)
            scope = sheets[n] if sheets and 0 <= n < len(sheets) else "sheet#%s" % lsid
        except ValueError:
            scope = "sheet#%s" % lsid
    refers = entry["refersTo"]
    if len(refers) > 90:
        refers = refers[:90] + "..."
    return "[%-14s] %-40s -> %s" % (scope, entry["name"], refers)


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2

    original = argv[1]
    repaired = argv[2] if len(argv) > 2 else None

    print("Original: %s" % original)
    sheets, names = parse_defined_names(read_part(original))
    sheet_count = len(sheets) if sheets else None
    print("  sheets: %s" % (sheet_count if sheet_count is not None else "unknown"))
    print("  defined names: %d" % len(names))
    print("")

    flagged = []
    seen = {}
    for e in names:
        problems = check(e, sheet_count)
        k = key(e)
        if k in seen:
            problems.append("duplicate of entry #%d" % seen[k])
        else:
            seen[k] = e["index"]
        if problems:
            flagged.append((e, problems))

    if flagged:
        print("Suspect entries (%d):" % len(flagged))
        for e, problems in flagged:
            print("  #%d %s" % (e["index"], describe(e, sheets)))
            for p in problems:
                print("        - %s" % p)
        print("")
    else:
        print("No entries failed the validity checks.")
        print("The damage may be structural (encoding, truncation, bad entities).")
        print("Diff against the repaired copy to be certain.\n")

    if repaired:
        print("Repaired: %s" % repaired)
        r_sheets, r_names = parse_defined_names(read_part(repaired))
        print("  defined names: %d" % len(r_names))
        before = {key(e): e for e in names}
        after = {key(e): e for e in r_names}

        gone = [before[k] for k in before if k not in after]
        added = [after[k] for k in after if k not in before]
        changed = [
            (before[k], after[k])
            for k in before
            if k in after and before[k]["refersTo"] != after[k]["refersTo"]
        ]

        print("")
        if gone:
            print("REMOVED BY THE REPAIR (%d) -- this is your answer:" % len(gone))
            for e in gone:
                print("  %s" % describe(e, sheets))
                print("    raw name repr: %r" % e["name"])
                print("    raw refersTo:  %r" % e["refersTo"])
        else:
            print("Nothing removed by name/scope. Check the 'changed' list below.")

        if changed:
            print("\nREFERSTO REWRITTEN (%d):" % len(changed))
            for b, a in changed:
                print("  %s" % b["name"])
                print("    before: %r" % b["refersTo"])
                print("    after:  %r" % a["refersTo"])

        if added:
            print("\nADDED BY THE REPAIR (%d):" % len(added))
            for e in added:
                print("  %s" % describe(e, r_sheets))
    else:
        print("Tip: save the repaired workbook under a new name and re-run as")
        print("  python %s original.xlsm repaired.xlsm" % os.path.basename(argv[0]))
        print("to get an exact diff.")

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
