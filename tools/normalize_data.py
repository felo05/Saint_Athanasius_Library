#!/usr/bin/env python3
"""Clean up and normalize data.csv so the same thing is always written the same way.

Usage (from the repository root):
    python3 tools/normalize_data.py                 # rewrite data.csv in place
    python3 tools/normalize_data.py --dry-run       # only write the report
    python3 tools/normalize_data.py --report changes.csv

Safe to run again after adding new books: it only changes what is inconsistent.
Every change is written to the report (row number as seen in Excel, column,
before, after, reason).

What it does
  * Every cell: replace Arabic ligature characters (ﻷ ﻻ ...) with normal
    letters, remove tatweel (ـ), turn repeated spaces into one space, trim,
    and tidy spacing around the Arabic comma. A line break between several
    people (authors, translators, their titles, saints in the category)
    becomes "، "; elsewhere it becomes a space.
  * Spelling variants: values that differ only by hamza forms (أ إ آ ا),
    ة/ه, ى/ي, ؤ/و, ئ/ي, diacritics or spacing around punctuation are unified
    to the spelling used most often in the file (or to one listed in
    PREFERRED_SPELLINGS when the common one is a misspelling). Author and
    translator names are unified together, as are author and translator
    titles (القس, الأب ...).
  * التصنيف: 'ا' -> 'أ'.
  * الطبعة: words -> numbers (الاولى -> 1, First -> 1, ...).
  * الجزء: 'ج1' -> '1'; lists of parts written '1\\2', '1,2', '1 & 2', ...
    -> '1، 2'; repairs values Excel turned into dates/times.
  * م. رف: tidy shelf codes ('د12- 1' -> 'د12-1').
  * Person titles always take "ال" (دكتور -> الدكتور); values that differ only
    by "ال" (بالهرم / الهرم) are variants too, except in book titles.
  * سنة النشر: fix clear typos (19962 -> 1996, 78 -> 1978, 88 -> 1988).
    Other impossible years are left as they are and listed in the report.
"""
import argparse
import collections
import csv
import io
import re
import sys
import unicodedata

DATA_FILE = "data.csv"

TASHKEEL = re.compile(r"[ً-ٰٟ]")
TATWEEL = "ـ"

EDITION_WORDS = {
    "الاولى": "1", "الأولى": "1", "الاولي": "1", "الأولي": "1", "first": "1",
    "الثانية": "2", "الثانيه": "2", "second": "2",
    "الثالثة": "3", "الثالثه": "3", "third": "3",
    "الرابعة": "4", "الرابعه": "4", "fourth": "4",
    "الخامسة": "5", "الخامسه": "5", "fifth": "5",
}
PART_WORDS = {"الاول": "1", "الأول": "1", "الثاني": "2", "الثانى": "2", "الثالث": "3"}
YEAR_FIXES = {"19962": "1996", "78": "1978", "88": "1988"}
# When one of these spellings is in a group of variants it wins, even if a
# misspelling is more common (e.g. "الإباء" means "pride", not "the fathers").
PREFERRED_SPELLINGS = {
    "الآباء السواح",
    "ولم يحبوا حياتهم حتى الموت",
    "في ذكرى شهداء المسيحية",
    "ألق على الرب همك فهو يعولك",
    "موسوعة المعرفة المسيحية آباء الكنيسة",
    "المجلس الوطني للثقافة والفنون والآداب",
    "مكتبة كنيسة الإخوة",
    "مكتبة الإخوة",
}
# Person titles always take "ال" (دكتور -> الدكتور, أسقف عام -> الأسقف العام),
# like القمص / القس / الأب. Only entries made up entirely of titles are changed:
# "أسقف الغربية" / "مطران دمياط" / "أستاذ اللاهوت" (title + what it belongs to)
# can't take "ال" in Arabic, so they stay as written.
TITLE_COLUMNS = {"لقب المؤلف", "لقب المترجم"}
TITLE_WORDS = {
    "دكتور", "دكتورة", "أسقف", "اسقف", "دياكون", "مطران", "عام", "مهندس", "أستاذ",
    "استاذ", "أرشيدياكون", "ارشيدياكون", "راهب", "اكليريكي", "إكليريكي", "أغنسطس",
    "اغنسطس", "شماس", "خادم", "مدرس", "خوري", "بطريرك",
}
# Titles that already carry "ال" and may appear in an all-title entry
AL_TITLES = {"القس", "القمص", "الأب", "الاب", "البطريرك", "الشماس", "الدكتور", "الدكتورة",
             "الأستاذ", "الاستاذ", "المهندس", "الدياكون", "الأرشيدياكون", "الراهب",
             "الأسقف", "الاسقف", "العام", "الإكليريكي", "الاكليريكي", "المتنيح"}
MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def header_key(h):
    """Header names contain line breaks ("سنة\\nالنشر"); compare them without."""
    return re.sub(r"\s+", " ", h.replace("﻿", "")).strip()


# Columns where a line break separates several people / saints: it becomes "، "
LIST_COLUMNS = {"لقب المؤلف", "اسم المؤلف", "لقب المترجم", "اسم المترجم", "تصنيف الكتاب"}


def clean_cell(value, is_list=False):
    # Arabic presentation forms (ligatures like ﻷ) -> normal letters
    value = "".join(
        unicodedata.normalize("NFKC", ch) if 0xFB50 <= ord(ch) <= 0xFDFF or 0xFE70 <= ord(ch) <= 0xFEFF else ch
        for ch in value
    )
    value = value.replace(TATWEEL, "")
    if is_list:
        lines = [re.sub(r"\s+", " ", ln).strip() for ln in value.strip().split("\n")]
        lines = [ln for ln in lines if ln]
        value = ""
        for ln in lines:
            if not value:
                value = ln
            elif value.endswith(" و") or value == "و":   # "... و" + next name: already joined
                value += " " + ln
            else:
                value += "، " + ln
    value = re.sub(r"\s+", " ", value).strip()
    # Arabic comma: no space before, one space after
    value = re.sub(r"\s*،\s*", "، ", value).strip()
    value = value.rstrip("،").strip()
    return value


def variant_key(value):
    """Two values with the same key are spelling variants of each other."""
    v = TASHKEEL.sub("", value)
    v = re.sub("[أإآٱ]", "ا", v)
    v = v.replace("ة", "ه").replace("ى", "ي").replace("ؤ", "و").replace("ئ", "ي")
    v = re.sub(r"\s*([-–—()،,.:/\\])\s*", r"\1", v)   # spacing around punctuation
    v = re.sub(r"\s+", " ", v).strip().lower()
    return v


def strip_al_key(value):
    """Like variant_key, but also ignores a leading ال / بال / وال / لل on each
    word, so "بالهرم" ~ "الهرم" and "سير قديسين" ~ "سير القديسين"."""
    return " ".join(re.sub(r"^(ال|بال|وال|لل)(?=\S\S)", "", w) for w in variant_key(value).split())


def add_al_to_titles(value):
    entries = []
    for entry in value.split("، "):
        words = entry.split(" ")
        all_titles = all(w in TITLE_WORDS or w in AL_TITLES or w == "-" for w in words)
        if all_titles:
            words = ["ال" + w if w in TITLE_WORDS else w for w in words]
        entries.append(" ".join(words))
    return "، ".join(entries)


def spelling_score(value):
    """Tie-breaker when two spellings are equally common: prefer written
    hamza and ة, then standard spacing around ' - '."""
    return (
        len(re.findall(r"(^|\s)ال", value)),
        len(re.findall("[أإآؤئ]", value)),
        value.count("ة"),
        value.count(" - "),
    )


def clean_parts(value):
    v = value
    if not v:
        return v
    # Excel turned "1-15" into a date "15-Jan" and "6:7" into a time "6:07"
    m = re.fullmatch(r"(\d{1,2})-([A-Za-z]{3})", v)
    if m and m.group(2).lower() in MONTHS:
        return f"{MONTHS[m.group(2).lower()]}-{m.group(1)}"
    m = re.fullmatch(r"(\d{1,2}):0?(\d{1,2})", v)
    if m:
        return f"{m.group(1)}، {m.group(2)}"
    v = PART_WORDS.get(v, v)
    v = re.sub(r"^ج\s*(?=\d)", "", v)                    # ج1 -> 1
    if re.fullmatch(r"\d+(\s*[\\,&،.\s]\s*\d+)+", v):    # 1\2, 1,2, 1 & 2, 1.2.3, "1 2"
        v = "، ".join(re.findall(r"\d+", v))
    v = re.sub(r"(\d)\s+([أ-ي])$", r"\1\2", v)          # "19 أ" -> "19أ"
    v = re.sub(r"\s*-\s*", "-", v)                       # "37- 1" -> "37-1"
    return v


def clean_shelf(value):
    v = re.sub(r"\s*-+\s*", "-", value)                  # "د12- 1", "د20--1" -> "د12-1"
    return v


# Part markers inside a book title: "ج1", "ج 2", "ج.3", "جزء1", "الجزء 2",
# "الجزء الأول", "جزء ثاني" ...
ORDINALS = {
    "الأول": 1, "الاول": 1, "أول": 1, "اول": 1,
    "الثاني": 2, "الثانى": 2, "ثاني": 2, "ثانى": 2,
    "الثالث": 3, "ثالث": 3, "الرابع": 4, "رابع": 4, "الخامس": 5, "خامس": 5,
    "السادس": 6, "سادس": 6, "السابع": 7, "سابع": 7, "الثامن": 8, "ثامن": 8,
    "التاسع": 9, "تاسع": 9, "العاشر": 10, "عاشر": 10,
}
_ORD = "|".join(sorted(map(re.escape, ORDINALS), key=len, reverse=True))
PART_MARKER = re.compile(
    r"(?<![\w])(?:"
    r"ج\s*\.?\s*(?P<n1>\d+)"                               # ج1  ج 2  ج.3
    r"|(?:ال)?جزء\s*(?P<n2>\d+)"                           # جزء1  الجزء 2
    r"|(?:ال)?جزء\s+(?P<w>" + _ORD + r")"                   # الجزء الأول  جزء ثاني
    r")(?![\w])"
)


def remove_part_from_title(title, part):
    """If the title names the same part number as the الجزء column, drop it
    from the title: "تفسير انجيل متى ج1" + part 1 -> "تفسير انجيل متى".
    Titles with more than one marker, or a different number, are left alone."""
    if not part.isdigit():
        return title
    markers = list(PART_MARKER.finditer(title))
    if len(markers) != 1:
        return title
    m = markers[0]
    n = m.group("n1") or m.group("n2") or ORDINALS.get(m.group("w"))
    if int(n) != int(part):
        return title
    t = title[:m.start()] + " " + title[m.end():]
    # Tidy what's left around the removed marker
    t = re.sub(r"\(\s*\)", " ", t)                          # empty brackets
    t = re.sub(r"\(\s+", "(", t)
    t = re.sub(r"\s+\)", ")", t)
    t = re.sub(r"\s*-\s*(?:-\s*)+", " - ", t)                # "- -" -> "-"
    t = re.sub(r"\s+", " ", t).strip()
    t = re.sub(r"^[\s\-–:،]+|[\s\-–:،]+$", "", t)           # separators left at the ends
    return t or title


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--file", default=DATA_FILE)
    ap.add_argument("--report", default="data-cleanup-report.csv")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    raw = open(args.file, "rb").read()
    text = raw.decode("utf-8-sig")
    rows = list(csv.reader(io.StringIO(text, newline="")))
    header, data = rows[0], rows[1:]
    col = {header_key(h): j for j, h in enumerate(header)}

    changes = []  # (excel_row, column, before, after, reason)

    def set_cell(i, j, new, reason):
        old = data[i][j]
        if new != old:
            changes.append((i + 2, header_key(header[j]), old, new, reason))
            data[i][j] = new

    # 1. Per-cell clean-up
    for i, row in enumerate(data):
        for j, v in enumerate(row):
            set_cell(i, j, clean_cell(v, header_key(header[j]) in LIST_COLUMNS), "spaces / ligatures / tatweel")

    # 2. Column-specific rules
    special = {
        "التصنيف": lambda v: "أ" if v == "ا" else v,
        "الطبعة": lambda v: EDITION_WORDS.get(v.lower(), v),
        "الجزء": clean_parts,
        "م. رف": clean_shelf,
        "سنة النشر": lambda v: YEAR_FIXES.get(v, v),
        "لقب المؤلف": add_al_to_titles,
        "لقب المترجم": add_al_to_titles,
    }
    reasons = {
        "التصنيف": "category letter", "الطبعة": "edition as number", "الجزء": "part format",
        "م. رف": "shelf code format", "سنة النشر": "year typo",
        "لقب المؤلف": "title with ال", "لقب المترجم": "title with ال",
    }
    for name, fn in special.items():
        if name not in col:
            continue
        j = col[name]
        for i in range(len(data)):
            set_cell(i, j, fn(data[i][j]), reasons[name])

    # Part number written in both the title and the الجزء column: keep it in the column only
    if "اسم الكتاب" in col and "الجزء" in col:
        jt, jp = col["اسم الكتاب"], col["الجزء"]
        for i, row in enumerate(data):
            set_cell(i, jt, remove_part_from_title(row[jt], row[jp]), "part removed from title")

    # 3. Unify spelling variants to the most common spelling.
    #    Columns in the same tuple share one vocabulary.
    groups = [
        ("اسم الكتاب",),
        ("السلسلة",),
        ("لقب المؤلف", "لقب المترجم"),
        ("اسم المؤلف", "اسم المترجم"),
        ("الناشر",),
        ("تصنيف الكتاب",),
    ]
    for names in groups:
        js = [col[n] for n in names if n in col]
        # Book titles keep their exact wording; elsewhere "ال" differences are variants
        key_fn = variant_key if names == ("اسم الكتاب",) else strip_al_key
        counts = collections.defaultdict(collections.Counter)
        for row in data:
            for j in js:
                if row[j]:
                    counts[key_fn(row[j])][row[j]] += 1
        canonical = {}
        for key, spellings in counts.items():
            if len(spellings) > 1:
                preferred = [s for s in spellings if s in PREFERRED_SPELLINGS]
                best = preferred[0] if preferred else max(
                    spellings, key=lambda s: (spellings[s], spelling_score(s), s))
                canonical[key] = best
        for i, row in enumerate(data):
            for j in js:
                if row[j]:
                    best = canonical.get(key_fn(row[j]))
                    if best:
                        set_cell(i, j, best, "spelling variant")

    # 4. Things to check by hand (not changed)
    notes = []
    if "سنة النشر" in col:
        j = col["سنة النشر"]
        for i, row in enumerate(data):
            v = row[j]
            if v and not (v.isdigit() and 1000 <= int(v) <= 2100):
                notes.append((i + 2, "سنة النشر", v, v, "CHECK: impossible year (not changed)"))
    if "اسم الكتاب" in col:
        j = col["اسم الكتاب"]
        for i, row in enumerate(data):
            if not row[j]:
                notes.append((i + 2, "اسم الكتاب", "", "", "CHECK: book has no title"))

    # Report: one line per changed cell (last change wins; first 'before')
    merged = collections.OrderedDict()
    for r, c, before, after, reason in changes:
        k = (r, c)
        if k in merged:
            merged[k] = (merged[k][0], after, merged[k][2] + " + " + reason)
        else:
            merged[k] = (before, after, reason)
    report_rows = [(r, c, b, a, why) for (r, c), (b, a, why) in merged.items() if b != a] + notes
    report_rows.sort(key=lambda x: (x[0], x[1]))
    with open(args.report, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(["الصف", "العمود", "قبل", "بعد", "السبب"])
        w.writerows(report_rows)

    changed_cells = sum(1 for x in report_rows if not x[4].startswith("CHECK"))
    by_reason = collections.Counter(x[4] for x in report_rows)
    print(f"{changed_cells} cells changed in {len({x[0] for x in report_rows if not x[4].startswith('CHECK')})} rows")
    for reason, n in by_reason.most_common():
        print(f"  {n:5}  {reason}")
    print(f"report: {args.report}")

    if not args.dry_run:
        out = io.StringIO()
        csv.writer(out, lineterminator="\r\n").writerows([header] + data)
        with open(args.file, "wb") as f:
            f.write(b"\xef\xbb\xbf" + out.getvalue().encode("utf-8"))
        print(f"wrote {args.file}")


if __name__ == "__main__":
    sys.exit(main())
