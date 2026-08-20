#!/usr/bin/env python3
"""Static audit of an inherited .xlsx workbook.

Reads formulas (never values) and reports five classes of defect:
hardcoded constants inside formulas, formulas that break the pattern of
their row, aggregate ranges that stop short of their block, links and
error remnants, and calculations nothing depends on.

Rankings in this report are structural. The script cannot tell you how
much of the final answer a defect moves, because it never evaluates the
workbook. It ranks by severity times reach and prints the cells to
perturb by hand to obtain the real ranking.

Standard library plus openpyxl. No network access, no writes to the
workbook under audit.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import math
import os
import re
import sys
from collections import defaultdict

try:
    from openpyxl import load_workbook
    from openpyxl.utils import column_index_from_string, get_column_letter
except ImportError:  # pragma: no cover
    sys.stderr.write("openpyxl is required: pip install openpyxl\n")
    raise SystemExit(3)

# --------------------------------------------------------------------------
# formula tokeniser
# --------------------------------------------------------------------------

_SHEETNAME = (
    r"(?:'(?:[^']|'')+'"
    r"|\[\d+\][^\s!,;()+\-*/^&<>=:\"'\[\]]+"
    r"|[^\s!,;()+\-*/^&<>=:\"'\[\]]+)"
)
SHEETPFX_RE = re.compile(_SHEETNAME + r"(?::" + _SHEETNAME + r")?!")
REFBODY_RE = re.compile(
    r"(?:\$?[A-Za-z]{1,3}\$?\d{1,7}:\$?[A-Za-z]{1,3}\$?\d{1,7}"
    r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
    r"|\$?\d{1,7}:\$?\d{1,7}"
    r"|\$?[A-Za-z]{1,3}\$?\d{1,7})"
    r"(?![A-Za-z0-9_.(])"
)
NAME_RE = re.compile(r"[A-Za-z_\\\u00c0-\uffff][A-Za-z0-9_.\\\u00c0-\uffff]*")
NUM_RE = re.compile(r"(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")
ERR_RE = re.compile(
    r"#REF!|#DIV/0!|#N/A|#VALUE!|#NAME\?|#NULL!|#NUM!|#SPILL!|#CALC!|#GETTING_DATA"
)
ADDR_RE = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})\Z")
COLONLY_RE = re.compile(r"(\$?)([A-Za-z]{1,3})\Z")
ROWONLY_RE = re.compile(r"(\$?)(\d{1,7})\Z")


class Tok:
    __slots__ = ("kind", "text", "pos", "func", "arg")

    def __init__(self, kind, text, pos, ctx):
        self.kind = kind
        self.text = text
        self.pos = pos
        self.func, self.arg = ctx

    def __repr__(self):  # pragma: no cover
        return f"Tok({self.kind},{self.text!r},{self.func}#{self.arg})"


def tokenize(formula):
    """Split a formula into tokens, tagging each with its enclosing
    function name and argument index (0-based)."""
    s = formula[1:] if formula.startswith("=") else formula
    toks = []
    stack = []
    prev_ident = None
    i, n = 0, len(s)
    while i < n:
        ch = s[i]
        ctx = (stack[-1][0], stack[-1][1]) if stack else ("", -1)
        if ch in " \t\r\n":
            i += 1
            continue
        if ch == '"':
            j = i + 1
            while j < n:
                if s[j] == '"':
                    if j + 1 < n and s[j + 1] == '"':
                        j += 2
                        continue
                    j += 1
                    break
                j += 1
            toks.append(Tok("str", s[i:j], i, ctx))
            i, prev_ident = j, None
            continue
        if ch == "#":
            m = ERR_RE.match(s, i)
            if m:
                toks.append(Tok("err", m.group(0), i, ctx))
                i, prev_ident = m.end(), None
                continue
        m = SHEETPFX_RE.match(s, i)
        if m:
            body = REFBODY_RE.match(s, m.end())
            if body:
                toks.append(Tok("ref", s[i:body.end()], i, ctx))
                i, prev_ident = body.end(), None
                continue
            nm = NAME_RE.match(s, m.end())
            if nm:
                toks.append(Tok("name", s[i:nm.end()], i, ctx))
                i, prev_ident = nm.end(), None
                continue
            err = ERR_RE.match(s, m.end())
            if err:
                toks.append(Tok("err", s[i:err.end()], i, ctx))
                i, prev_ident = err.end(), None
                continue
        m = REFBODY_RE.match(s, i)
        if m:
            toks.append(Tok("ref", m.group(0), i, ctx))
            i, prev_ident = m.end(), None
            continue
        if ch.isdigit() or (ch == "." and i + 1 < n and s[i + 1].isdigit()):
            m = NUM_RE.match(s, i)
            if m:
                toks.append(Tok("num", m.group(0), i, ctx))
                i, prev_ident = m.end(), None
                continue
        m = NAME_RE.match(s, i)
        if m:
            j = m.end()
            k = j
            while k < n and s[k] in " \t":
                k += 1
            if k < n and s[k] == "(":
                toks.append(Tok("func", m.group(0), i, ctx))
                prev_ident = m.group(0)
            else:
                toks.append(Tok("name", m.group(0), i, ctx))
                prev_ident = None
            i = j
            continue
        if ch == "(":
            toks.append(Tok("op", "(", i, ctx))
            stack.append([(prev_ident or "").upper(), 0])
            i, prev_ident = i + 1, None
            continue
        if ch == ")":
            toks.append(Tok("op", ")", i, ctx))
            if stack:
                stack.pop()
            i, prev_ident = i + 1, None
            continue
        if ch in ",;":
            if stack:
                stack[-1][1] += 1
            toks.append(Tok("op", ",", i, ctx))
            i, prev_ident = i + 1, None
            continue
        toks.append(Tok("op", ch, i, ctx))
        i, prev_ident = i + 1, None
    return toks


# --------------------------------------------------------------------------
# canonical (R1C1-equivalent) form
# --------------------------------------------------------------------------


def _canon_addr(a, row, col):
    m = ADDR_RE.match(a)
    if m:
        cabs, cl, rabs, rn = m.groups()
        c, r = column_index_from_string(cl.upper()), int(rn)
        cs = f"C{c}" if cabs else ("C" if c == col else f"C[{c - col:+d}]")
        rs = f"R{r}" if rabs else ("R" if r == row else f"R[{r - row:+d}]")
        return rs + cs
    m = COLONLY_RE.match(a)
    if m:
        cabs, cl = m.groups()
        c = column_index_from_string(cl.upper())
        return f"C{c}" if cabs else ("C" if c == col else f"C[{c - col:+d}]")
    m = ROWONLY_RE.match(a)
    if m:
        rabs, rn = m.groups()
        r = int(rn)
        return f"R{r}" if rabs else ("R" if r == row else f"R[{r - row:+d}]")
    return a.upper()


def canon_ref(text, row, col):
    m = SHEETPFX_RE.match(text)
    pfx, body = "", text
    if m:
        pfx, body = m.group(0).upper(), text[m.end():]
    return pfx + ":".join(_canon_addr(p, row, col) for p in body.split(":"))


def canonical(formula, row, col):
    """R1C1-equivalent normal form. Two formulas that Excel would show as
    the same fill pattern produce the same string."""
    out = []
    for t in tokenize(formula):
        if t.kind == "ref":
            out.append(canon_ref(t.text, row, col))
        elif t.kind in ("func", "name"):
            out.append(t.text.upper())
        else:
            out.append(t.text)
    return "".join(out)


# --------------------------------------------------------------------------
# reference resolution
# --------------------------------------------------------------------------


class Ref:
    __slots__ = ("sheet", "r1", "c1", "r2", "c2", "external", "cross", "raw")

    def __init__(self, sheet, r1, c1, r2, c2, external, cross, raw):
        self.sheet, self.raw = sheet, raw
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2
        self.external, self.cross = external, cross


def _split_addr(a, default_row, default_col_max):
    m = ADDR_RE.match(a)
    if m:
        return int(m.group(4)), column_index_from_string(m.group(2).upper())
    m = COLONLY_RE.match(a)
    if m:
        return None, column_index_from_string(m.group(2).upper())
    m = ROWONLY_RE.match(a)
    if m:
        return int(m.group(2)), None
    return None, None


def resolve_ref(text, host_sheet, sheet_bounds):
    m = SHEETPFX_RE.match(text)
    sheet, body, external, cross = host_sheet, text, False, False
    if m:
        raw_sheet = m.group(0)[:-1]
        body = text[m.end():]
        if "[" in raw_sheet:
            external = True
        name = raw_sheet.split(":")[0].strip("'").replace("''", "'")
        name = re.sub(r"^\[\d+\]", "", name)
        sheet, cross = name, True
    parts = body.split(":")
    a = _split_addr(parts[0], None, None)
    b = _split_addr(parts[-1], None, None) if len(parts) > 1 else a
    maxr, maxc = sheet_bounds.get(sheet, (1, 1))
    r1 = a[0] if a[0] is not None else 1
    c1 = a[1] if a[1] is not None else 1
    r2 = b[0] if b[0] is not None else maxr
    c2 = b[1] if b[1] is not None else maxc
    if r1 > r2:
        r1, r2 = r2, r1
    if c1 > c2:
        c1, c2 = c2, c1
    return Ref(sheet, r1, c1, r2, c2, external, cross, text)


# --------------------------------------------------------------------------
# literal classification
# --------------------------------------------------------------------------

BENIGN = {0.0, 1.0, -1.0, 12.0, 100.0}
STRUCTURAL_ARGS = {
    "VLOOKUP": {2, 3}, "HLOOKUP": {2, 3}, "XLOOKUP": {3, 4, 5},
    "INDEX": {1, 2}, "MATCH": {2}, "OFFSET": {1, 2, 3, 4},
    "ROUND": {1}, "ROUNDUP": {1}, "ROUNDDOWN": {1}, "TRUNC": {1},
    "LEFT": {1}, "RIGHT": {1}, "MID": {1, 2}, "REPT": {1},
    "LARGE": {1}, "SMALL": {1}, "RANK": {2}, "SUBTOTAL": {0},
    "CHOOSE": {0}, "WEEKDAY": {1}, "DATE": {0, 1, 2}, "EDATE": {1},
    "EOMONTH": {1}, "INDIRECT": {1}, "AGGREGATE": {0, 1},
    "NETWORKDAYS": set(), "SUMIFS": set(), "COUNTIFS": set(),
}


def classify_literal(tok, prev_op):
    """Return (severity, label) or None when the literal is structural."""
    try:
        v = float(tok.text)
    except ValueError:
        return None
    if v in BENIGN:
        return None
    if prev_op == "^":
        return None
    fn = tok.func
    if fn in STRUCTURAL_ARGS and tok.arg in STRUCTURAL_ARGS[fn]:
        return None
    if v != int(v):
        return ("high", "rate or factor")
    iv = int(v)
    if 1900 <= iv <= 2200:
        return ("low", "period literal")
    if abs(iv) >= 1000:
        return ("medium", "magnitude")
    return ("medium", "count or factor")


# --------------------------------------------------------------------------
# workbook load
# --------------------------------------------------------------------------

AGGREGATES = {
    "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "MEDIAN",
    "STDEV", "STDEV.S", "STDEV.P", "PRODUCT", "AVERAGEA",
}
TOTAL_WORDS = ("total", "toplam", "subtotal", "ara toplam", "genel toplam", "sum of")


def formula_text(value):
    if isinstance(value, str) and value.startswith("="):
        return value
    text = getattr(value, "text", None)
    if isinstance(text, str) and text.startswith("="):
        return text
    return None


def read_workbook(path, max_cells):
    wb = load_workbook(path, data_only=False, read_only=True, keep_links=True)
    grid = {}        # sheet -> {(row, col): raw value}
    formulas = {}    # (sheet, row, col) -> formula text
    bounds = {}
    total = 0
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            cells, maxr, maxc = {}, 1, 1
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if v is None or (isinstance(v, str) and not v.strip()):
                        continue
                    total += 1
                    if total > max_cells:
                        raise MemoryError(
                            f"workbook exceeds --max-cells ({max_cells}); "
                            "raise the limit or audit sheet by sheet"
                        )
                    key = (cell.row, cell.column)
                    cells[key] = v
                    maxr = max(maxr, cell.row)
                    maxc = max(maxc, cell.column)
                    ft = formula_text(v)
                    if ft:
                        formulas[(name, cell.row, cell.column)] = ft
            grid[name] = cells
            bounds[name] = (maxr, maxc)
        defined = []
        try:
            for dn, obj in wb.defined_names.items():
                defined.append((dn, getattr(obj, "value", "") or ""))
        except Exception:
            pass
        ext = []
        try:
            for link in (wb._external_links or []):
                target = getattr(getattr(link, "file_link", None), "Target", None)
                if target:
                    ext.append(str(target))
        except Exception:
            pass
    finally:
        wb.close()
    return grid, formulas, bounds, defined, ext


def addr(sheet, row, col):
    return f"{sheet}!{get_column_letter(col)}{row}"


# --------------------------------------------------------------------------
# analysis
# --------------------------------------------------------------------------


def build_graph(formulas, bounds, grid, sheetnames):
    """precedents[cell] = set(cells it reads); dependents = reverse."""
    precedents = defaultdict(set)
    dependents = defaultdict(set)
    sheet_edges = defaultdict(int)
    externals, errors, crosslinks = [], [], []
    unknown_sheets = set()
    for cell, ftext in formulas.items():
        sheet, row, col = cell
        for t in tokenize(ftext):
            if t.kind == "err":
                errors.append((cell, ftext, t.text))
                continue
            if t.kind != "ref":
                continue
            ref = resolve_ref(t.text, sheet, bounds)
            if ref.external:
                externals.append((cell, ftext, t.text))
                continue
            if ref.sheet not in grid:
                unknown_sheets.add(ref.sheet)
                continue
            if ref.cross and ref.sheet != sheet:
                crosslinks.append((cell, ref.sheet, t.text))
                sheet_edges[(sheet, ref.sheet)] += 1
            area = (ref.r2 - ref.r1 + 1) * (ref.c2 - ref.c1 + 1)
            cells = grid[ref.sheet]
            if area <= 4096:
                span = (
                    (ref.sheet, r, c)
                    for r in range(ref.r1, ref.r2 + 1)
                    for c in range(ref.c1, ref.c2 + 1)
                    if (r, c) in cells
                )
            else:
                span = (
                    (ref.sheet, r, c)
                    for (r, c) in cells
                    if ref.r1 <= r <= ref.r2 and ref.c1 <= c <= ref.c2
                )
            for p in span:
                if p == cell:
                    continue
                precedents[cell].add(p)
                dependents[p].add(cell)
    return precedents, dependents, sheet_edges, externals, errors, crosslinks, unknown_sheets


def transitive_reach(dependents, targets, cap=20000):
    """Number of distinct formula cells downstream of each target cell.

    Computed per target rather than for the whole graph, so a large
    workbook costs one traversal per cell actually reported.
    """
    reach = {}
    for start in targets:
        seen, stack = set(), [start]
        while stack:
            node = stack.pop()
            for d in dependents.get(node, ()):
                if d not in seen:
                    seen.add(d)
                    if len(seen) >= cap:
                        stack = []
                        break
                    stack.append(d)
        reach[start] = len(seen)
    return reach


def find_cycles(precedents, limit=20):
    WHITE, GREY, BLACK = 0, 1, 2
    colour = defaultdict(int)
    cycles = []
    for start in list(precedents):
        if colour[start] != WHITE:
            continue
        colour[start] = GREY
        path = [start]
        stack = [(start, iter(sorted(precedents.get(start, ()))))]
        while stack:
            node, it = stack[-1]
            nxt = next(it, None)
            if nxt is None:
                colour[node] = BLACK
                stack.pop()
                path.pop()
                continue
            if colour[nxt] == GREY:
                if nxt in path:
                    cycles.append(path[path.index(nxt):] + [nxt])
                    if len(cycles) >= limit:
                        return cycles
                continue
            if colour[nxt] == BLACK:
                continue
            colour[nxt] = GREY
            path.append(nxt)
            stack.append((nxt, iter(sorted(precedents.get(nxt, ())))))
    return cycles


def scan_hardcodes(formulas):
    out = []
    for cell, ftext in sorted(formulas.items()):
        toks = tokenize(ftext)
        for idx, t in enumerate(toks):
            if t.kind != "num":
                continue
            prev = toks[idx - 1] if idx else None
            prev_op = prev.text if prev is not None and prev.kind == "op" else ""
            verdict = classify_literal(t, prev_op)
            if verdict is None:
                continue
            sev, label = verdict
            out.append({
                "kind": "hardcode",
                "cell": cell,
                "formula": ftext,
                "literal": t.text,
                "label": label,
                "severity": sev,
                "note": (
                    f"literal {t.text} ({label}) inside "
                    + (f"{t.func}() argument {t.arg + 1}" if t.func else "arithmetic")
                ),
            })
    return out


def scan_row_consistency(grid, formulas, min_run, share):
    out = []
    for sheet, cells in sorted(grid.items()):
        by_row = defaultdict(list)
        for (r, c), v in cells.items():
            ft = formula_text(v)
            if ft is not None:
                by_row[r].append((c, "F", ft))
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                by_row[r].append((c, "K", v))
        for r, items in sorted(by_row.items()):
            items.sort()
            run = []
            for entry in items + [(None, None, None)]:
                if run and entry[0] == run[-1][0] + 1:
                    run.append(entry)
                    continue
                out.extend(_check_run(sheet, r, run, min_run, share))
                run = [entry] if entry[0] is not None else []
    return out


def _check_run(sheet, r, run, min_run, share):
    if len(run) < min_run:
        return []
    patterns = []
    for c, kind, v in run:
        patterns.append(canonical(v, r, c) if kind == "F" else "#CONSTANT")
    counts = defaultdict(int)
    for p in patterns:
        counts[p] += 1
    modal, k = max(counts.items(), key=lambda kv: (kv[1], kv[0]))
    if modal == "#CONSTANT" or k < 3 or k / len(run) < share:
        return []
    example = next(c for (c, _, _), p in zip(run, patterns) if p == modal)
    findings = []
    for (c, kind, v), p in zip(run, patterns):
        if p == modal:
            continue
        sev = "high" if p == "#CONSTANT" else "medium"
        what = ("hardcoded value {!r} where the row calculates".format(v)
                if p == "#CONSTANT" else "formula differs from the rest of the row")
        findings.append({
            "kind": "row_break",
            "cell": (sheet, r, c),
            "formula": v if kind == "F" else repr(v),
            "severity": sev,
            "note": what,
            "majority": modal,
            "majority_share": f"{k}/{len(run)}",
            "example": addr(sheet, r, example),
        })
    return findings


def _label_for_row(cells, r, c):
    for cc in range(1, c):
        v = cells.get((r, cc))
        if isinstance(v, str) and v.strip():
            return v.strip()
    return ""


def _is_total_label(text):
    t = (text or "").strip().lower()
    return any(w in t for w in TOTAL_WORDS)


def collect_aggregates(formulas, bounds):
    """(host cell, function, Ref) for every single-range aggregate call."""
    out = []
    for cell, ftext in sorted(formulas.items()):
        sheet = cell[0]
        toks = tokenize(ftext)
        for t in toks:
            if t.kind != "ref" or ":" not in t.text:
                continue
            if t.func not in AGGREGATES:
                continue
            ref = resolve_ref(t.text, sheet, bounds)
            if ref.external:
                continue
            out.append((cell, t.func, ref))
    return out


def scan_short_ranges(aggs, grid, gap):
    hosts = defaultdict(set)
    for cell, fn, ref in aggs:
        hosts[ref.sheet].add(cell[1:])
    findings = []
    for cell, fn, ref in aggs:
        cells = grid.get(ref.sheet)
        if cells is None:
            continue
        vertical = ref.c1 == ref.c2
        horizontal = ref.r1 == ref.r2
        if vertical and not horizontal:
            c = ref.c1
            if (ref.r2, c) not in cells or (ref.r1, c) not in cells:
                continue
            missed, r = [], ref.r2 + 1
            blanks = 0
            while blanks <= gap and r < ref.r2 + 200:
                if (r, c) not in cells:
                    blanks += 1
                    r += 1
                    continue
                if (ref.sheet, r, c) == cell or (r, c) in hosts[ref.sheet]:
                    break
                if _is_total_label(_label_for_row(cells, r, c)):
                    break
                missed.append(addr(ref.sheet, r, c))
                blanks = 0
                r += 1
            if missed:
                findings.append({
                    "kind": "short_range",
                    "cell": cell,
                    "formula": None,
                    "severity": "high",
                    "note": (
                        f"{fn}({ref.raw}) ends at row {ref.r2}; "
                        f"{len(missed)} populated cell(s) below are excluded: "
                        + ", ".join(missed[:6])
                        + ("..." if len(missed) > 6 else "")
                    ),
                })
        elif horizontal and not vertical:
            r = ref.r1
            if (r, ref.c2) not in cells or (r, ref.c1) not in cells:
                continue
            missed, c = [], ref.c2 + 1
            blanks = 0
            while blanks <= gap and c < ref.c2 + 60:
                if (r, c) not in cells:
                    blanks += 1
                    c += 1
                    continue
                if (ref.sheet, r, c) == cell or (r, c) in hosts[ref.sheet]:
                    break
                missed.append(addr(ref.sheet, r, c))
                blanks = 0
                c += 1
            if missed:
                findings.append({
                    "kind": "short_range",
                    "cell": cell,
                    "formula": None,
                    "severity": "high",
                    "note": (
                        f"{fn}({ref.raw}) ends at column "
                        f"{get_column_letter(ref.c2)}; "
                        f"{len(missed)} populated cell(s) to the right are "
                        "excluded: " + ", ".join(missed[:6])
                        + ("..." if len(missed) > 6 else "")
                    ),
                })
    return findings


def scan_double_counts(aggs, grid):
    by_sheet = defaultdict(list)
    for cell, fn, ref in aggs:
        by_sheet[ref.sheet].append((cell, fn, ref))
    findings = []
    for sheet, items in by_sheet.items():
        for cell, fn, ref in items:
            if ref.c1 != ref.c2:
                continue
            for other, ofn, oref in items:
                if other is cell or other[0] != sheet:
                    continue
                if other[2] != ref.c1:
                    continue
                if not (ref.r1 <= other[1] <= ref.r2):
                    continue
                if oref.c1 == ref.c1 and ref.r1 <= oref.r1 and oref.r2 <= ref.r2:
                    findings.append({
                        "kind": "double_count",
                        "cell": cell,
                        "formula": None,
                        "severity": "high",
                        "note": (
                            f"{fn}({ref.raw}) spans {addr(*other)}, which is "
                            f"itself {ofn}({oref.raw}) over the same column"
                        ),
                    })
    return findings


def terminal_blocks(formulas, dependents, declared_outputs):
    terminals = {
        cell for cell in formulas
        if not dependents.get(cell) and addr(*cell) not in declared_outputs
    }
    by_sheet = defaultdict(set)
    for sheet, r, c in terminals:
        by_sheet[sheet].add((r, c))
    blocks = []
    for sheet, cells in by_sheet.items():
        remaining = set(cells)
        while remaining:
            seed = remaining.pop()
            block, stack = {seed}, [seed]
            while stack:
                r, c = stack.pop()
                for nr, nc in ((r + 1, c), (r - 1, c), (r, c + 1), (r, c - 1)):
                    if (nr, nc) in remaining:
                        remaining.discard((nr, nc))
                        block.add((nr, nc))
                        stack.append((nr, nc))
            rows = [r for r, _ in block]
            cols = [c for _, c in block]
            blocks.append({
                "sheet": sheet,
                "size": len(block),
                "range": (
                    f"{get_column_letter(min(cols))}{min(rows)}:"
                    f"{get_column_letter(max(cols))}{max(rows)}"
                ),
                "cells": sorted(block),
            })
    blocks.sort(key=lambda b: -b["size"])
    return blocks


SEVERITY_BASE = {"high": 6.0, "medium": 3.0, "low": 1.0}


def score_finding(f, reach):
    r = reach.get(f["cell"], 0) if f.get("cell") else 0
    f["reach"] = r
    f["score"] = round(SEVERITY_BASE[f["severity"]] * (1 + math.log10(1 + r)), 2)
    return f


# --------------------------------------------------------------------------
# report
# --------------------------------------------------------------------------


def md_escape(text):
    return str(text).replace("|", "\\|").replace("\n", " ")


def render(path, grid, formulas, bounds, findings, blocks, sheet_edges,
           externals, errors, crosslinks, unknown_sheets, defined, ext_files,
           cycles, reach, perturb, outputs, top):
    L = []
    n_cells = sum(len(c) for c in grid.values())
    stamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    L.append(f"# Model audit — {os.path.basename(path)}")
    L.append("")
    L.append(
        f"Sheets {len(grid)} · populated cells {n_cells} · formula cells "
        f"{len(formulas)} · constants {n_cells - len(formulas)} · "
        f"generated {stamp}"
    )
    L.append("")
    L.append(
        "Ranking below is **structural**: severity times downstream reach. "
        "This script never evaluates the workbook, so it cannot state how "
        "much of the answer any finding moves. Run the perturbation "
        "procedure to get that ranking."
    )
    L.append("")

    ranked = sorted(findings, key=lambda f: (-f["score"], addr(*f["cell"])))
    L.append("## Ranked findings")
    L.append("")
    if not ranked:
        L.append("No findings in the four defect classes.")
    else:
        L.append("| # | Score | Class | Cell | Reach | Finding |")
        L.append("|---|-------|-------|------|-------|---------|")
        for i, f in enumerate(ranked[:top], 1):
            L.append(
                f"| {i} | {f['score']} | {f['kind']} | `{addr(*f['cell'])}` | "
                f"{f['reach']} | {md_escape(f['note'])} |"
            )
        if len(ranked) > top:
            L.append("")
            L.append(f"{len(ranked) - top} further findings below the cut, listed by class.")
    L.append("")

    L.append("## Perturbation list — how to get the real ranking")
    L.append("")
    L.append(
        "Reach counts downstream formula cells, not currency. To rank by "
        "how much of the answer each item moves: open a copy, note the "
        "output values, change one cell at a time by the amount shown, "
        "record the change in each output, then undo before the next."
    )
    L.append("")
    if outputs:
        L.append("Declared outputs: " + ", ".join(f"`{o}`" for o in outputs))
    else:
        L.append(
            "No outputs declared. Re-run with `--outputs Sheet!B12,Sheet!B13` "
            "so this list can be filtered to cells that actually reach them."
        )
    L.append("")
    L.append("| # | Cell | Current content | Reach | Perturb by |")
    L.append("|---|------|-----------------|-------|------------|")
    for i, (cell, content, r, how) in enumerate(perturb, 1):
        L.append(
            f"| {i} | `{addr(*cell)}` | `{md_escape(content)}` | {r} | {how} |"
        )
    L.append("")

    def section(title, kind, cols, rowfn):
        items = [f for f in findings if f["kind"] == kind]
        L.append(f"## {title} ({len(items)})")
        L.append("")
        if not items:
            L.append("None found.")
            L.append("")
            return
        L.append("| " + " | ".join(cols) + " |")
        L.append("|" + "|".join("---" for _ in cols) + "|")
        for f in sorted(items, key=lambda x: (-x["score"], addr(*x["cell"]))):
            L.append("| " + " | ".join(rowfn(f)) + " |")
        L.append("")

    section(
        "Hardcoded constants inside formulas", "hardcode",
        ["Cell", "Literal", "Kind", "Severity", "Formula"],
        lambda f: [
            f"`{addr(*f['cell'])}`", f"`{f['literal']}`", f["label"],
            f["severity"], f"`{md_escape(f['formula'])}`",
        ],
    )
    section(
        "Formulas that break their row", "row_break",
        ["Cell", "This cell", "Row majority", "Compare with", "Note"],
        lambda f: [
            f"`{addr(*f['cell'])}`", f"`{md_escape(f['formula'])}`",
            f"`{md_escape(f['majority'])}` ({f['majority_share']})",
            f"`{f['example']}`", md_escape(f["note"]),
        ],
    )
    section(
        "Ranges that stop short", "short_range",
        ["Cell", "Formula", "Finding"],
        lambda f: [
            f"`{addr(*f['cell'])}`",
            f"`{md_escape(formulas.get(f['cell'], ''))}`",
            md_escape(f["note"]),
        ],
    )
    section(
        "Ranges that double count", "double_count",
        ["Cell", "Formula", "Finding"],
        lambda f: [
            f"`{addr(*f['cell'])}`",
            f"`{md_escape(formulas.get(f['cell'], ''))}`",
            md_escape(f["note"]),
        ],
    )

    L.append(f"## Links and error remnants")
    L.append("")
    L.append(f"External workbook references: {len(externals)}")
    if externals:
        L.append("")
        L.append("| Cell | Reference | Formula |")
        L.append("|------|-----------|---------|")
        for cell, ftext, ref in externals[:40]:
            L.append(f"| `{addr(*cell)}` | `{md_escape(ref)}` | `{md_escape(ftext)}` |")
    if ext_files:
        L.append("")
        L.append("Linked workbook targets recorded in the file: "
                 + ", ".join(f"`{t}`" for t in ext_files))
    L.append("")
    L.append(f"Error remnants (`#REF!` and friends): {len(errors)}")
    if errors:
        L.append("")
        L.append("| Cell | Error | Formula |")
        L.append("|------|-------|---------|")
        for cell, ftext, tok in errors[:40]:
            L.append(f"| `{addr(*cell)}` | `{tok}` | `{md_escape(ftext)}` |")
    bad_names = [(n, v) for n, v in defined if "#REF" in str(v)]
    if bad_names:
        L.append("")
        L.append("Defined names pointing at deleted cells: "
                 + ", ".join(f"`{n}` → `{v}`" for n, v in bad_names))
    if unknown_sheets:
        L.append("")
        L.append("References to sheets not present in this file: "
                 + ", ".join(f"`{s}`" for s in sorted(unknown_sheets)))
    L.append("")
    L.append(f"Cross-sheet references: {len(crosslinks)}")
    L.append("")

    L.append("## Dependency map")
    L.append("")
    L.append("| From sheet | To sheet | References |")
    L.append("|------------|----------|------------|")
    for (a, b), n in sorted(sheet_edges.items(), key=lambda kv: -kv[1]):
        L.append(f"| {a} | {b} | {n} |")
    if not sheet_edges:
        L.append("| — | — | 0 |")
    L.append("")
    inbound = defaultdict(int)
    for (a, b), n in sheet_edges.items():
        inbound[b] += n
    orphans = [s for s in grid if inbound.get(s, 0) == 0 and len(grid) > 1]
    L.append("Sheets no other sheet reads: "
             + (", ".join(f"`{s}`" for s in orphans) if orphans else "none"))
    L.append("")
    if cycles:
        L.append("Circular references detected:")
        L.append("")
        for cyc in cycles[:5]:
            L.append("- " + " → ".join(f"`{addr(*c)}`" for c in cyc))
        L.append("")
    else:
        L.append("No circular references.")
        L.append("")

    L.append("## Calculations nothing depends on")
    L.append("")
    L.append(
        "A formula cell with no dependents is either a reported output or a "
        "dead calculation. Static analysis cannot separate the two; declared "
        "outputs are excluded, everything else is listed for the reader to "
        "judge. Blocks are contiguous groups, largest first."
    )
    L.append("")
    if not blocks:
        L.append("Every formula cell feeds another.")
    else:
        L.append("| Sheet | Block | Cells | First cell formula |")
        L.append("|-------|-------|-------|--------------------|")
        for b in blocks[:25]:
            first = b["cells"][0]
            L.append(
                f"| {b['sheet']} | `{b['range']}` | {b['size']} | "
                f"`{md_escape(formulas.get((b['sheet'],) + first, ''))}` |"
            )
    L.append("")
    L.append("## What this audit cannot tell you")
    L.append("")
    L.append(
        "- Which finding moves the answer most. Reach is a count of cells, "
        "not of currency; a widely-referenced input can still be worth "
        "little. Perturb to find out."
    )
    L.append(
        "- Whether a hardcoded literal is wrong. It reports that a number "
        "lives inside a formula where nobody will find it."
    )
    L.append(
        "- Whether a formula that breaks its row is a mistake or a "
        "deliberate exception. It reports that the row is not uniform."
    )
    L.append(
        "- Anything about values. The workbook was read with "
        "`data_only=False`, so cached results were never consulted."
    )
    L.append("")
    return "\n".join(L)


# --------------------------------------------------------------------------


def input_candidates(grid, formulas, dependents, limit=2000):
    """Numeric constants that at least one formula reads, most-read first."""
    out = []
    for sheet, cells in grid.items():
        for (r, c), v in cells.items():
            key = (sheet, r, c)
            if key in formulas or isinstance(v, bool):
                continue
            if not isinstance(v, (int, float)):
                continue
            n = len(dependents.get(key, ()))
            if n:
                out.append((n, key, v))
    out.sort(key=lambda t: (-t[0], addr(*t[1])))
    return out[:limit]


def build_perturbation_list(findings, grid, formulas, inputs, reach, limit):
    seen, rows = set(), []
    for f in sorted(findings, key=lambda x: -x["score"]):
        cell = f["cell"]
        if cell in seen:
            continue
        seen.add(cell)
        content = formulas.get(cell) or str(grid[cell[0]].get(cell[1:], ""))
        how = ("replace the flagged literal with an input cell, then move it "
               "±10%") if f["kind"] == "hardcode" else "±10% of its value"
        rows.append((cell, content, reach.get(cell, 0), how))
        if len(rows) >= limit:
            break
    ranked_inputs = sorted(
        ((reach.get(key, 0), key, v) for _, key, v in inputs),
        key=lambda t: (-t[0], addr(*t[1])),
    )
    for r, key, v in ranked_inputs[:limit]:
        if key in seen:
            continue
        seen.add(key)
        rows.append((key, v, r, "±10% of its value"))
    rows.sort(key=lambda t: -t[2])
    return rows


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("workbook")
    ap.add_argument("--out", help="write the markdown report here")
    ap.add_argument("--json", dest="json_out", help="also write findings as JSON")
    ap.add_argument("--outputs", default="",
                    help="comma-separated output cells, e.g. Summary!B12")
    ap.add_argument("--top", type=int, default=25, help="rows in the ranked table")
    ap.add_argument("--min-run", type=int, default=4,
                    help="shortest run of adjacent cells tested for row consistency")
    ap.add_argument("--share", type=float, default=0.6,
                    help="fraction of a run that must share a pattern")
    ap.add_argument("--gap", type=int, default=1,
                    help="blank cells tolerated when extending a summed block")
    ap.add_argument("--max-cells", type=int, default=400000)
    ap.add_argument("--strict", action="store_true",
                    help="exit 2 when any high-severity finding exists")
    args = ap.parse_args(argv)

    grid, formulas, bounds, defined, ext_files = read_workbook(
        args.workbook, args.max_cells)
    outputs = [o.strip() for o in args.outputs.split(",") if o.strip()]

    (precedents, dependents, sheet_edges, externals, errors,
     crosslinks, unknown_sheets) = build_graph(
        formulas, bounds, grid, list(grid))
    cycles = find_cycles(precedents)

    aggs = collect_aggregates(formulas, bounds)
    findings = []
    findings += scan_hardcodes(formulas)
    findings += scan_row_consistency(grid, formulas, args.min_run, args.share)
    findings += scan_short_ranges(aggs, grid, args.gap)
    findings += scan_double_counts(aggs, grid)

    inputs = input_candidates(grid, formulas, dependents)
    targets = {f["cell"] for f in findings} | {key for _, key, _ in inputs}
    reach = transitive_reach(dependents, targets)
    for f in findings:
        score_finding(f, reach)

    blocks = terminal_blocks(formulas, dependents, set(outputs))
    perturb = build_perturbation_list(
        findings, grid, formulas, inputs, reach, 10)

    report = render(
        args.workbook, grid, formulas, bounds, findings, blocks, sheet_edges,
        externals, errors, crosslinks, unknown_sheets, defined, ext_files,
        cycles, reach, perturb, outputs, args.top)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(report + "\n")
        sys.stderr.write(f"report written to {args.out}\n")
    else:
        sys.stdout.write(report + "\n")

    if args.json_out:
        payload = {
            "workbook": os.path.abspath(args.workbook),
            "sheets": list(grid),
            "formula_cells": len(formulas),
            "findings": [
                {**{k: v for k, v in f.items() if k != "cell"},
                 "address": addr(*f["cell"])}
                for f in sorted(findings, key=lambda x: -x["score"])
            ],
            "dead_blocks": [
                {k: v for k, v in b.items() if k != "cells"} for b in blocks
            ],
            "sheet_edges": [
                {"from": a, "to": b, "n": n} for (a, b), n in sheet_edges.items()
            ],
            "externals": [addr(*c) for c, _, _ in externals],
            "errors": [addr(*c) for c, _, _ in errors],
        }
        with open(args.json_out, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2, ensure_ascii=False)
        sys.stderr.write(f"json written to {args.json_out}\n")

    if args.strict and any(f["severity"] == "high" for f in findings):
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
