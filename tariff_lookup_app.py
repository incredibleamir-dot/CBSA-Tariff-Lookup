import re
import sqlite3
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from ttkbootstrap.widgets.dateentry import DateEntry
from ttkbootstrap.widgets.tableview import Tableview

if getattr(sys, "frozen", False):  # running as PyInstaller exe
    DB_PATH = Path(sys.executable).resolve().parent / "tariff_app.db"
else:
    DB_PATH = Path(__file__).resolve().parent / "tariff_app.db"

APP_VERSION = "1.0.0"
APP_AUTHOR = "Amir Arshad"
APP_GITHUB = "https://github.com/incredibleamir-dot"
APP_REPO = "https://github.com/incredibleamir-dot/CBSA-Tariff-Lookup"

TREATMENT_COLS = ["MFN", "General Tariff", "AUT", "NZT", "CCCT", "LDCT", "GPT",
                  "UST", "MXT", "CIAT", "CT", "CRT", "IT", "NT", "SLT", "PT",
                  "COLT", "JT", "PAT", "HNT", "KRT", "CEUT", "UAT", "CPTPT",
                  "UKT", "CPUKT"]

COUNTRY_ALIASES = {
    "usa": "United States of America", "us": "United States of America",
    "united states": "United States of America", "america": "United States of America",
    "uk": "United Kingdom", "u.k.": "United Kingdom", "britain": "United Kingdom",
    "south korea": "South Korea", "korea": "South Korea",
    "viet nam": "Vietnam", "uae": "United Arab Emirates",
}

US_ORIGINS = {"united states of america", "united states", "usa", "us", "america", "puerto rico"}

# Tariffs grantable by law but not encoded in CBSA's country-list "Other" column:
# UK acceded to CPTPP — UK goods can claim CPUKT as well as UKT (Act s.52.82).
EXTRA_TREATMENTS = {"United Kingdom": ["CPUKT"]}

try:
    from tabulate import tabulate
except ImportError:
    tabulate = None


def norm_code(s: str) -> str:
    return re.sub(r"\D", "", s or "")


def parse_advalorem(rate) -> float | None:
    """'Free' -> 0.0, '6.5%' -> 6.5, specific/None -> None."""
    if rate is None:
        return None
    t = str(rate).strip()
    if not t or t.upper() in ("NONE", "N/A", "NA", "NULL"):
        return None
    if t.lower() == "free":
        return 0.0
    m = re.search(r"([\d]+(?:\.[\d]+)?)\s*%", t)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            return None
    return None


def get_breadcrumb(con, item_norm: str, rev: str | None = None) -> list[str]:
    """Parent-heading descriptions for an HS code, chapter-first.

    e.g. 0402.10.10 -> ["Milk and cream, concentrated ...", "In powder ..."].
    """
    rev = rev or active_rev(con)
    out = []
    for _tariff, desc in con.execute(
            "SELECT TARIFF, DESC1 FROM tphs WHERE rev=? AND ? LIKE TARIFF_NORM || '%' "
            "AND length(TARIFF_NORM) < length(?) ORDER BY length(TARIFF_NORM)",
            (rev, item_norm, item_norm)):
        if desc and str(desc) != "None":
            out.append(str(desc).strip())
    return out


# rate-unit ->canonical UOM, and mass conversions to kg
_RATE_UNIT_MAP = {"kg": "KGM", "kgs": "KGM", "kilogram": "KGM", "kilograms": "KGM",
                 "tonne": "TNE", "tonnes": "TNE", "t": "TNE",
                 "litre": "LTR", "litres": "LTR", "l": "LTR", "liter": "LTR",
                 "each": "NMB", "number": "NMB", "dozen": "DZN", "pair": "PAR",
                 "metre": "MTR", "meter": "MTR", "m": "MTR", "gram": "GRM", "g": "GRM"}
_MASS_KG = {"GRM": 0.001, "KGM": 1.0, "TNE": 1000.0}


def parse_specific(rate) -> tuple[float, str | None] | None:
    """Parse 'X ¢/kg', '$X/tonne', 'X ¢each' etc -> (CAD per unit, UOM or None).

    Bare amounts like '1.9 ¢' return (cad, None) — caller assumes the row's UOM.
    Returns None for ad-valorem, compound or unparseable rates.
    """
    if rate is None:
        return None
    t = str(rate).strip().lower().replace("cents", "¢")
    m = re.match(r"^(\$)?\s*([\d]*\.?[\d]+)\s*(¢|c|\$)?\s*/?\s*([a-z]*)\s*$", t)
    sym = (m.group(1) or m.group(3)) if m else None
    if not m or not sym:
        return None
    amt, unit = float(m.group(2)), m.group(4)
    if not unit:
        return (amt / 100.0 if sym in ("¢", "c") else amt, None)
    uom = _RATE_UNIT_MAP.get(unit)
    if uom is None:
        return None
    cad = amt / 100.0 if sym in ("¢", "c") else amt
    return (cad, uom)


def convert_qty(qty: float, from_uom: str, to_uom: str) -> float | None:
    """Convert qty between UOMs. Mass family converts freely; else exact match only."""
    if from_uom == to_uom:
        return qty
    if from_uom in _MASS_KG and to_uom in _MASS_KG:
        return qty * _MASS_KG[from_uom] / _MASS_KG[to_uom]
    return None


PERIOD_WINDOWS = {
    "up_to_2025-08-31": (None, date(2025, 8, 31)),
    "2025-09-01_to_2026-09-07": (date(2025, 9, 1), date(2026, 9, 7)),
    "2026-09-08+": (date(2026, 9, 8), None),
}


def surtax_applies(effective: str, period: str, import_date: date | None) -> bool:
    """True if a surtax row was in force on import_date (None = show all)."""
    if import_date is None:
        return True
    try:
        eff = date.fromisoformat(str(effective))
    except ValueError:
        return True
    if import_date < eff:
        return False
    start, end = PERIOD_WINDOWS.get(period, (None, None))
    if start and import_date < start:
        return False
    if end and import_date > end:
        return False
    return True


def _pdf_text(t) -> str:
    """Sanitize to WinAnsi-safe text for Helvetica (reportlab)."""
    return (str(t).replace("★", "*").replace("›", ">").replace("¢", "cents")
            .replace("—", "-").replace("–", "-").replace("•", "-").replace("⚠", "!"))


def export_pdf(path: str, meta: dict, duties: list[dict],
               surtax: list[dict], summary: list[dict]) -> str:
    """Modern styled PDF report: meta + duties + surtax + summary tables."""
    import datetime as _dt
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, HRFlowable)

    NAVY, ACCENT, LIGHT, GREEN_BG = (colors.HexColor("#1F4E79"),
                                     colors.HexColor("#2E75B6"),
                                     colors.HexColor("#DDEBF7"),
                                     colors.HexColor("#E2EFDA"))
    RED = colors.HexColor("#C00000")
    styles = getSampleStyleSheet()
    title = ParagraphStyle("Title2", parent=styles["Title"], textColor=NAVY,
                           fontSize=20, spaceAfter=2)
    sub = ParagraphStyle("Sub", parent=styles["Normal"], textColor=ACCENT,
                         fontSize=10, spaceAfter=8)
    cell = ParagraphStyle("Cell", parent=styles["Normal"], fontSize=8, leading=10)
    cell_h = ParagraphStyle("CellH", parent=cell, textColor=colors.white,
                            fontName="Helvetica-Bold")
    small = ParagraphStyle("Small", parent=styles["Normal"], fontSize=8,
                           textColor=colors.grey)

    doc = SimpleDocTemplate(path, pagesize=landscape(A4),
                            leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=44,
                            title="Canada Tariff Report", author="Tariff Lookup App")
    story = [Paragraph("Canada Customs Tariff Report", title),
             Paragraph(f"Origin {_pdf_text(meta.get('Origin', ''))}  •  "
                       f"Generated {_dt.datetime.now():%Y-%m-%d %H:%M}", sub),
             HRFlowable(width="100%", color=ACCENT, thickness=1.5), Spacer(1, 10)]

    meta_rows = [[Paragraph(f"<b>{_pdf_text(k)}</b>", cell),
                   Paragraph(_pdf_text(v), cell)] for k, v in meta.items()]
    story += [Table(meta_rows, colWidths=[130, 560],
                    style=TableStyle([("BACKGROUND", (0, 0), (0, -1), LIGHT),
                                       ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                                       ("VALIGN", (0, 0), (-1, -1), "TOP")])),
              Spacer(1, 12)]

    def styled_table(headers, rows, widths, header_bg, best_idx=()):
        body = [[Paragraph(f"<b>{_pdf_text(h)}</b>", cell_h) for h in headers]]
        for r in rows:
            body.append([Paragraph(_pdf_text(c), cell) for c in r])
        ts = [("BACKGROUND", (0, 0), (-1, 0), header_bg),
              ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
              ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
              ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
              ("VALIGN", (0, 0), (-1, -1), "TOP")]
        for i in best_idx:
            ts.append(("BACKGROUND", (0, i + 1), (-1, i + 1), GREEN_BG))
        return Table(body, colWidths=widths, repeatRows=1,
                     style=TableStyle(ts))

    story.append(Paragraph("<b>Duties by treatment</b> (* = lowest option)", cell))
    dcols = ["HS", "Description", "Treatment", "Rate", "Duty", "Best"]
    drows = [[d.get("HS", ""), d.get("Description", ""), d.get("Treatment", ""),
              d.get("Rate", ""), d.get("Duty", ""), d.get("Best", "")] for d in duties]
    best_idx = tuple(i for i, d in enumerate(duties) if d.get("Best") == "★")
    story.append(styled_table(dcols, drows, [80, 290, 85, 75, 105, 40], NAVY, best_idx))
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>US reciprocal surtax</b>", cell))
    if surtax:
        srows = [[s.get("HS", ""), s.get("Period", ""), s.get("Tariff item", ""),
                  s.get("Effective", ""), s.get("Surtax", "")] for s in surtax]
        story.append(styled_table(["HS", "Period", "Tariff item", "Effective", "Surtax"],
                                  srows, [80, 210, 100, 100, 90], RED))
    else:
        story.append(Paragraph("No surtax rows for this selection.", cell))
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>Summary — one line per HS</b>", cell))
    scol = ["HS", "Description", "Lowest", "Rate", "Duty", "GST", "Surtax"]
    grows = [[g.get("HS", ""), g.get("Description", ""), g.get("Lowest", ""),
              g.get("Rate", ""), g.get("Duty", ""), g.get("GST", ""),
              g.get("Surtax", "")] for g in summary]
    story.append(styled_table(scol, grows, [75, 200, 70, 60, 70, 70, 125], ACCENT))
    story.append(Spacer(1, 12))
    story.append(Paragraph("For information only — verify classifications, rates and surtax "
                           "periods with the CBSA Customs Tariff and Canada Gazette orders.",
                           small))

    def _foot(canvas, _doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.grey)
        canvas.drawString(36, 24, "Tariff Lookup App — T2026-2 + US surtax (all periods)")
        canvas.drawRightString(landscape(A4)[0] - 36, 24, f"Page {_doc.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=_foot, onLaterPages=_foot)
    return path


def export_results(path: str, meta: dict, duties: list[dict],
                   surtax: list[dict], summary: list[dict]) -> str:
    """Write summary/duties/surtax/info to .xlsx (4 sheets), .csv (sections) or .pdf."""
    if path.lower().endswith(".pdf"):
        return export_pdf(path, meta, duties, surtax, summary)
    import pandas as pd
    low = path.lower()
    if low.endswith(".csv"):
        import csv as _csv
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = _csv.writer(f)
            w.writerow(["TARIFF REPORT"])
            for k, v in meta.items():
                w.writerow([k, v])
            w.writerow([])
            w.writerow(["SUMMARY"])
            if summary:
                w.writerow(list(summary[0].keys()))
                w.writerows([list(r.values()) for r in summary])
            w.writerow([])
            w.writerow(["DUTIES"])
            if duties:
                w.writerow(list(duties[0].keys()))
                w.writerows([list(r.values()) for r in duties])
            w.writerow([])
            w.writerow(["SURTAX"])
            if surtax:
                w.writerow(list(surtax[0].keys()))
                w.writerows([list(r.values()) for r in surtax])
    else:
        with pd.ExcelWriter(path) as xw:
            pd.DataFrame(summary).to_excel(xw, sheet_name="Summary", index=False)
            pd.DataFrame(duties).to_excel(xw, sheet_name="Duties", index=False)
            pd.DataFrame(surtax).to_excel(xw, sheet_name="Surtax", index=False)
            pd.DataFrame([{"Item": k, "Value": v} for k, v in meta.items()]).to_excel(
                xw, sheet_name="Info", index=False)
    return path


def active_rev(con) -> str:
    """Active tariff revision (db_meta), fallback to bundled T2026-2."""
    try:
        r = con.execute("SELECT value FROM db_meta WHERE key='active_rev'").fetchone()
        return r[0] if r else "T2026-2"
    except Exception:
        return "T2026-2"


def installed_revs(con) -> list[str]:
    try:
        return [r[0] for r in con.execute(
            "SELECT DISTINCT rev FROM tphs ORDER BY rev")]
    except Exception:
        return ["T2026-2"]


def rev_info(con, rev: str) -> dict:
    out = {"rev": rev, "rows": 0, "effective": "", "source": ""}
    try:
        out["rows"] = con.execute("SELECT COUNT(*) FROM tphs WHERE rev=?",
                                  (rev,)).fetchone()[0]
        for k in ("effective", "source"):
            r = con.execute("SELECT value FROM db_meta WHERE key=?",
                            (f"rev:{rev}:{k}",)).fetchone()
            if r:
                out[k] = r[0]
    except Exception:
        pass
    return out


def diff_revs(con, rev_a: str, rev_b: str):
    """A-vs-B tariff changes: [change, tariff, description, detail] rows.

    change in ADDED / REMOVED / RATE CHANGE / DESCRIPTION CHANGE.
    """
    cols = ["TARIFF_NORM", "TARIFF", "DESC1", "DESC2", "DESC3"] + TREATMENT_COLS

    def snap(rev):
        d = {}
        qcols = ", ".join(f'"{c}"' for c in cols)
        for r in con.execute(f"SELECT {qcols} FROM tphs WHERE rev=?", (rev,)):
            d[r[0]] = {c: ("" if v is None else str(v).strip()) for c, v in zip(cols, r)}
        return d

    A, B = snap(rev_a), snap(rev_b)
    out = []
    for norm in sorted(set(A) | set(B), key=lambda n: (len(n), n)):
        if norm not in A:
            b = B[norm]
            out.append(["ADDED", b["TARIFF"], b["DESC1"],
                        "new in " + rev_b + "; " + "; ".join(
                            f"{c}={b[c]}" for c in TREATMENT_COLS if b[c])])
        elif norm not in B:
            a = A[norm]
            out.append(["REMOVED", a["TARIFF"], a["DESC1"], "gone in " + rev_b])
        else:
            a, b = A[norm], B[norm]
            rd = [c for c in TREATMENT_COLS if a[c] != b[c]]
            dd = [c for c in ("DESC1", "DESC2", "DESC3") if a[c] != b[c]]
            if rd:
                out.append(["RATE CHANGE", b["TARIFF"], b["DESC1"],
                            "; ".join(f"{c}: {a[c] or '—'} → {b[c] or '—'}" for c in rd)])
            elif dd:
                out.append(["DESCRIPTION CHANGE", b["TARIFF"], b["DESC1"],
                            "; ".join(f"{c} changed" for c in dd)])
    return out


def get_connection():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def country_rev_for(con, rev: str | None = None) -> str:
    """Country-list revision paired with a tariff revision (db_meta mapping)."""
    rev = rev or active_rev(con)
    try:
        r = con.execute("SELECT value FROM db_meta WHERE key=?",
                        (f"rev:{rev}:country_rev",)).fetchone()
        if r:
            return r[0]
        r = con.execute("SELECT rev FROM country_map ORDER BY rev DESC LIMIT 1").fetchone()
        return r[0] if r else "T2026"
    except Exception:
        return "T2026"


def load_countries(con, rev: str | None = None) -> list[str]:
    crev = country_rev_for(con, rev)
    return [r[0] for r in con.execute(
        "SELECT country FROM country_map WHERE rev=? ORDER BY country", (crev,))]


def resolve_country(con, text: str, rev: str | None = None) -> str | None:
    crev = country_rev_for(con, rev)
    t = (text or "").strip()
    if not t:
        return None
    alias = COUNTRY_ALIASES.get(t.lower())
    if alias:
        return alias
    row = con.execute("SELECT country FROM country_map WHERE rev=? AND lower(country)=?",
                      (crev, t.lower())).fetchone()
    if row:
        return row[0]
    row = con.execute("SELECT country FROM country_map WHERE rev=? AND lower(country) LIKE ?",
                      (crev, f"%{t.lower()}%")).fetchone()
    return row[0] if row else None


def eligible_treatments(con, country: str, rev: str | None = None) -> list[str]:
    """Ordered eligible duty columns for an origin country.

    MFN is only included for MFN beneficiaries (e.g. Russia and Belarus
    had MFN withdrawn — General Tariff only).
    """
    rev = rev or active_rev(con)
    crev = country_rev_for(con, rev)
    row = con.execute("SELECT mfn, gpt, ldct, other FROM country_map WHERE rev=? AND country=?",
                      (crev, country)).fetchone()
    out = []
    if row is None:
        out = ["MFN", "General Tariff"]
    else:
        if (row[0] or "").lower() == "yes":
            out.append("MFN")
        out.append("General Tariff")
        if (row[1] or "").lower() == "yes" and "GPT" not in out:
            out.append("GPT")
        if (row[2] or "").lower() == "yes" and "LDCT" not in out:
            out.append("LDCT")
        other = (row[3] or "").strip()
        if other:
            for code in re.split(r"[,\s]+", other):
                code = code.strip()
                if code and code in TREATMENT_COLS and code not in out:
                    out.append(code)
    for code in EXTRA_TREATMENTS.get(country, []):
        if code not in out:
            out.append(code)
    return out


def find_tariff_row(con, query: str, rev: str | None = None):
    """Best TPHS row for a code: exact digits match, latest EFF_DATE on ties."""
    rev = rev or active_rev(con)
    n = norm_code(query)
    if len(n) < 4:
        return None
    rows = con.execute("SELECT * FROM tphs WHERE rev=? AND TARIFF_NORM=? ORDER BY EFF_DATE DESC",
                       (rev, n)).fetchall()
    if rows:
        return rows[0]
    # fallback: prefix — user typed short code, take shortest longer match, latest date
    rows = con.execute("SELECT * FROM tphs WHERE rev=? AND TARIFF_NORM LIKE ? "
                       "ORDER BY length(TARIFF_NORM), EFF_DATE DESC LIMIT 5",
                       (rev, n + "%")).fetchall()
    return rows[0] if rows else None


def search_tariff(con, text: str, limit: int = 100, rev: str | None = None):
    rev = rev or active_rev(con)
    t = (text or "").strip()
    if not t:
        return []
    if re.search(r"\d", t):
        n = norm_code(t)
        if len(n) >= 3:
            return con.execute(
                "SELECT TARIFF, DESC1, UOM FROM tphs "
                "WHERE rev=? AND TARIFF_NORM LIKE ? ORDER BY length(TARIFF_NORM), TARIFF "
                "LIMIT ?", (rev, f"%{n}%", limit)).fetchall()
        return con.execute(
            "SELECT TARIFF, DESC1, UOM FROM tphs WHERE rev=? AND TARIFF LIKE ? LIMIT ?",
            (rev, f"%{t}%", limit)).fetchall()
    like = f"%{t}%"
    return con.execute(
        "SELECT TARIFF, DESC1, UOM FROM tphs "
        "WHERE rev=? AND DESC_FULL LIKE ? ORDER BY TARIFF LIMIT ?", (rev, like, limit)).fetchall()


def get_counter_rows(con, item_norm: str):
    return con.execute("SELECT tariff_item, hs_chapter, hs_heading, description, "
                       "effective_date, rate, period FROM counter_tariffs "
                       "WHERE ITEM_NORM=? ORDER BY effective_date", (item_norm,)).fetchall()


def _ascii_table(headers: list[str], rows: list[list]) -> str:
    """Grid-style ASCII table via tabulate, with a fixed-width fallback."""
    if tabulate is not None:
        return tabulate(rows, headers=headers, tablefmt="grid")
    widths = [len(h) for h in headers]
    for r in rows:
        for i, cell in enumerate(r):
            widths[i] = max(widths[i], len(str(cell)))
    bar = "+" + "+".join("-" * (w + 2) for w in widths) + "+"
    fmt = "|" + "|".join(" {:<%d} " % w for w in widths) + "|"
    out = [bar, fmt.format(*headers), bar.replace("-", "=")]
    for r in rows:
        out.append(fmt.format(*[str(c) for c in r]))
    out.append(bar)
    return "\n".join(out)


def origin_note(con, country: str, rev: str | None = None) -> str | None:
    """Warning when MFN was withdrawn (Russia, Belarus — General Tariff only)."""
    if "MFN" not in eligible_treatments(con, country, rev):
        return (f"NOTE: Canada has withdrawn MFN treatment from {country} — "
                "only the General Tariff applies.")
    return None


def _row_has_rates(row, treat) -> bool:
    d = dict(row)
    return any(str(d.get(c) or "").strip() not in ("", "None") for c in treat)


def _nearest_rated_ancestor(con, query_norm: str, treat, rev: str):
    """Nearest parent line carrying rates (statistical suffixes inherit it)."""
    q = ("SELECT * FROM tphs WHERE rev=? AND ? LIKE TARIFF_NORM || '%' "
         "AND length(TARIFF_NORM) < ? ORDER BY length(TARIFF_NORM) DESC LIMIT 10")
    for r in con.execute(q, (rev, query_norm, len(query_norm))):
        if _row_has_rates(r, treat):
            return r
    return None


def _first_rated_descendant(con, query_norm: str, treat, rev: str):
    """Closest child line carrying rates (for rateless header/fragment rows)."""
    q = ("SELECT * FROM tphs WHERE rev=? AND TARIFF_NORM LIKE ? "
         "AND length(TARIFF_NORM) > ? ORDER BY length(TARIFF_NORM), EFF_DATE DESC LIMIT 25")
    for r in con.execute(q, (rev, query_norm + "%", len(query_norm))):
        if _row_has_rates(r, treat):
            return r
    return None


def get_duty_data(con, tariff_query: str, country: str, value: float | None = 1000.0,
                  qty: float | None = None, qty_unit: str | None = None,
                  rev: str | None = None):
    """Structured duty data for one HS code. Returns None if HS not found.

    rows: [treatment, rate_str, duty_str, adv_or_None, duty_amt_or_None].
    best: (treatment, adv_or_None, duty_amt) with lowest computable duty, else None.
    """
    row = find_tariff_row(con, tariff_query, rev)
    if row is None:
        return None
    rev = rev or active_rev(con)
    if value is not None and value <= 0:
        value = None
    if qty is not None and qty <= 0:
        qty = None
    d = dict(row)
    treat = eligible_treatments(con, country, rev)
    resolved_from = None
    inherited_from = None
    if not _row_has_rates(row, treat):
        alt = _first_rated_descendant(con, norm_code(tariff_query), treat, rev)
        if alt is not None:
            resolved_from = row["TARIFF"]
            row = alt
            d = dict(row)
    if not _row_has_rates(row, treat):
        anc = _nearest_rated_ancestor(con, norm_code(tariff_query), treat, rev)
        if anc is not None:
            resolved_from = row["TARIFF"]
            inherited_from = anc["TARIFF"]
            row = anc
            d = dict(row)
    show_money = value is not None and value > 0
    own = " ".join(x for x in [d.get("DESC1"), d.get("DESC2"), d.get("DESC3")]
                   if x and str(x) != "None").strip()
    crumb = get_breadcrumb(con, norm_code(d.get("TARIFF")), rev)
    context = " › ".join(crumb + ([own] if own else ""))
    uom = (d.get("UOM") or "").strip() or "-"
    rows: list[list] = []
    best = None
    for col in treat:
        rate = d.get(col)
        if rate is None or str(rate).strip() in ("", "None"):
            continue
        adv = parse_advalorem(rate)
        duty_amt = None
        if adv is not None:
            if adv == 0:
                duty_amt = 0.0
                duty_str = "$0.00"
            elif show_money:
                duty_amt = value * adv / 100.0
                duty_str = f"${duty_amt:,.2f}"
            else:
                duty_str = "-"
        else:
            spec = parse_specific(rate)
            if spec and qty:
                per_unit, rate_uom = spec
                assumed = rate_uom or (qty_unit or uom)
                conv = convert_qty(qty, (qty_unit or uom), assumed)
                if conv is not None:
                    duty_amt = per_unit * conv
                    tag = "" if rate_uom else " (per-unit assumed)"
                    duty_str = f"${duty_amt:,.2f} ({qty:g} {qty_unit or uom}){tag}"
                else:
                    duty_str = f"{rate} (qty unit mismatch)"
            else:
                duty_str = f"{rate}" + ("" if qty else " (enter qty to compute)")
        if duty_amt is not None and (best is None or duty_amt < best[2]):
            best = (col, adv, duty_amt)
        rows.append([col, str(rate), duty_str, adv, duty_amt])
    has_money = show_money or any(r[4] is not None for r in rows)
    gst_amt = None
    if show_money:
        total_duty = best[2] if best else 0.0
        base = value + total_duty
        gst_amt = base * 0.05
        gst = (f"GST 5% on duty-paid value = ${gst_amt:,.2f} "
               f"(duty-paid ${base:,.2f}). PST/QST/HST not levied at border.")
    else:
        gst = ("GST 5% applies on duty-paid value (enter a value and/or quantity "
               "to compute $ duty and GST).")
    return {"hs": d.get("TARIFF"), "eff": d.get("EFF_DATE"), "desc": own,
            "context": context or own, "uom": uom, "country": country,
            "treatments": treat, "rows": rows, "best": best, "gst": gst,
            "gst_amt": gst_amt, "show_money": has_money,
            "norm": norm_code(d.get("TARIFF")), "note": origin_note(con, country, rev),
            "resolved_from": resolved_from, "inherited_from": inherited_from}


def build_summary_table(items):
    """Per-HS summary rows for tables/exports.

    items: [(duty_data_or_None, query_code, surtax_rows)].
    Returns [dict(HS, Description, Lowest, Rate, Duty, GST, Surtax)].
    """
    out = []
    for data, query, srows in items:
        if data is None:
            out.append({"HS": query, "Description": "NOT FOUND", "Lowest": "—",
                        "Rate": "—", "Duty": "—", "GST": "—", "Surtax": "—"})
            continue
        best = data["best"]
        if best:
            brow = next((r for r in data["rows"] if r[0] == best[0]), None)
            out.append({"HS": data["hs"],
                        "Description": data["context"][:70],
                        "Lowest": best[0],
                        "Rate": brow[1] if brow else "—",
                        "Duty": f"${best[2]:,.2f}",
                        "GST": f"${data['gst_amt']:,.2f}"
                        if data.get("gst_amt") is not None else "—",
                        "Surtax": "; ".join(f"{s[3]} ({s[2]})" for s in srows) or "—"})
        else:
            out.append({"HS": data["hs"], "Description": data["context"][:70],
                        "Lowest": "no computable duty", "Rate": "—", "Duty": "—",
                        "GST": "—", "Surtax": "; ".join(
                            f"{s[3]} ({s[2]})" for s in srows) or "—"})
    return out


def get_surtax_data(con, item_norm: str, import_date: date | None = None):
    """Structured US surtax rows: [period, tariff_item, effective_date, surtax_str].

    import_date filters to rows in force on that date (None = all rows).
    """
    out = []
    for c in get_counter_rows(con, item_norm):
        cd = dict(c)
        if not surtax_applies(cd["effective_date"], cd["period"], import_date):
            continue
        out.append([cd["period"], cd["tariff_item"], cd["effective_date"],
                    f"+{cd['rate']}" + ("" if str(cd["rate"]).endswith("%") else "%")])
    return out


def build_report(con, tariff_query: str, country: str, value: float | None = 1000.0,
                 qty: float | None = None, qty_unit: str | None = None,
                 import_date: date | None = None, rev: str | None = None) -> str:
    if value is not None and value <= 0:
        value = None
    if qty is not None and qty <= 0:
        qty = None
    data = get_duty_data(con, tariff_query, country, value, qty, qty_unit, rev)
    if data is None:
        return f"[{tariff_query}]  NOT FOUND in tariff schedule.\n"
    show_money = data["show_money"]
    best = data["best"]
    lines = []
    lines.append(f"HS: {data['hs']}  (eff {data['eff']})"
                   + (f"  [rate inherited from parent {data['inherited_from']}]"
                      if data.get("inherited_from")
                      else (f"  [closest rated line for {data['resolved_from']}]"
                            if data.get("resolved_from") else "")))
    lines.append(f"Desc: {data['context'][:220]}")
    lines.append(f"UOM: {data['uom']}   Origin: {country}")
    if data.get("note"):
        lines.append(data["note"])
    lines.append("")
    duty_rows = [[r[0] + (" *" if best and r[0] == best[0] else ""), r[1], r[2]]
                 for r in data["rows"]]
    show_duty_col = any(r[2].startswith("$") for r in duty_rows)
    if duty_rows:
        if show_duty_col:
            _dh = f"Duty on ${value:,.2f}" if value else "Duty $"
            lines.append(_ascii_table(["Treatment", "Rate", _dh], duty_rows))
        else:
            lines.append(_ascii_table(["Treatment", "Rate"],
                                      [[r[0], r[1]] for r in duty_rows]))
    else:
        lines.append("(no rates found for eligible treatments)")
    lines.append("")
    if best:
        _adv = best[1]
        lines.append(f"* Lowest duty option: {best[0]} @ "
                     + ("Free" if _adv == 0 else (f"{_adv}%" if _adv is not None
                                                  else f"${best[2]:,.2f}")))
    else:
        lines.append("=> No computable duty (enter a value and/or quantity).")
    lines.append(data["gst"])
    # US reciprocal surtax
    if country and country.lower() in US_ORIGINS or (country or "").lower() in US_ORIGINS:
        srows = get_surtax_data(con, data["norm"], import_date)
        if srows:
            lines.append("")
            lines.append(f"US RECIPROCAL SURTAX - {len(srows)} record(s), stacks on top:")
            lines.append(_ascii_table(["Period", "Tariff item", "Effective", "Surtax"],
                                      srows))
        else:
            lines.append("US surtax list: this HS not listed - no reciprocal surtax.")
    lines.append("")
    return "\n".join(lines)


# ---------------- CBSA download + versioned import ----------------
CBSA_BASE = "https://www.cbsa-asfc.gc.ca"
CBSA_MENU = CBSA_BASE + "/trade-commerce/tariff-tarif/menu-eng.html"


def _http_get(url: str, timeout: int = 60) -> str:
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8", errors="replace")


def cbsa_find_latest() -> dict:
    """Newest CBSA tariff revision with an Access database.

    Returns {year, rev, effective, zip_url, countries_url}.
    """
    import urllib.parse
    menu = _http_get(CBSA_MENU)
    years = sorted(set(re.findall(r"/trade-commerce/tariff-tarif/(20\d\d)/menu-eng\.html", menu)))
    if not years:
        raise RuntimeError("CBSA menu page: no tariff year found")
    year = years[-1]
    ypage = _http_get(f"{CBSA_BASE}/trade-commerce/tariff-tarif/{year}/menu-eng.html")
    parts = re.split(r">\s*(T20\d\d(?:-\d+)?)\s*</span>\s*<span[^>]*>\s*Effective date:\s*"
                     r"(\d{4}-\d{2}-\d{2})", ypage)
    for i in range(1, len(parts), 3):
        rev, eff = parts[i], parts[i + 1]
        body = parts[i + 2] if i + 2 < len(parts) else ""
        zip_href = None
        for m in re.finditer(r'<a\s+href="([^"]+\.zip)"[^>]*>(.*?)</a>', body, re.S):
            if "Microsoft Access" in m.group(2):
                zip_href = m.group(1)
                break
        if zip_href:
            return {"year": year, "rev": rev, "effective": eff,
                    "zip_url": urllib.parse.urljoin(CBSA_BASE + "/", zip_href),
                    "countries_url": f"{CBSA_BASE}/trade-commerce/tariff-tarif/"
                                     f"{year}/html/countries-pays-eng.html"}
    raise RuntimeError("CBSA year page: no Access database link found")


def download_file(url: str, dest: str, hook=None, timeout: int = 300) -> str:
    """Download with optional hook(done_bytes, total_bytes_or_None). Returns dest."""
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        total = r.headers.get("Content-Length")
        total = int(total) if total else None
        done = 0
        with open(dest, "wb") as f:
            while True:
                chunk = r.read(256 * 1024)
                if not chunk:
                    break
                f.write(chunk)
                done += len(chunk)
                if hook:
                    hook(done, total)
    return dest


def parse_countries_html(html: str):
    """CBSA countries page -> [(country, mfn, gpt, ldct, other)]."""
    out = []
    for r in re.findall(r"<tr>(.*?)</tr>", html, re.S):
        if '<th scope="row">' not in r:
            continue
        cells = re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", r, re.S)
        clean = []
        for c in cells:
            c = re.sub(r"<span[^>]*>(.*?)</span>", r"\1", c, flags=re.S)
            c = re.sub(r"<[^>]+>", "", c).replace("&nbsp;", "").strip()
            clean.append(c)
        if len(clean) == 5:
            country, mfn, gpt, ldct, other = clean
            out.append((country,
                        "yes" if mfn.strip().lower() == "yes" else "no",
                        "yes" if gpt.strip().lower() == "yes" else "no",
                        "yes" if ldct.strip().lower() == "yes" else "no",
                        other))
    return out


def read_accdb_tphs(accdb_path: str):
    """Read TPHS table from an Access .accdb via ODBC. Returns (columns, rows)."""
    try:
        import pyodbc
    except ImportError:
        raise RuntimeError("pyodbc is not installed (pip install pyodbc)")
    try:
        con = pyodbc.connect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
                             f"DBQ={accdb_path}")
    except Exception as e:
        raise RuntimeError(f"Cannot open Access file (need MS Access Database Engine): {e}")
    cur = con.cursor()
    cols = [c.column_name for c in cur.columns(table="TPHS")]
    if "TARIFF" not in cols:
        raise RuntimeError("TPHS table not found in Access file")
    rows = cur.execute("SELECT * FROM TPHS").fetchall()
    con.close()
    return cols, [tuple(r) for r in rows]


def _tphs_frame(cols, rows):
    import pandas as pd
    df = pd.DataFrame(rows, columns=cols)
    df["TARIFF_NORM"] = df["TARIFF"].astype(str).str.replace(r"\D", "", regex=True)
    for c in ("DESC1", "DESC2", "DESC3"):
        if c not in df.columns:
            df[c] = ""
    df["DESC_FULL"] = (df["DESC1"].fillna("").astype(str) + " | "
                       + df["DESC2"].fillna("").astype(str) + " | "
                       + df["DESC3"].fillna("").astype(str)
                       ).str.replace(r"\s*\|\s*None\s*", "", regex=True).str.strip(" |")
    return df


def import_tariff_rev(db_path: str, rev: str, effective: str, source: str,
                      tphs_cols, tphs_rows, countries) -> dict:
    """Store one tariff revision (replacing same rev) + paired country list.

    Returns {tphs_rows, countries, prev_rev} for change analysis.
    """
    import datetime as _dt
    con = sqlite3.connect(db_path)
    prev = None
    try:
        r = con.execute("SELECT value FROM db_meta WHERE key='active_rev'").fetchone()
        prev = r[0] if r else None
    except Exception:
        pass
    df = _tphs_frame(list(tphs_cols), [list(r) for r in tphs_rows])
    df["rev"] = rev
    con.execute("DELETE FROM tphs WHERE rev=?", (rev,))
    df.to_sql("tphs", con, index=False, if_exists="append")
    year = re.search(r"20\d\d", rev)
    crev = "T" + year.group(0) if year else "T2026"
    con.execute("DELETE FROM country_map WHERE rev=?", (crev,))
    con.executemany("INSERT INTO country_map(country, mfn, gpt, ldct, other, rev) "
                    "VALUES (?,?,?,?,?,?)", [(c, m, g, l, o, crev) for c, m, g, l, o in countries])
    now = _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    for k, v in {f"rev:{rev}:effective": effective, f"rev:{rev}:source": source,
                 f"rev:{rev}:country_rev": crev, f"rev:{rev}:imported": now,
                 "active_rev": rev}.items():
        con.execute("INSERT OR REPLACE INTO db_meta VALUES (?,?)", (k, v))
    con.execute("CREATE INDEX IF NOT EXISTS ix_tphs_rev_norm ON tphs(rev, TARIFF_NORM)")
    con.execute("CREATE INDEX IF NOT EXISTS ix_cmap_rev ON country_map(rev, country)")
    con.commit()
    n_t = con.execute("SELECT COUNT(*) FROM tphs WHERE rev=?", (rev,)).fetchone()[0]
    n_c = con.execute("SELECT COUNT(*) FROM country_map WHERE rev=?", (crev,)).fetchone()[0]
    con.close()
    return {"tphs_rows": n_t, "countries": n_c, "prev_rev": prev if prev != rev else None}


class App:
    def __init__(self):
        self.root = ttkb.Window(themename="flatly")
        self.root.title("Canada Tariff Lookup  •  T2026-2 + US Surtax")
        self.root.geometry("1220x760")
        self.root.minsize(1000, 620)

        self.con = get_connection()
        self.rev = active_rev(self.con)
        self.countries = load_countries(self.con, self.rev)
        self.basket: list[str] = []
        self._results_cache: list[tuple] = []
        ttkb.Style().configure("Treeview", font=("Consolas", 11), rowheight=28)
        # light-theme widget colors (flatly)
        self.LB_BG = "#ffffff"
        self.LB_FG = "#212529"
        self.LB_SEL = "#0d6efd"
        self.LB_HL = "#adb5bd"
        self.CANVAS_BG = self.root.cget("background")

        mono = tkfont.Font(family="Consolas", size=11)
        title_font = tkfont.Font(size=16, weight="bold")
        sub_font = tkfont.Font(size=10)

        # ----- header -----
        header = ttkb.Frame(self.root, padding=(16, 12, 16, 6))
        header.pack(fill=X)
        ttkb.Label(header, text="🇨🇦  Canada Tariff Lookup",
                   font=title_font, bootstyle="inverse-dark").pack(side=LEFT)
        self.rev_combo = ttkb.Combobox(header, values=installed_revs(self.con),
                                       width=10, bootstyle="info")
        self.rev_combo.set(self.rev)
        self.rev_combo.pack(side=LEFT, padx=(16, 0))
        self.rev_combo.bind("<<ComboboxSelected>>", self.on_rev_change)
        ttkb.Button(header, text="Compare…", bootstyle="secondary-outline",
                    command=self.open_compare).pack(side=LEFT, padx=(6, 0))
        ttkb.Button(header, text="Database…", bootstyle="secondary-outline",
                    command=self.open_updater).pack(side=LEFT, padx=(6, 0))
        ttkb.Button(header, text="About", bootstyle="secondary-outline",
                    command=self.open_about).pack(side=LEFT, padx=(6, 0))
        self.subtitle = ttkb.Label(header, text="", font=sub_font, bootstyle="secondary")
        self.subtitle.pack(side=LEFT, padx=16)
        ttkb.Separator(self.root, bootstyle="secondary").pack(fill=X, padx=16, pady=4)

        # ----- main split -----
        main = ttkb.Frame(self.root, padding=(16, 6, 16, 0))
        main.pack(fill=BOTH, expand=True)
        main.columnconfigure(0, weight=0, minsize=500)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        side = ttkb.Frame(main, width=500)
        side.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        side.grid_propagate(False)
        side.columnconfigure(0, weight=1)

        body = ttkb.Frame(main)
        body.grid(row=0, column=1, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        # ----- sidebar: HS codes -----
        hs_box = ttkb.LabelFrame(side, text="  HS codes  ", padding=10, bootstyle="info")
        hs_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        hs_box.columnconfigure(0, weight=1)

        self.hs_entry = ttkb.Entry(hs_box)
        self.hs_entry.grid(row=0, column=0, sticky="ew")
        self.hs_entry.bind("<Return>", lambda _e: self.on_search())

        btn_row = ttkb.Frame(hs_box)
        btn_row.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttkb.Button(btn_row, text="🔍 Search", bootstyle="primary",
                    command=self.on_search).pack(side=LEFT, padx=(0, 6))
        ttkb.Button(btn_row, text="➕ Add typed", bootstyle="success-outline",
                    command=self.on_add_typed).pack(side=LEFT)

        ttkb.Label(hs_box, text="Matches — double-click to add:",
                   bootstyle="secondary").grid(row=2, column=0, sticky="w", pady=(10, 2))
        res_frame = ttkb.Frame(hs_box)
        res_frame.grid(row=3, column=0, sticky="ew")
        res_frame.columnconfigure(0, weight=1)
        self.results_box = tk.Listbox(res_frame, height=8, exportselection=False,
                                      font=mono, bg=self.LB_BG, fg=self.LB_FG,
                                      selectbackground=self.LB_SEL, selectforeground="#ffffff",
                                      relief="flat", highlightthickness=1,
                                      highlightbackground=self.LB_HL)
        self.results_box.grid(row=0, column=0, sticky="ew")
        res_sb = ttkb.Scrollbar(res_frame, orient=VERTICAL, command=self.results_box.yview,
                                bootstyle="info-round")
        res_sb.grid(row=0, column=1, sticky="ns")
        self.results_box.configure(yscrollcommand=res_sb.set)
        self.results_box.bind("<Double-Button-1>", lambda _e: self.on_add_selected())

        ttkb.Label(hs_box, text="Selected HS codes:", bootstyle="secondary").grid(
            row=4, column=0, sticky="w", pady=(10, 2))
        self.basket_box = tk.Listbox(hs_box, height=5, exportselection=False,
                                     font=mono, bg=self.LB_BG, fg=self.LB_FG,
                                     selectbackground=self.LB_SEL, selectforeground="#ffffff",
                                     relief="flat", highlightthickness=1,
                                     highlightbackground=self.LB_HL)
        self.basket_box.grid(row=5, column=0, sticky="ew")

        brow = ttkb.Frame(hs_box)
        brow.grid(row=6, column=0, sticky="ew", pady=(8, 0))
        ttkb.Button(brow, text="✖ Remove", bootstyle="danger-outline",
                    command=self.on_remove).pack(side=LEFT, padx=(0, 6))
        ttkb.Button(brow, text="Clear", bootstyle="secondary-outline",
                    command=self.on_clear_basket).pack(side=LEFT)

        # ----- sidebar: origin + value -----
        org_box = ttkb.LabelFrame(side, text="  Origin & value  ", padding=10, bootstyle="info")
        org_box.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        org_box.columnconfigure(0, weight=1)

        ttkb.Label(org_box, text="Origin country:", bootstyle="secondary").grid(
            row=0, column=0, sticky="w")
        self.country_var = tk.StringVar()
        self.country_entry = ttkb.Entry(org_box, textvariable=self.country_var)
        self.country_entry.grid(row=1, column=0, sticky="ew", pady=(2, 0))
        self.country_entry.bind("<KeyRelease>", self.on_country_type)
        self.country_entry.bind("<Return>", self.on_country_enter)
        self.suggest_box = tk.Listbox(org_box, height=4, exportselection=False,
                                      font=mono, bg=self.LB_BG, fg=self.LB_FG,
                                      selectbackground=self.LB_SEL, selectforeground="#ffffff",
                                      relief="flat", highlightthickness=1,
                                      highlightbackground=self.LB_HL)
        self.suggest_box.grid(row=2, column=0, sticky="ew", pady=(4, 0))
        self.suggest_box.bind("<Double-Button-1>", lambda _e: self.on_country_pick())
        self.suggest_box.bind("<Return>", lambda _e: self.on_country_pick())
        self._refresh_suggest("")

        ttkb.Label(org_box, text="Value for duty (CAD) — blank = $1,000:",
                   bootstyle="secondary").grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.value_entry = ttkb.Entry(org_box)
        self.value_entry.grid(row=4, column=0, sticky="ew", pady=(2, 0))
        self.value_entry.insert(0, "1000")

        ttkb.Label(org_box, text="Quantity (for specific duties, e.g. ¢/kg):",
                   bootstyle="secondary").grid(row=5, column=0, sticky="w", pady=(10, 0))
        qframe = ttkb.Frame(org_box)
        qframe.grid(row=6, column=0, sticky="ew", pady=(2, 0))
        qframe.columnconfigure(0, weight=1)
        self.qty_entry = ttkb.Entry(qframe)
        self.qty_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.unit_combo = ttkb.Combobox(qframe, values=["KGM", "TNE", "GRM", "LTR",
                                                        "NMB", "DZN", "PAR", "MTR"],
                                        width=8)
        self.unit_combo.grid(row=0, column=1, sticky="e")

        ttkb.Label(org_box, text="Import date (US surtax shown only if in force then):",
                   bootstyle="secondary").grid(row=7, column=0, sticky="w", pady=(10, 0))
        dframe = ttkb.Frame(org_box)
        dframe.grid(row=8, column=0, sticky="ew", pady=(2, 0))
        dframe.columnconfigure(0, weight=1)
        self.date_entry = DateEntry(dframe, dateformat="%Y-%m-%d", bootstyle="info")
        self.date_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.date_all_var = tk.BooleanVar(value=False)
        ttkb.Checkbutton(dframe, text="All periods", variable=self.date_all_var,
                         bootstyle="info-round-toggle").grid(row=0, column=1, sticky="e")

        # ----- results pane (native ttkbootstrap tables) -----
        res_pane = ttkb.LabelFrame(body, text="  Applicable duty & tax  ", padding=10,
                                   bootstyle="success")
        res_pane.grid(row=0, column=0, rowspan=2, sticky="nsew")
        res_pane.columnconfigure(0, weight=1)
        res_pane.rowconfigure(0, weight=1)
        self.results_scroll = ttkb.Scrollbar(res_pane, orient=VERTICAL,
                                             bootstyle="success-round")
        self.results_scroll.grid(row=0, column=1, sticky="ns")
        canvas = tk.Canvas(res_pane, highlightthickness=0, bg=self.CANVAS_BG)
        canvas.grid(row=0, column=0, sticky="nsew")
        self.results_scroll.configure(command=canvas.yview)
        canvas.configure(yscrollcommand=self.results_scroll.set)
        self.results_area = ttkb.Frame(canvas)
        self._results_win = canvas.create_window((0, 0), window=self.results_area,
                                                 anchor="nw")

        def _sync_scroll(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfigure(self._results_win, width=canvas.winfo_width())

        self.results_area.bind("<Configure>", _sync_scroll)
        canvas.bind("<Configure>", _sync_scroll)
        self._show_results_placeholder()

        act = ttkb.Frame(body)
        act.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttkb.Button(act, text="⚖  Show duties & taxes", bootstyle="success",
                    command=self.on_calculate).pack(side=LEFT, padx=(0, 6))
        ttkb.Button(act, text="Clear results", bootstyle="secondary-outline",
                    command=self._show_results_placeholder).pack(side=LEFT)
        ttkb.Button(act, text="Export Excel", bootstyle="info-outline",
                    command=lambda: self.on_export("xlsx")).pack(side=LEFT, padx=(6, 0))
        ttkb.Button(act, text="Export CSV", bootstyle="info-outline",
                    command=lambda: self.on_export("csv")).pack(side=LEFT, padx=(6, 0))
        ttkb.Button(act, text="Export PDF", bootstyle="info-outline",
                    command=lambda: self.on_export("pdf")).pack(side=LEFT, padx=(6, 0))
        self.status = ttkb.Label(act, text="T2026-2 + US surtax loaded",
                                 bootstyle="secondary")
        self.status.pack(side=RIGHT)

        # ----- footer status bar -----
        ttkb.Separator(self.root, bootstyle="secondary").pack(fill=X, padx=16, pady=(6, 0))
        self.footer = ttkb.Label(self.root, text="", bootstyle="secondary",
                                  padding=(16, 4))
        self.footer.pack(anchor="w")
        self.refresh_db_labels()

    def run(self):
        self.root.mainloop()

    # ---- handlers ----
    def on_search(self):
        q = self.hs_entry.get().strip()
        rows = search_tariff(self.con, q, rev=self.rev)
        self._results_cache = rows
        self.results_box.delete(0, "end")
        for r in rows[:100]:
            self.results_box.insert("end", f"{r[0]} — {(r[1] or '')[:60]}")
        self.status.configure(text=f"{len(rows)} match(es)")

    def _selected_result_code(self) -> str | None:
        sel = self.results_box.curselection()
        if not sel or not self._results_cache:
            return None
        return self._results_cache[sel[0]][0]

    def on_add_selected(self):
        code = self._selected_result_code()
        if code:
            self.add_code(code)

    def on_add_typed(self):
        raw = self.hs_entry.get().strip()
        if not raw:
            return
        for part in re.split(r"[,;\s]+", raw):
            if part.strip():
                self.add_code(part.strip())
        self.hs_entry.delete(0, "end")

    def add_code(self, code: str):
        row = find_tariff_row(self.con, code, self.rev)
        canon = row["TARIFF"] if row is not None else code
        if canon in self.basket:
            self.status.configure(text=f"{canon} already in list")
            return
        if row is None:
            self.status.configure(text=f"{code}: not found — not added")
            return
        self.basket.append(canon)
        self.basket_box.insert("end", canon)
        uom = (row["UOM"] or "").strip()
        if uom in ("KGM", "TNE", "GRM", "LTR", "NMB", "DZN", "PAR", "MTR") \
                and not self.unit_combo.get():
            self.unit_combo.set(uom)
        self.status.configure(text=f"Added {canon} ({len(self.basket)} total)")

    def _get_import_date(self):
        if self.date_all_var.get():
            return None
        try:
            d = self.date_entry.get_date()
            if isinstance(d, datetime):
                return d.date()
            if isinstance(d, date):
                return d
        except Exception:
            pass
        try:
            return date.fromisoformat(self.date_entry.entry.get().strip())
        except Exception:
            return None

    def on_remove(self):
        sel = self.basket_box.curselection()
        if sel:
            self.basket_box.delete(sel[0])
            del self.basket[sel[0]]

    def on_clear_basket(self):
        self.basket.clear()
        self.basket_box.delete(0, "end")

    def _refresh_suggest(self, prefix: str):
        p = prefix.lower()
        matches = [c for c in self.countries if p in c.lower()][:8] if p else \
            ["United States of America", "China", "Mexico", "Germany",
             "United Kingdom", "Japan", "South Korea", "India"]
        self.suggest_box.delete(0, "end")
        for m in matches:
            self.suggest_box.insert("end", m)

    def on_country_type(self, _event):
        self._refresh_suggest(self.country_var.get())

    def on_country_pick(self):
        sel = self.suggest_box.curselection()
        if sel:
            self.country_var.set(self.suggest_box.get(sel[0]))
            self.country_entry.icursor("end")

    def on_country_enter(self, _event):
        cur = self.country_var.get()
        resolved = resolve_country(self.con, cur, self.rev)
        if resolved:
            self.country_var.set(resolved)
        self._refresh_suggest(self.country_var.get())

    def _clear_results_area(self):
        for w in self.results_area.winfo_children():
            w.destroy()

    def _show_results_placeholder(self):
        self._clear_results_area()
        ttkb.Label(self.results_area,
                   text="Add HS codes, pick an origin country, then press "
                        "“⚖  Show duties & taxes”.",
                   bootstyle="secondary").pack(anchor="w", pady=20)
        if hasattr(self, "status"):
            self.status.configure(text="T2026-2 + US surtax loaded")

    def on_export(self, kind):
        last = getattr(self, "_last", None)
        if not last or not (last["duties"] or last["surtax"]):
            self.status.configure(text="Calculate first, then export")
            return
        if kind == "xlsx":
            ext, ftype = ".xlsx", [("Excel workbook", "*.xlsx")]
        elif kind == "pdf":
            ext, ftype = ".pdf", [("PDF report", "*.pdf")]
        else:
            ext, ftype = ".csv", [("CSV file", "*.csv")]
        path = filedialog.asksaveasfilename(defaultextension=ext, filetypes=ftype,
                                            initialfile=f"tariff-report{ext}")
        if not path:
            return
        try:
            export_results(path, last["meta"], last["duties"],
                           last["surtax"], last.get("summary", []))
            self.status.configure(text=f"Exported {path}")
        except Exception as e:
            self.status.configure(text=f"Export failed: {e}")

    # ---- tariff revisions ----
    def refresh_db_labels(self):
        info = rev_info(self.con, self.rev)
        try:
            n_c = self.con.execute("SELECT COUNT(*) FROM country_map WHERE rev=?",
                                   (country_rev_for(self.con, self.rev),)).fetchone()[0]
            n_s = self.con.execute("SELECT COUNT(*) FROM counter_tariffs").fetchone()[0]
        except Exception:
            n_c, n_s = 0, 0
        self.subtitle.configure(
            text=f"{self.rev} (eff {info.get('effective') or '?'})  •  US surtax, all periods")
        self.footer.configure(
            text=f"{info['rows']:,} tariff lines ({self.rev})  •  {n_c} countries  •  "
                 f"{n_s:,} US surtax rows")
        self.root.title(f"Canada Tariff Lookup  •  {self.rev} + US Surtax")

    def on_rev_change(self, _event=None):
        new = self.rev_combo.get().strip()
        if new and new != self.rev:
            self.rev = new
            self.con.execute("INSERT OR REPLACE INTO db_meta VALUES ('active_rev',?)", (new,))
            self.con.commit()
            self.countries = load_countries(self.con, self.rev)
            self._refresh_suggest(self.country_var.get())
            self.refresh_db_labels()
            self.status.configure(text=f"Switched to {new} — recalculate to refresh")

    def open_about(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("About")
        dlg.geometry("560x470")
        dlg.resizable(False, False)
        body = ttkb.Frame(dlg, padding=20)
        body.pack(fill=BOTH, expand=True)
        ttkb.Label(body, text="🇨🇦  Canada Tariff Lookup",
                   font=tkfont.Font(size=16, weight="bold")).pack(anchor="w")
        ttkb.Label(body, text=f"Version {APP_VERSION}  •  Customs Tariff T2026-2 + US surtax",
                   bootstyle="secondary").pack(anchor="w", pady=(0, 10))
        ttkb.Label(body, text=f"Developed by {APP_AUTHOR}",
                   font=tkfont.Font(weight="bold")).pack(anchor="w")

        def link(parent, text, url):
            lb = tk.Label(parent, text=text, fg="#0d6efd", cursor="hand2",
                          font=tkfont.Font(underline=True))
            lb.pack(anchor="w")
            lb.bind("<Button-1>", lambda _e: webbrowser.open(url))
            return lb

        link(body, "GitHub: incredibleamir-dot", APP_GITHUB)
        link(body, "Repository: CBSA-Tariff-Lookup", APP_REPO)
        ttkb.Separator(body, bootstyle="secondary").pack(fill=X, pady=12)
        ttkb.Label(body, text="Data transparency", font=tkfont.Font(weight="bold")).pack(anchor="w")
        ttkb.Label(body, wraplength=500, justify="left", text=(
            "• Customs Tariff (rates, countries): official CBSA data, kept current with "
            "the built-in Database updater — every imported revision is stored and "
            "selectable, with A-vs-B change comparison.\n\n"
            "• US reciprocal surtax list: a STATIC snapshot compiled 2026-10-01 from "
            "Finance Canada's complete list of US products subject to counter tariffs "
            "(covers 2025-03-04, 2025-03-13, 2025-04-09 and 2026-09-08 measures). "
            "It is hard-coded into the database and is NOT refreshed by the updater. "
            "Always confirm live surtax status via Finance Canada / CBSA Customs Notices "
            "before transacting.")).pack(anchor="w")
        ttkb.Separator(body, bootstyle="secondary").pack(fill=X, pady=12)
        ttkb.Label(body, wraplength=500, justify="left", bootstyle="secondary", text=(
            "For information only — verify classifications, rates and surtax periods "
            "with the CBSA Customs Tariff and Canada Gazette orders.")).pack(anchor="w")
        ttkb.Button(body, text="Close", bootstyle="secondary",
                    command=dlg.destroy).pack(pady=(12, 0))

    def open_compare(self):
        revs = installed_revs(self.con)
        if len(revs) < 2:
            self.status.configure(text="Only one revision installed — import another to compare")
            return
        dlg = tk.Toplevel(self.root)
        dlg.title("Compare tariff revisions")
        dlg.geometry("950x620")
        top = ttkb.Frame(dlg, padding=12)
        top.pack(fill=X)
        ttkb.Label(top, text="From:").pack(side=LEFT)
        cb_a = ttkb.Combobox(top, values=revs, width=10)
        cb_a.set(revs[-2])
        cb_a.pack(side=LEFT, padx=6)
        ttkb.Label(top, text="To:").pack(side=LEFT)
        cb_b = ttkb.Combobox(top, values=revs, width=10)
        cb_b.set(revs[-1])
        cb_b.pack(side=LEFT, padx=6)
        fcombo = ttkb.Combobox(top, values=["All", "ADDED", "REMOVED", "RATE CHANGE",
                                            "DESCRIPTION CHANGE"], width=20)
        fcombo.set("All")
        fcombo.pack(side=LEFT, padx=6)
        ttkb.Button(top, text="Compare", bootstyle="primary",
                    command=lambda: run()).pack(side=LEFT, padx=6)
        counts = ttkb.Label(top, text="", bootstyle="info")
        counts.pack(side=LEFT, padx=12)
        holder = ttkb.Frame(dlg, padding=(12, 0, 12, 12))
        holder.pack(fill=BOTH, expand=True)
        state = {"rows": []}

        def show(rows):
            for w in holder.winfo_children():
                w.destroy()
            Tableview(holder,
                      coldata=[{"text": "Change", "width": 150, "stretch": False},
                               {"text": "HS", "width": 110, "stretch": False},
                               {"text": "Description", "stretch": True},
                               {"text": "Detail", "stretch": True}],
                      rowdata=rows, paginated=True, pagesize=20, searchable=True,
                      bootstyle="info", yscrollbar=True, height=20).pack(fill=BOTH, expand=True)

        def apply_filter(_e=None):
            f = fcombo.get()
            show([r for r in state["rows"] if f == "All" or r[0] == f])

        def run():
            a, b = cb_a.get().strip(), cb_b.get().strip()
            if not a or not b or a == b:
                counts.configure(text="Pick two different revisions")
                return
            rows = diff_revs(self.con, a, b)
            state["rows"] = rows
            by = {}
            for ch, *_rest in rows:
                by[ch] = by.get(ch, 0) + 1
            counts.configure(text=f"{a} → {b}: {len(rows)} changes (" +
                                  ", ".join(f"{k} {v}" for k, v in sorted(by.items())) + ")")
            apply_filter()

        fcombo.bind("<<ComboboxSelected>>", apply_filter)
        run()

    # ---- CBSA updater ----
    def _bg(self, fn, done):
        def w():
            try:
                res = (True, fn())
            except Exception as e:
                res = (False, str(e))
            try:
                self.root.after(0, lambda: done(*res))
            except Exception:
                pass
        threading.Thread(target=w, daemon=True).start()

    def _poll_updater(self):
        try:
            while self._upd_logbuf:
                self._upd_log.insert("end", self._upd_logbuf.pop(0) + "\n")
                self._upd_log.see("end")
            tot = self._upd_prog.get("total")
            if tot:
                self._upd_bar.configure(
                    value=100 * self._upd_prog.get("done", 0) / tot)
            if not self._upd_prog.get("finished"):
                self.root.after(250, self._poll_updater)
        except Exception:
            pass

    def open_updater(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Tariff database")
        dlg.geometry("660x540")
        body = ttkb.Frame(dlg, padding=14)
        body.pack(fill=BOTH, expand=True)
        ttkb.Label(body, text="Installed revisions:",
                   font=tkfont.Font(weight="bold")).pack(anchor="w")
        lines = []
        for r in installed_revs(self.con):
            ri = rev_info(self.con, r)
            lines.append(f"{r} — eff {ri.get('effective') or '?'} — {ri['rows']:,} lines"
                         + ("   [active]" if r == self.rev else ""))
        ttkb.Label(body, text="\n".join(lines)).pack(anchor="w", pady=(0, 10))
        self._upd_status = ttkb.Label(body, text="")
        self._upd_status.pack(anchor="w")
        brow = ttkb.Frame(body)
        brow.pack(fill=X, pady=6)
        ttkb.Button(brow, text="Check CBSA for updates", bootstyle="info",
                    command=self._upd_check).pack(side=LEFT, padx=(0, 6))
        self._upd_import_btn = ttkb.Button(brow, text="Download & import latest",
                                           bootstyle="success", state="disabled",
                                           command=self._upd_import)
        self._upd_import_btn.pack(side=LEFT)
        self._upd_bar = ttkb.Progressbar(body, bootstyle="success-striped", maximum=100)
        self._upd_bar.pack(fill=X, pady=8)
        self._upd_log = tk.Text(body, height=12, font=("Consolas", 10))
        self._upd_log.pack(fill=BOTH, expand=True)
        self._upd_info = None
        self._upd_logbuf = []
        self._upd_prog = {}

    def _upd_check(self):
        self._upd_status.configure(text="Contacting CBSA…")
        self._bg(cbsa_find_latest, self._upd_checked)

    def _upd_checked(self, ok, res):
        if not ok:
            self._upd_status.configure(text=f"Check failed: {res}")
            return
        self._upd_info = res
        if res["rev"] in installed_revs(self.con):
            self._upd_status.configure(
                text=f"Latest on CBSA is {res['rev']} (eff {res['effective']}) — already installed")
            self._upd_import_btn.configure(state="disabled")
        else:
            self._upd_status.configure(
                text=f"Latest on CBSA: {res['rev']} (eff {res['effective']}) — NEW, ready to download")
            self._upd_import_btn.configure(state="normal")

    def _upd_import(self):
        info = getattr(self, "_upd_info", None)
        if not info:
            return
        self._upd_import_btn.configure(state="disabled")
        self._upd_bar.configure(value=0)
        self._upd_prog = {"done": 0, "total": None, "finished": False}

        def hooked(done, total):
            self._upd_prog["done"] = done
            self._upd_prog["total"] = total

        def job():
            import tempfile
            import zipfile
            tmpd = tempfile.mkdtemp(prefix="cbsa_tariff_")
            zp = str(Path(tmpd) / "latest.zip")
            self._upd_logbuf.append("Downloading " + info["zip_url"])
            download_file(info["zip_url"], zp, hook=hooked)
            names = zipfile.ZipFile(zp).namelist()
            acc = next(n for n in names if n.lower().endswith((".accdb", ".mdb")))
            zipfile.ZipFile(zp).extractall(tmpd)
            cols, rows = read_accdb_tphs(str(Path(tmpd) / acc))
            self._upd_logbuf.append(f"Access table read: {len(rows):,} rows")
            countries = parse_countries_html(_http_get(info["countries_url"]))
            self._upd_logbuf.append(f"Country list: {len(countries)} entries")
            return {**import_tariff_rev(str(DB_PATH), info["rev"], info["effective"],
                                        info["zip_url"], cols, rows, countries),
                    "rev": info["rev"]}

        self._bg(job, self._upd_imported)
        self._poll_updater()

    def _upd_imported(self, ok, res):
        self._upd_prog["finished"] = True
        self._poll_updater()
        if not ok:
            self._upd_status.configure(text=f"Import failed: {res}")
            messagebox.showerror("Import failed", str(res))
            return
        prev = res.get("prev_rev")
        summary = ""
        if prev:
            diff = diff_revs(self.con, prev, res["rev"])
            by = {}
            for ch, *_rest in diff:
                by[ch] = by.get(ch, 0) + 1
            summary = ", ".join(f"{k} {v}" for k, v in sorted(by.items())) or "no changes"
            self._upd_logbuf.append(f"Changes vs {prev}: {summary or 'none'}")
            self._poll_updater()
        self.rev = res["rev"]
        self.rev_combo.configure(values=installed_revs(self.con))
        self.rev_combo.set(self.rev)
        self.countries = load_countries(self.con, self.rev)
        self._refresh_suggest(self.country_var.get())
        self.refresh_db_labels()
        self._upd_status.configure(text=f"Imported {res['rev']}: {res['tphs_rows']:,} lines")
        messagebox.showinfo("Import complete",
                            f"{res['rev']}: {res['tphs_rows']:,} tariff lines imported.\n"
                            + (f"Changes vs {prev}: {summary}" if prev else ""))

    def on_calculate(self):
        country_raw = self.country_var.get().strip()
        country = resolve_country(self.con, country_raw, self.rev)
        if not country:
            self.status.configure(text="Pick a valid origin country first")
            return
        self.country_var.set(country)
        raw = self.value_entry.get().strip().replace(",", "")
        try:
            value = float(raw) if raw else 1000.0
        except ValueError:
            value = 1000.0
        if value <= 0:
            value = 1000.0
        qraw = self.qty_entry.get().strip().replace(",", "")
        try:
            qty = float(qraw) if qraw else None
        except ValueError:
            qty = None
        if qty is not None and qty <= 0:
            qty = None
        qty_unit = self.unit_combo.get().strip() or None
        import_date = self._get_import_date()
        if not self.basket:
            self.status.configure(text="Add at least one HS code")
            return
        is_us = country.lower() in US_ORIGINS

        self._clear_results_area()
        note = origin_note(self.con, country, self.rev)
        info = (f"Origin: {country}    Value for duty: ${value:,.2f}"
                + (f"    Qty: {qty:g} {qty_unit or ''}" if qty else "")
                + (f"    Import date: {import_date.isoformat()}" if import_date
                   else "    Surtax: all periods")
                + f"    Treatments: {', '.join(eligible_treatments(self.con, country, self.rev))}"
                + (f"    // {note}" if note else ""))
        ttkb.Label(self.results_area, text=info, bootstyle="info",
                   font=tkfont.Font(weight="bold")).pack(anchor="w", pady=(0, 8))

        duty_rowdata, surtax_rowdata = [], []
        exp_duties, exp_surtax, per_hs = [], [], []
        for code in self.basket:
            data = get_duty_data(self.con, code, country, value, qty, qty_unit, self.rev)
            if data is None:
                per_hs.append((None, code, []))
                continue
            best = data["best"]
            for (treat, rate, duty, _adv, _amt) in data["rows"]:
                star = "★" if best and treat == best[0] else ""
                duty_rowdata.append([data["hs"], data["context"][:80], treat, rate,
                                     duty, star])
                exp_duties.append({"HS": data["hs"], "Description": data["context"],
                                   "Treatment": treat, "Rate": rate,
                                   "Duty": duty, "Best": star})
            _srows = (get_surtax_data(self.con, data["norm"], import_date) if is_us else [])
            for s in _srows:
                surtax_rowdata.append([data["hs"]] + s)
                exp_surtax.append({"HS": data["hs"], "Period": s[0],
                                   "Tariff item": s[1], "Effective": s[2],
                                   "Surtax": s[3]})
            per_hs.append((data, code, _srows))

        ttkb.Label(self.results_area, text="Duties by treatment",
                   bootstyle="success", font=tkfont.Font(weight="bold")).pack(anchor="w")
        if value:
            duty_header = f"Duty on ${value:,.2f}"
        elif qty:
            duty_header = f"Duty ({qty:g} {qty_unit or ''})"
        else:
            duty_header = "Duty"
        Tableview(self.results_area,
                  coldata=[{"text": "HS", "width": 110, "stretch": False},
                           {"text": "Description", "stretch": True},
                           {"text": "Treatment", "width": 110, "stretch": False},
                           {"text": "Rate", "width": 110, "stretch": False},
                           {"text": duty_header, "width": 130, "stretch": False},
                           {"text": "Best", "width": 55, "stretch": False}],
                  rowdata=duty_rowdata,
                  paginated=True, pagesize=15, searchable=True,
                  bootstyle="primary", yscrollbar=True,
                  height=min(max(len(duty_rowdata), 3), 12)).pack(fill=X, expand=False,
                                                                  pady=(4, 12))

        ttkb.Label(self.results_area, text="US reciprocal surtax"
                   + (f" (in force {import_date.isoformat()})" if import_date
                      else " (all periods)"),
                   bootstyle="danger", font=tkfont.Font(weight="bold")).pack(anchor="w")
        if not is_us:
            ttkb.Label(self.results_area, text="N/A — origin is not the US.",
                       bootstyle="secondary").pack(anchor="w", pady=(4, 12))
        elif not surtax_rowdata:
            ttkb.Label(self.results_area,
                       text="These HS codes are not on the US surtax list — no surtax.",
                       bootstyle="secondary").pack(anchor="w", pady=(4, 12))
        else:
            Tableview(self.results_area,
                      coldata=[{"text": "HS", "width": 110, "stretch": False},
                               {"text": "Period", "width": 200, "stretch": False},
                               {"text": "Tariff item", "width": 110, "stretch": False},
                               {"text": "Effective", "width": 100, "stretch": False},
                               {"text": "Surtax", "width": 80, "stretch": False}],
                      rowdata=surtax_rowdata,
                      paginated=True, pagesize=10, searchable=False,
                      bootstyle="danger", height=min(max(len(surtax_rowdata), 3), 10)
                      ).pack(fill=X, expand=False, pady=(4, 12))

        ttkb.Label(self.results_area, text="Summary — one line per HS",
                   bootstyle="info", font=tkfont.Font(weight="bold")).pack(anchor="w")
        sum_table = build_summary_table(per_hs)
        Tableview(self.results_area,
                  coldata=[{"text": "HS", "width": 110, "stretch": False},
                           {"text": "Description", "stretch": True},
                           {"text": "Lowest", "width": 90, "stretch": False},
                           {"text": "Rate", "width": 90, "stretch": False},
                           {"text": "Duty", "width": 90, "stretch": False},
                           {"text": "GST", "width": 90, "stretch": False},
                           {"text": "Surtax", "width": 130, "stretch": False}],
                  rowdata=[[r["HS"], r["Description"], r["Lowest"], r["Rate"],
                            r["Duty"], r["GST"], r["Surtax"]] for r in sum_table],
                  paginated=False, searchable=False,
                  bootstyle="info", yscrollbar=True,
                  height=min(max(len(sum_table), 3), 10)).pack(fill=X, expand=False,
                                                                pady=(4, 12))
        self._last = {"meta": {"Origin": country,
                               "Value for duty": f"${value:,.2f}",
                               "Quantity": f"{qty:g} {qty_unit or ''}" if qty else "—",
                               "Import date": import_date.isoformat() if import_date
                               else "all periods",
                               "Tariff database": f"{self.rev} "
                               f"(eff {rev_info(self.con, self.rev).get('effective') or '?'})",
                               "Treatments": ", ".join(
                                   eligible_treatments(self.con, country, self.rev))},
                       "duties": exp_duties, "surtax": exp_surtax,
                       "summary": sum_table}
        self.status.configure(text=f"Done — {len(self.basket)} HS x {country}")


if __name__ == "__main__":
    if not DB_PATH.exists():
        raise SystemExit(f"Missing {DB_PATH} — keep tariff_app.db next to this script.")
    App().run()
