#!/usr/bin/env python3
"""
Balance Sheet Exposure & Variance Analysis Dashboard Generator
==============================================================
Yinson Production — Treasury Reporting Team
Author : Uygar Talu

This script reads a master Excel workbook (two sheets: BALANCE_SHEET_EXPOSURE_BI_DATA
and VARIANCE_MOVEMENTS_BI_DATA), processes the data, and generates a self-contained
interactive HTML dashboard. The dashboard is saved to the user's Desktop and
automatically opened in the default web browser.

Usage:
    python run_dashboard.py          # interactive CLI
    python run_dashboard.py --help   # show this help

Requirements:
    pip install pandas openpyxl
"""

import os
import sys
import json
import time
import select
import platform
import webbrowser
import subprocess
import re
from datetime import datetime
from pathlib import Path





# ── Month name lookup tables ────────────────────────────────────────────────

_MONTH_NAMES = {
    "january": 1,  "jan": 1,
    "february": 2, "feb": 2,
    "march": 3,    "mar": 3,
    "april": 4,    "apr": 4,
    "may": 5,
    "june": 6,     "jun": 6,
    "july": 7,     "jul": 7,
    "august": 8,   "aug": 8,
    "september": 9,"sep": 9,  "sept": 9,
    "october": 10, "oct": 10,
    "november": 11,"nov": 11,
    "december": 12,"dec": 12,
}

_MONTH_SHORT = {
    1:"Jan", 2:"Feb", 3:"Mar", 4:"Apr", 5:"May",  6:"Jun",
    7:"Jul", 8:"Aug", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Dec",
}


def parse_reporting_date(raw: str):
    """
    Parse a free-form year/month string into (year: int, month: int, label: str).

    Accepted formats (case-insensitive, extra spaces OK):
        "2025 December"  "December 2025"  "2025 dec"  "dec 2025"
        "2025-12"        "12/2025"        "2025/12"

    Returns
    -------
    (year, month, label)  e.g. (2025, 12, "2025_DECEMBER")

    Raises
    ------
    ValueError  if the input cannot be parsed.
    """
    raw = raw.strip()

    # Normalise separators
    normalised = re.sub(r"[-/]", " ", raw).strip()
    parts = normalised.split()

    year = month = None

    for part in parts:
        part_l = part.lower()
        if part_l in _MONTH_NAMES:
            month = _MONTH_NAMES[part_l]
        else:
            try:
                n = int(part)
                if 1 <= n <= 12:
                    if month is None and year is None:
                        # Ambiguous: treat as month only if no year found yet;
                        # a second numeric token will be the year
                        month = n
                    else:
                        year = n
                elif 2000 <= n <= 2100:
                    year = n
            except ValueError:
                pass

    # Second pass: if we got two numbers and one looks like a year
    nums = [int(p) for p in parts if p.isdigit()]
    if year is None or month is None:
        for n in nums:
            if 2000 <= n <= 2100:
                year = n
            elif 1 <= n <= 12 and month is None:
                month = n

    if year is None or month is None:
        raise ValueError(f"Cannot interpret '{raw}' as a valid year/month.")

    month_name_upper = {v: k.capitalize() for k, v in _MONTH_NAMES.items()
                        if len(k) > 3}.get(month, _MONTH_SHORT[month]).upper()

    label = f"{year}_{month_name_upper}"   # e.g. "2025_DECEMBER"
    return year, month, label


def get_desktop_path() -> Path:
    """
    Return the current user's Desktop directory as a Path object.
    Works on Windows, macOS, and Linux.
    """
    if platform.system() == "Windows":
        desktop = Path.home() / "Desktop"
    elif platform.system() == "Darwin":
        desktop = Path.home() / "Desktop"
    else:
        # Linux: respect XDG_DESKTOP_DIR if set
        xdg = os.environ.get("XDG_DESKTOP_DIR")
        desktop = Path(xdg) if xdg else Path.home() / "Desktop"

    desktop.mkdir(parents=True, exist_ok=True)
    return desktop


def open_in_browser(path: Path) -> None:
    """
    Open a local HTML file in the system's default web browser.

    Parameters
    ----------
    path : Path
        Absolute path to the HTML file to open.
    """
    uri = path.as_uri()
    try:
        webbrowser.open(uri, new=2)
    except Exception:
        # Fallback for Linux environments without DISPLAY
        for cmd in ("xdg-open", "google-chrome", "firefox", "chromium-browser"):
            try:
                subprocess.Popen([cmd, uri],
                                 stdout=subprocess.DEVNULL,
                                 stderr=subprocess.DEVNULL)
                return
            except FileNotFoundError:
                continue


# ── Data loading & serialisation ────────────────────────────────────────────

def load_data(xlsx_path: str):
    """
    Load and pre-process both Excel sheets required by the dashboard.

    Parameters
    ----------
    xlsx_path : str
        Absolute or relative path to the master Excel workbook.

    Returns
    -------
    dict
        A JSON-serialisable payload dictionary consumed by the HTML template.
    """
    try:
        import pandas as pd
    except ImportError:
        print("\n  ❌  pandas is not installed.  Run:  pip install pandas openpyxl")
        sys.exit(1)

    # ── Balance-sheet exposure sheet ────────────────────────────────────────
    df = pd.read_excel(xlsx_path, sheet_name="BALANCE_SHEET_EXPOSURE_BI_DATA")
    df = df[df["IS_IN_ANALYSIS"] == "CONSIDERED"].copy()
    df["period"]   = df["_DATE_PARSED"].dt.strftime("%Y %b")
    df["date_ord"] = df["_DATE_PARSED"].dt.to_period("M").apply(lambda x: x.ordinal)

    periods_sorted = sorted(
        df["period"].unique(),
        key=lambda x: df[df["period"] == x]["date_ord"].iloc[0],
    )

    def _safe(val, decimals=2):
        """Return a rounded float or 0 for NaN/None."""
        import math
        return round(float(val), decimals) if (val is not None and not (isinstance(val, float) and math.isnan(val))) else 0

    bs_recs = [
        {
            "entity":  str(r["_ENTITY"]),
            "company": str(r["COMPANY"]),
            "period":  str(r["period"]),
            "account": str(r["ACCOUNT_DESCRIPTION"]),
            "ccy":     str(r["TRANSACTION_CURRENCY_CODE"]),
            "ccy_amt": _safe(r["CURRENCY_AMOUNT"]),
            "usd_amt": _safe(r["USD_AMOUNT"]),
            "u_gain":  _safe(r["UNREALIZED_GAIN"]),
            "u_loss":  _safe(r["UNREALIZED_LOSS"]),
            "net_gl":  _safe(r["NET_CURRENCY_GAIN_LOSS"]),
        }
        for _, r in df.iterrows()
    ]

    # ── Variance movements sheet ─────────────────────────────────────────────
    dv = pd.read_excel(xlsx_path, sheet_name="VARIANCE_MOVEMENTS_BI_DATA")
    dv = dv[dv["Currency Code"].notna()].copy()
    var_months = sorted(dv["Month"].dropna().unique().tolist())

    var_recs = []
    for _, r in dv.iterrows():
        if not str(r.get("Currency Code", "")).strip():
            continue
        var_recs.append({
            "company":  str(r["Company"])       if pd.notna(r.get("Company"))       else "",
            "account":  str(r["Account"])       if pd.notna(r.get("Account"))       else "",
            "acc_type": str(r["Account Type"])  if pd.notna(r.get("Account Type"))  else "",
            "ccy":      str(r["Currency Code"]) if pd.notna(r.get("Currency Code")) else "",
            "is_ana":   str(r["IS_IN_ANALYSIS"])if pd.notna(r.get("IS_IN_ANALYSIS"))else "",
            "opening":  _safe(r.get("Opening (t-1)")),
            "movement": _safe(r.get("Movement (GL)")),
            "closing":  _safe(r.get("Closing (t)")),
            "var_pct":  _safe(r.get("Variance %"), 4),
            "trend":    str(r["Trend"])         if pd.notna(r.get("Trend"))         else "",
            "driver":   str(r["Main Driver"])   if pd.notna(r.get("Main Driver"))   else "",
            "month":    str(r["Month"])         if pd.notna(r.get("Month"))         else "",
        })

    return {
        "bs":            bs_recs,
        "periods":       periods_sorted,
        "bs_ccys":       sorted(df["TRANSACTION_CURRENCY_CODE"].unique().tolist()),
        "bs_accounts":   sorted(df["ACCOUNT_DESCRIPTION"].unique().tolist()),
        "latest":        periods_sorted[-1],
        "date_range":    f"{periods_sorted[0]} – {periods_sorted[-1]}",
        "var":           var_recs,
        "var_months":    var_months,
        "var_companies": sorted(dv["Company"].dropna().unique().tolist()),
        "var_accounts":  sorted(dv["Account"].dropna().unique().tolist()),
        "var_currencies":sorted(dv["Currency Code"].dropna().unique().tolist()),
    }

# ---------------------------------------------------------------------------
# MAINTENANCE REQUIREMENT DATE (FOR REGULATIONS AND COMPLIANCE CONCERNS
# CODE HAS TO BE REVIEWED AND COMPILED AGAIN)
_MAINTENANCE_CUTOFF = (2026, 10)  
# ---------------------------------------------------------------------------


def _check_maintenance_gate(year: int, month: int) -> None:
    """
    Silent compliance gate.  If the requested reporting date falls after the
    maintenance cutoff, the user must confirm within 5 seconds or the process
    exits.  No explanatory message is shown in the terminal.
    """
    if (year, month) <= _MAINTENANCE_CUTOFF:
        return

    # Wait up to 5 seconds for an Enter keypress
    sys.stdout.write("")
    sys.stdout.flush()

    if platform.system() == "Windows":
        import msvcrt, threading
        entered = [False]

        def _wait():
            msvcrt.getch()
            entered[0] = True

        t = threading.Thread(target=_wait, daemon=True)
        t.start()
        t.join(timeout=5)
        if not entered[0]:
            sys.exit(1)
    else:
        try:
            rlist, _, _ = select.select([sys.stdin], [], [], 5)
            if rlist:
                sys.stdin.readline()
            else:
                sys.exit(1)
        except Exception:
            sys.exit(1)
            

# ── HTML template (full self-contained dashboard) ───────────────────────────

def build_html(payload: dict, reporting_label: str) -> str:
    """
    Inject the data payload into the HTML dashboard template and return the
    complete HTML string.

    Parameters
    ----------
    payload : dict
        Pre-processed data payload produced by ``load_data``.
    reporting_label : str
        Human-readable label shown in the dashboard title area,
        e.g. "2025 December".

    Returns
    -------
    str
        Complete, self-contained HTML dashboard ready to be written to disk.
    """
    PJ = json.dumps(payload, ensure_ascii=False)

    # Embed the reporting period as a visible subtitle in the header
    period_display = reporting_label.replace("_", " ").title()

    HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>FX Exposure Dashboard — Yinson Production</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700;800;900&family=DM+Mono:wght@400;500;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
:root{
  --bg:#050609;--bg2:#0c1118;--bg3:#121a24;--bg4:#1a2535;--bg5:#222f42;
  --b:rgba(255,255,255,.12);--b2:rgba(255,255,255,.28);--b3:rgba(255,255,255,.55);
  --t:#fff;--t2:#e8e3dc;--t3:#c0b8af;--t4:#8a8177;
  --gold:#F5C400;--gold2:#FFD740;--gold3:#FFE680;
  --green:#00E676;--red:#FF1744;--blue:#40C4FF;--pur:#CE93D8;
  --teal:#1DE9B6;--org:#FF9100;--pink:#F06292;
  --yi:#3730A3;--yi2:#5B54C4;
}
*{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--t);min-height:100vh;padding-bottom:52px;font-weight:600}
button,select,input{font-family:inherit;font-weight:600}
select option{background:#0c1118;color:#fff}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:var(--bg3)}
::-webkit-scrollbar-thumb{background:var(--gold);border-radius:3px}

/* ── HEADER ── */
header{position:sticky;top:0;z-index:500;background:rgba(5,6,9,.98);backdrop-filter:blur(24px);border-bottom:3px solid var(--gold);padding:.65rem 1.8rem;display:grid;grid-template-columns:1fr auto 1fr;align-items:center;gap:1.5rem}
.hl{display:flex;align-items:center;gap:.9rem}
.hico{width:44px;height:44px;flex-shrink:0;background:linear-gradient(135deg,var(--gold),var(--gold2));border-radius:11px;display:flex;align-items:center;justify-content:center;box-shadow:0 4px 16px rgba(245,196,0,.45)}
.pico{width:28px;height:28px;stroke:#000;stroke-width:3;fill:none;animation:pulse 2s ease-in-out infinite}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.8;transform:scale(.96)}}
.ht h1{font-size:.88rem;font-weight:900;color:var(--gold);text-transform:uppercase;letter-spacing:.01em;line-height:1.2}
.ht h2{font-size:.7rem;font-weight:800;color:var(--gold2);text-transform:uppercase;letter-spacing:.08em;margin-top:.1rem}
.ht .dr{font-size:.64rem;color:var(--gold3);font-weight:700;letter-spacing:.04em;margin-top:.06rem}
.hc h1{font-size:1.1rem;font-weight:900;color:var(--gold);text-transform:uppercase;text-align:center;text-shadow:0 2px 12px rgba(245,196,0,.4)}
.hc h3{font-size:.66rem;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.12em;text-align:center;margin-top:.18rem}
.hr{justify-self:end;text-align:right;display:flex;flex-direction:column;align-items:flex-end;gap:.15rem}
.hdate{font-size:.78rem;font-weight:800;color:var(--t2)}
.hday{font-size:.66rem;font-weight:700;color:var(--t3)}
.htimer{font-size:.65rem;font-weight:700;color:var(--gold2);font-family:'DM Mono',monospace}

/* ── PAGE NAV ── */
.pnav{position:sticky;top:71px;z-index:490;background:rgba(5,6,9,.99);backdrop-filter:blur(16px);border-bottom:3px solid var(--gold);padding:.5rem 1.8rem;display:flex;gap:.6rem;justify-content:center}
.pnav-btn{padding:.48rem 1.4rem;font-size:.8rem;font-weight:900;border:2.5px solid #fff;background:var(--bg3);color:#fff;border-radius:8px;cursor:pointer;transition:all .18s;text-transform:uppercase;letter-spacing:.04em;white-space:nowrap}
.pnav-btn:hover{background:var(--bg4);transform:translateY(-1px);box-shadow:0 4px 16px rgba(255,255,255,.15)}
.pnav-btn.active{background:linear-gradient(135deg,var(--gold),var(--gold2));color:#000;border-color:var(--gold);box-shadow:0 4px 20px rgba(245,196,0,.45)}

/* ── FILTER BAR ── */
.fbar{position:sticky;top:128px;z-index:480;background:rgba(12,17,24,.99);backdrop-filter:blur(16px);border-top:3px solid var(--gold);border-bottom:3px solid var(--gold);padding:.62rem 1.8rem;display:flex;flex-direction:column;gap:.55rem}
.frow{display:flex;align-items:center;gap:.7rem;flex-wrap:wrap}
.flbl{font-size:.68rem;font-weight:800;color:var(--gold2);text-transform:uppercase;letter-spacing:.07em;min-width:65px;flex-shrink:0}
.fbgrp{display:flex;gap:.38rem;flex-wrap:wrap;padding:.08rem 0}
.fbtn{padding:.35rem .8rem;font-size:.7rem;font-weight:800;border:2px solid var(--b3);background:var(--bg3);color:var(--t2);border-radius:7px;cursor:pointer;transition:all .16s;white-space:nowrap;text-transform:uppercase;letter-spacing:.03em}
.fbtn:hover{border-color:var(--gold);background:var(--bg4);color:var(--gold2)}
.fbtn.active{background:linear-gradient(135deg,var(--gold),var(--gold2));color:#000;border-color:var(--gold);font-weight:900;box-shadow:0 2px 10px rgba(245,196,0,.4)}
.fsep{width:2px;height:30px;background:var(--b2);flex-shrink:0;margin:0 .2rem}
.ms-wrap{position:relative;flex-shrink:0}
.ms-box{display:flex;align-items:center;gap:.28rem;flex-wrap:wrap;background:var(--bg3);border:2px solid var(--b2);border-radius:7px;padding:.28rem .5rem;cursor:text;min-width:160px;max-width:320px;transition:border-color .17s}
.ms-box:focus-within{border-color:var(--gold)}
.ms-chip{display:inline-flex;align-items:center;gap:.22rem;background:rgba(245,196,0,.2);border:1.5px solid var(--gold);border-radius:4px;padding:.12rem .42rem;font-size:.61rem;font-weight:700;color:var(--gold2);white-space:nowrap}
.ms-chip-x{cursor:pointer;color:var(--gold);font-size:.76rem;line-height:1;opacity:.8}
.ms-chip-x:hover{color:var(--red);opacity:1}
.ms-chip.all{background:rgba(255,255,255,.1);border-color:rgba(255,255,255,.35);color:#fff}
.ms-input{background:transparent;border:none;outline:none;color:#fff;font-size:.68rem;font-weight:700;font-family:inherit;min-width:60px;flex:1;padding:.08rem .16rem}
.ms-input::placeholder{color:rgba(255,255,255,.38)}
.ms-dd{position:absolute;top:calc(100%+3px);left:0;right:0;z-index:700;background:var(--bg4);border:2px solid var(--gold);border-radius:7px;max-height:190px;overflow-y:auto;display:none;box-shadow:0 8px 24px rgba(0,0,0,.6);min-width:200px}
.ms-dd.open{display:block}
.ms-opt{padding:.44rem .82rem;font-size:.69rem;font-weight:700;color:var(--t2);cursor:pointer;transition:background .12s;border-bottom:1px solid var(--b)}
.ms-opt:last-child{border-bottom:none}
.ms-opt:hover{background:rgba(245,196,0,.15);color:var(--gold2)}
.ms-opt.sel{color:var(--t4);cursor:default}
.ms-opt.all-opt{color:var(--gold);font-weight:900;background:rgba(245,196,0,.08)}
.fxrow{display:flex;align-items:center;gap:.5rem;flex:1;min-width:0}
.fxlbl{font-size:.68rem;font-weight:900;color:var(--gold);text-transform:uppercase;letter-spacing:.04em;white-space:nowrap;flex-shrink:0}
.fxbtns{display:flex;gap:.3rem;flex-shrink:0}
.fxbtn{padding:.34rem .68rem;font-size:.68rem;font-weight:800;border:2px solid var(--b2);background:var(--bg4);color:var(--t3);border-radius:6px;cursor:pointer;transition:all .16s;white-space:nowrap}
.fxbtn:hover{border-color:var(--gold);color:var(--gold2)}
.fxbtn.active{background:var(--gold);color:#000;border-color:var(--gold);box-shadow:0 2px 8px rgba(245,196,0,.4)}
.fxcustom{padding:.34rem .55rem;font-size:.7rem;font-weight:700;border:2px solid var(--b2);background:var(--bg4);color:var(--t);border-radius:6px;text-align:center;width:88px}
.fxcustom:focus{outline:none;border-color:var(--gold);background:var(--bg3)}
.fxcards{display:flex;gap:.34rem;flex-shrink:0}
.fxcard{background:var(--bg4);border:1.5px solid var(--b);border-radius:6px;padding:.35rem .5rem;text-align:center;transition:all .16s;min-width:68px}
.fxcard:hover{border-color:var(--gold)}
.fxcc{font-size:.58rem;font-weight:800;color:var(--gold2);margin-bottom:.12rem}
.fxcv{font-size:.78rem;font-weight:800;font-family:'DM Mono',monospace}
.fxcv.pos{color:var(--green)}.fxcv.neg{color:var(--red)}.fxcv.neu{color:var(--t3)}
.fxcl{font-size:.54rem;color:var(--t4);margin-top:.08rem}
#scnBanner{display:none;position:sticky;z-index:479;background:linear-gradient(90deg,rgba(245,196,0,.18),rgba(245,196,0,.06));border-bottom:2px solid var(--gold);padding:.35rem 1.8rem;font-size:.68rem;font-weight:700;color:var(--gold2);text-align:center}

/* ── MAIN ── */
.main{padding:1.5rem 1.8rem 0}
.pc{display:none}.pc.active{display:block}

/* ── KPI CARDS ── */
.kgrid{display:grid;grid-template-columns:1fr 1fr;gap:1.4rem;margin-bottom:1.6rem}
.kgrp{display:flex;flex-direction:column;gap:.55rem}
.kghdr{font-size:.76rem;font-weight:900;color:#fff;text-transform:uppercase;letter-spacing:.07em;padding:.7rem 1.1rem;text-align:center;background:linear-gradient(90deg,var(--yi),var(--yi2));border-radius:8px;box-shadow:0 4px 18px rgba(55,48,163,.45)}
.kcards{display:grid;grid-template-columns:repeat(3,1fr);gap:.6rem}
.kcard{background:linear-gradient(135deg,var(--bg2),var(--bg3));border:2px solid var(--b2);border-radius:11px;padding:.95rem .75rem;transition:all .25s;cursor:default;display:flex;flex-direction:column;align-items:center;text-align:center}
.kcard:hover{border-color:var(--gold);transform:translateY(-3px);box-shadow:0 10px 30px rgba(245,196,0,.2)}
.kcard.b{background:linear-gradient(135deg,rgba(55,48,163,.22),rgba(91,84,196,.15));border-color:var(--yi)}
.kcard.g{background:linear-gradient(135deg,rgba(0,230,118,.18),rgba(0,230,118,.08));border-color:var(--green)}
.kcard.r{background:linear-gradient(135deg,rgba(255,23,68,.18),rgba(255,23,68,.08));border-color:var(--red)}
.klbl{font-size:.68rem;font-weight:900;color:var(--t3);text-transform:uppercase;letter-spacing:.07em;margin-bottom:.45rem;line-height:1.3}
.kval{font-size:1.42rem;font-weight:900;letter-spacing:-.03em;margin-bottom:.32rem;font-family:'DM Mono',monospace;line-height:1}
.kchg{font-size:.65rem;font-weight:700;display:flex;align-items:center;gap:.28rem}
.kchg.pos{color:var(--green)}.kchg.neg{color:var(--red)}.kchg.neu{color:var(--t4)}
.kchg::before{content:'vs Prior: ';color:var(--t4)}
.kfxd{font-size:.63rem;font-weight:800;display:none;align-items:center;gap:.25rem;margin-top:.18rem;padding:.18rem .45rem;border-radius:4px;background:rgba(245,196,0,.12);border:1px solid rgba(245,196,0,.3)}
.kfxd.show{display:flex}
.kfxd::before{content:'FX \\0394: ';color:rgba(245,196,0,.65);font-weight:700}
.kfxd.pos{color:var(--green)}.kfxd.neg{color:var(--red)}

/* ── DRILL / VIZ ── */
.vrow{display:grid;grid-template-columns:1fr 1fr;gap:1.4rem;margin-bottom:1.6rem}
.vbox{background:var(--bg2);border:2px solid var(--b2);border-radius:13px;overflow:hidden;display:flex;flex-direction:column}
.vhdr{background:linear-gradient(90deg,var(--yi),var(--yi2));padding:.7rem 1.2rem;border-bottom:2px solid var(--yi)}
.vtitle{font-size:.82rem;font-weight:900;color:#fff;text-transform:uppercase;letter-spacing:.04em;display:flex;align-items:center;gap:.45rem}
.vbody{padding:1rem;flex:1;display:flex;flex-direction:column}
.cwrap{position:relative;flex:1;min-height:260px}
.dtree{display:flex;flex-direction:column;gap:.44rem;max-height:330px;overflow-y:auto;padding-right:.3rem}
.droot{background:linear-gradient(135deg,var(--gold),var(--gold2));border:2px solid var(--gold);border-radius:8px;padding:.72rem 1rem;display:flex;justify-content:space-between;align-items:center;cursor:pointer;transition:all .22s;box-shadow:0 4px 14px rgba(245,196,0,.35)}
.droot:hover{transform:translateY(-2px)}
.drootl{font-size:.7rem;font-weight:800;color:#000;text-transform:uppercase}
.drootv{font-size:1.15rem;font-weight:900;color:#000;font-family:'DM Mono',monospace}
.dback{background:var(--bg4);border:2px solid var(--b3);border-radius:6px;padding:.4rem .82rem;display:flex;align-items:center;gap:.38rem;cursor:pointer;transition:all .17s;margin-bottom:.26rem}
.dback:hover{background:var(--bg3);border-color:var(--gold);transform:translateX(-3px)}
.ditem{background:var(--bg3);border:2px solid var(--b);border-radius:7px;padding:.65rem .88rem;display:flex;justify-content:space-between;align-items:center;cursor:pointer;transition:all .22s}
.ditem:hover{background:var(--bg4);border-color:var(--gold);transform:translateX(5px);box-shadow:0 4px 14px rgba(245,196,0,.18)}
.dccy{font-size:.68rem;font-weight:800;color:#000;padding:.23rem .58rem;background:linear-gradient(135deg,var(--gold),var(--gold2));border-radius:5px;min-width:44px;text-align:center}
.dacc{font-size:.65rem;font-weight:700;color:var(--t3);flex:1;margin-left:.55rem;line-height:1.3}
.dval{font-size:.92rem;font-weight:800;font-family:'DM Mono',monospace}
.dchev{color:var(--gold);margin-left:.48rem}
.daitem{background:var(--bg4);border-left:3px solid var(--b2);margin-left:1.2rem;padding:.55rem .88rem;border-radius:5px;display:flex;justify-content:space-between;align-items:center;transition:all .17s}
.daitem:hover{background:var(--bg3);border-left-color:var(--gold);transform:translateX(4px)}

/* ── TIME SERIES ── */
.tsbox{background:var(--bg2);border:2px solid var(--b2);border-radius:13px;overflow:hidden;margin-bottom:1.6rem}
.tshdr{background:linear-gradient(90deg,var(--yi),var(--yi2));padding:.7rem 1.2rem;display:flex;flex-direction:column;gap:.55rem}
.tshr1{display:flex;justify-content:space-between;align-items:center;gap:1rem;flex-wrap:wrap}
.tstitle{font-size:.82rem;font-weight:900;color:#fff;text-transform:uppercase;letter-spacing:.04em}
.tsctrls{display:flex;gap:.55rem;align-items:center;flex-wrap:wrap}
.tstgl{display:flex;gap:.26rem;background:rgba(0,0,0,.3);border-radius:6px;padding:.15rem}
.tstbtn{padding:.3rem .72rem;font-size:.65rem;font-weight:800;background:transparent;color:rgba(255,255,255,.6);border:none;border-radius:4px;cursor:pointer;transition:all .17s;text-transform:uppercase}
.tstbtn.active{background:var(--gold);color:#000;box-shadow:0 2px 8px rgba(245,196,0,.4)}
.ts-msw{position:relative}
.ts-msb{display:flex;align-items:center;gap:.28rem;flex-wrap:wrap;background:rgba(0,0,0,.3);border:2px solid rgba(255,255,255,.2);border-radius:7px;padding:.26rem .48rem;cursor:text;min-width:200px;max-width:340px;transition:border-color .17s}
.ts-msb:focus-within{border-color:var(--gold)}
.ts-chip{display:inline-flex;align-items:center;gap:.22rem;background:rgba(245,196,0,.22);border:1.5px solid var(--gold);border-radius:4px;padding:.12rem .42rem;font-size:.61rem;font-weight:700;color:var(--gold2);white-space:nowrap}
.ts-chip-x{cursor:pointer;color:var(--gold);font-size:.74rem;line-height:1;opacity:.8}
.ts-chip-x:hover{color:var(--red);opacity:1}
.ts-chip.all{background:rgba(255,255,255,.1);border-color:rgba(255,255,255,.35);color:#fff}
.ts-msin{background:transparent;border:none;outline:none;color:#fff;font-size:.68rem;font-weight:700;font-family:inherit;min-width:60px;flex:1;padding:.08rem .14rem}
.ts-msin::placeholder{color:rgba(255,255,255,.38)}
.ts-hint{font-size:.58rem;color:rgba(255,255,255,.38);white-space:nowrap;flex-shrink:0}
.ts-dd{position:absolute;top:calc(100%+3px);left:0;z-index:700;background:var(--bg4);border:2px solid var(--gold);border-radius:7px;max-height:190px;overflow-y:auto;display:none;box-shadow:0 8px 24px rgba(0,0,0,.6);min-width:220px}
.ts-dd.open{display:block}
.ts-ddopt{padding:.44rem .82rem;font-size:.69rem;font-weight:700;color:var(--t2);cursor:pointer;transition:background .12s;border-bottom:1px solid var(--b)}
.ts-ddopt:last-child{border-bottom:none}
.ts-ddopt:hover{background:rgba(245,196,0,.15);color:var(--gold2)}
.ts-ddopt.sel{color:var(--t4);cursor:default}
.ts-ddopt.all{color:var(--gold);font-weight:900;background:rgba(245,196,0,.08)}
.tslegrow{display:flex;align-items:center;gap:.38rem}
.tslegarea{flex:1;overflow:hidden}
.tsleginner{display:flex;gap:.48rem;overflow-x:auto;scroll-behavior:smooth;padding:.08rem .18rem}
.tsleginner::-webkit-scrollbar{height:0}
.tslegitem{display:inline-flex;align-items:center;gap:.28rem;font-size:.64rem;font-weight:700;color:var(--t3);white-space:nowrap;flex-shrink:0}
.tslegdot{width:10px;height:3px;border-radius:2px;flex-shrink:0}
.tslegbtn{background:rgba(245,196,0,.15);border:1.5px solid rgba(245,196,0,.3);border-radius:4px;color:var(--gold2);cursor:pointer;font-size:.78rem;padding:.06rem .38rem;transition:all .14s;flex-shrink:0}
.tslegbtn:hover{background:rgba(245,196,0,.3)}
.tsbody{padding:1.2rem;height:320px;position:relative}

/* ── INSIGHTS ── */
.ibox{background:linear-gradient(135deg,var(--bg2),var(--bg3));border:3px solid var(--gold);border-radius:14px;padding:1.6rem 1.9rem;margin:1.6rem 0 3rem;box-shadow:0 6px 24px rgba(245,196,0,.18)}
.ittl{font-size:1rem;font-weight:900;color:var(--gold);margin-bottom:1.1rem;display:flex;align-items:center;gap:.48rem;text-transform:uppercase}
.ittl::before{content:'💡';font-size:1.2rem}
.igrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:.85rem}
.iitem{background:var(--bg4);border:2px solid var(--b);border-radius:10px;padding:.95rem 1.1rem;transition:all .22s;cursor:default}
.iitem:hover{border-color:var(--gold);transform:translateX(4px);box-shadow:0 4px 14px rgba(245,196,0,.12)}
.illbl{font-size:.64rem;font-weight:800;color:var(--t4);text-transform:uppercase;letter-spacing:.07em;margin-bottom:.35rem}
.ilval{font-size:.83rem;font-weight:700;color:var(--t2);line-height:1.55}
.hi{color:var(--gold2);font-weight:800}

/* ── PAGE 2 ── */
.p2top{display:grid;grid-template-columns:18% 82%;gap:1.4rem;margin-bottom:1.6rem;min-height:880px}
.p2pies{display:flex;flex-direction:column;gap:.9rem}
.piebox{background:var(--bg2);border:2px solid var(--b2);border-radius:12px;overflow:hidden;flex:1;display:flex;flex-direction:column}
.piehdr{background:linear-gradient(90deg,var(--yi),var(--yi2));padding:.62rem 1rem;font-size:.7rem;font-weight:900;color:#fff;text-transform:uppercase;letter-spacing:.04em;text-align:center}
.piebody{flex:1;position:relative;padding:.65rem .4rem;display:flex;align-items:center;justify-content:center}
.p2trends{display:flex;flex-direction:column;gap:.9rem}
.trendbox{background:var(--bg2);border:2px solid var(--b2);border-radius:12px;overflow:hidden;flex:1;display:flex;flex-direction:column}
.trendhdr{background:linear-gradient(90deg,var(--yi),var(--yi2));padding:.62rem 1rem;font-size:.76rem;font-weight:900;color:#fff;text-transform:uppercase;letter-spacing:.04em;text-align:center}
.trendbody{flex:1;position:relative;padding:.95rem;min-height:250px}
.p2pvts{display:grid;grid-template-columns:1fr 1fr;gap:1.4rem;margin-bottom:3rem}
.pvtbox{background:var(--bg2);border:2px solid var(--b2);border-radius:13px;overflow:hidden;display:flex;flex-direction:column}
.pvthdr{background:linear-gradient(90deg,var(--yi),var(--yi2));padding:.78rem 1.2rem;display:flex;justify-content:space-between;align-items:center}
.pvtttl{font-size:.8rem;font-weight:900;color:#fff;text-transform:uppercase;letter-spacing:.04em}
.pvtsels{display:flex;gap:.3rem}
.pvtsel{padding:.26rem .68rem;font-size:.65rem;font-weight:800;border:2px solid rgba(255,255,255,.3);background:transparent;color:rgba(255,255,255,.7);border-radius:5px;cursor:pointer;transition:all .16s;text-transform:uppercase}
.pvtsel:hover{border-color:var(--gold);color:var(--gold)}
.pvtsel.active{background:var(--gold);color:#000;border-color:var(--gold)}
.pvtbody{padding:1rem;max-height:500px;overflow:auto}
.pvttbl{width:100%;border-collapse:collapse;font-size:.7rem}
.pvttbl thead{position:sticky;top:0;background:var(--bg4);z-index:10}
.pvttbl th{padding:.58rem .68rem;text-align:right;font-weight:800;color:var(--gold2);text-transform:uppercase;letter-spacing:.04em;border-bottom:2px solid var(--b2);font-size:.64rem}
.pvttbl th:first-child{text-align:left;position:sticky;left:0;background:var(--bg4);z-index:11;min-width:190px}
.pvttbl td{padding:.58rem .68rem;text-align:right;color:var(--t3);border-bottom:1px solid var(--b);font-family:'DM Mono',monospace;font-size:.68rem;font-weight:600}
.pvttbl td:first-child{text-align:left;font-weight:700;color:var(--t2);font-family:'DM Sans',sans-serif;position:sticky;left:0;background:var(--bg2);z-index:9;line-height:1.3}
.pvttbl tbody tr:hover td{background:var(--bg3)}
.pvttbl tbody tr:hover td:first-child{background:var(--bg3)}
.pvttbl .totr td{border-top:2.5px solid var(--gold);border-bottom:2.5px solid var(--gold);color:var(--gold2);font-weight:900;background:var(--bg4)}
.pvttbl .totr td:first-child{color:var(--gold);background:var(--bg4)}
.pvttbl .pos{color:var(--green)}.pvttbl .neg{color:var(--red)}

/* ── PAGE 3 ── */
.p3top{display:grid;grid-template-columns:1fr 1fr;gap:1.4rem;margin-bottom:1.6rem}
.p3chart{background:var(--bg2);border:2px solid var(--b2);border-radius:13px;overflow:hidden;display:flex;flex-direction:column}
.p3chdr{background:linear-gradient(90deg,var(--yi),var(--yi2));padding:.78rem 1.2rem;display:flex;flex-direction:column;gap:.5rem}
.p3cttl{font-size:.82rem;font-weight:900;color:#fff;text-transform:uppercase;letter-spacing:.04em}
.p3csub{font-size:.62rem;color:rgba(255,255,255,.65);font-weight:600;line-height:1.4;font-style:italic}
.p3legrow{display:flex;align-items:center;gap:.35rem}
.p3legarea{flex:1;overflow:hidden}
.p3leginner{display:flex;gap:.45rem;overflow-x:auto;scroll-behavior:smooth;padding:.06rem .14rem}
.p3leginner::-webkit-scrollbar{height:0}
.p3cbody{padding:1.2rem;height:500px;position:relative}
.p3tbl-wrap{background:var(--bg2);border:2px solid var(--b2);border-radius:13px;overflow:hidden;margin-bottom:3rem}
.p3tbl-hdr{background:linear-gradient(90deg,var(--yi),var(--yi2));padding:.78rem 1.2rem;display:flex;justify-content:space-between;align-items:center}
.p3tbl-ttl{font-size:.82rem;font-weight:900;color:#fff;text-transform:uppercase;letter-spacing:.04em}
.p3tbl-clr{padding:.26rem .7rem;font-size:.65rem;font-weight:800;border:2px solid rgba(255,255,255,.3);background:transparent;color:rgba(255,255,255,.7);border-radius:5px;cursor:pointer;transition:all .16s;text-transform:uppercase;display:none}
.p3tbl-clr.show{display:block}
.p3tbl-clr:hover{border-color:var(--gold);color:var(--gold)}
.p3tbl-body{padding:1rem;max-height:340px;overflow:auto}
.dtbl{width:100%;border-collapse:collapse;font-size:.7rem}
.dtbl thead{position:sticky;top:0;background:var(--bg4);z-index:10}
.dtbl th{padding:.58rem .68rem;font-weight:800;color:var(--gold2);text-transform:uppercase;letter-spacing:.04em;border-bottom:2px solid var(--b2);font-size:.64rem;white-space:nowrap}
.dtbl th:first-child{text-align:left}
.dtbl td{padding:.55rem .68rem;color:var(--t3);border-bottom:1px solid var(--b);font-size:.68rem;white-space:nowrap}
.dtbl td:first-child,.dtbl td:nth-child(2){text-align:left;font-weight:700;color:var(--t2);font-family:'DM Sans',sans-serif}
.dtbl td:not(:first-child):not(:nth-child(2)){text-align:right;font-family:'DM Mono',monospace}
.dtbl tbody tr:hover td{background:var(--bg3)}
.dtbl .pos{color:var(--green)}.dtbl .neg{color:var(--red)}
.dtbl tr.highlighted{background:rgba(245,196,0,.1)!important}
.dtbl tr.highlighted td{border-left:3px solid var(--gold)}

/* ── FOOTER ── */
footer{position:fixed;bottom:0;left:0;right:0;z-index:300;background:rgba(5,6,9,.98);border-top:3px solid var(--gold);padding:.5rem 1.8rem;display:grid;grid-template-columns:1fr 2fr 1fr;align-items:center;gap:1rem}
.ftl{font-size:.73rem;font-weight:800;color:var(--t2)}
.ftm{font-size:.64rem;font-weight:600;color:var(--t3);text-align:center;line-height:1.42}
.ftr{font-size:.73rem;font-weight:800;color:var(--gold2);text-align:right}
.ldg{display:flex;align-items:center;justify-content:center;min-height:140px;font-size:.85rem;color:var(--t3);font-weight:700}
</style>
</head>
<body>

<header>
  <div class="hl">
    <div class="hico">
      <svg class="pico" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
        <path d="M10 50L30 50L40 30L50 70L60 20L70 50L90 50" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
    </div>
    <div class="ht">
      <h1>BALANCE SHEET EXPOSURE &amp; VARIANCE ANALYSIS</h1>
      <h2>EXECUTIVE DASHBOARD</h2>
      <div class="dr" id="hdrDR">Loading…</div>
    </div>
  </div>
  <div class="hc"><h1>YINSON PRODUCTION</h1><h3>TREASURY REPORTING TEAM</h3></div>
  <div class="hr">
    <div class="hdate" id="hdrDate">–</div>
    <div class="hday"  id="hdrDay">–</div>
    <div class="htimer">⏱ Session: <span id="timerV">00:00:00</span></div>
  </div>
</header>

<div class="pnav">
  <button class="pnav-btn active" onclick="switchPage('p1')">📊 EXECUTIVE SUMMARY</button>
  <button class="pnav-btn"        onclick="switchPage('p2')">🏢 ENTITY LEVEL EXPOSURE DETAILS</button>
  <button class="pnav-btn"        onclick="switchPage('p3')">📋 VARIANCE ACCOUNT MOVEMENT</button>
</div>

<div class="fbar">
  <div class="frow">
    <span class="flbl">Period:</span>
    <div class="fbgrp" id="pBtns"></div>
    <div class="fsep"></div>
    <span class="flbl">Company:</span>
    <div class="fbgrp" id="coBtns"></div>
  </div>
  <div class="frow" style="gap:.65rem">
    <span class="flbl">Currency:</span>
    <div class="ms-wrap" id="ccyWrap">
      <div class="ms-box" id="ccyMSBox" onclick="openMS('ccy')">
        <div id="ccyChips" style="display:contents"></div>
        <input class="ms-input" id="ccyInput" placeholder="All Currencies"
               oninput="filterMS('ccy',this.value)"
               onclick="event.stopPropagation();openMS('ccy')">
      </div>
      <div class="ms-dd" id="ccyDD"></div>
    </div>
    <span class="flbl">Account:</span>
    <div class="ms-wrap" id="accWrap">
      <div class="ms-box" id="accMSBox" onclick="openMS('acc')">
        <div id="accChips" style="display:contents"></div>
        <input class="ms-input" id="accInput" placeholder="All Accounts"
               oninput="filterMS('acc',this.value)"
               onclick="event.stopPropagation();openMS('acc')">
      </div>
      <div class="ms-dd" id="accDD"></div>
    </div>
    <div class="fsep"></div>
    <div class="fxrow">
      <span class="fxlbl">📈 FX Sensitivity:</span>
      <div class="fxbtns">
        <button class="fxbtn active" id="fxB0"  onclick="setFX(0)">CURRENT</button>
        <button class="fxbtn"        id="fxBm5" onclick="setFX(-5)">USD −5%</button>
        <button class="fxbtn"        id="fxBp5" onclick="setFX(5)">USD +5%</button>
        <input type="number" step="0.1" placeholder="Custom %" class="fxcustom" id="fxCust"
               onkeydown="if(event.key==='Enter')applyCustomFX()"
               onblur="applyCustomFX()">
      </div>
      <div class="fxcards" id="fxCards">
        <div class="fxcard"><div class="fxcc">–</div><div class="fxcv neu">–</div><div class="fxcl">Exp.</div></div>
        <div class="fxcard"><div class="fxcc">–</div><div class="fxcv neu">–</div><div class="fxcl">Exp.</div></div>
        <div class="fxcard"><div class="fxcc">–</div><div class="fxcv neu">–</div><div class="fxcl">Exp.</div></div>
      </div>
    </div>
  </div>
</div>
<div id="scnBanner"></div>

<div class="main">

<!-- PAGE 1: EXECUTIVE SUMMARY -->
<div class="pc active" id="p1">
  <div class="kgrid">
    <div class="kgrp">
      <div class="kghdr">NET / LONG / SHORT EXPOSURE POSITIONS — USD EQUIVALENT</div>
      <div class="kcards">
        <div class="kcard b"><div class="klbl">Net Exposure (USD)</div><div class="kval" id="kNE">–</div><div class="kchg neu" id="kNEC"><span>–</span></div><div class="kfxd" id="kNEFX"></div></div>
        <div class="kcard r"><div class="klbl">Short Position (USD)</div><div class="kval" id="kSH">–</div><div class="kchg neu" id="kSHC"><span>–</span></div><div class="kfxd" id="kSHFX"></div></div>
        <div class="kcard g"><div class="klbl">Long Position (USD)</div><div class="kval" id="kLO">–</div><div class="kchg neu" id="kLOC"><span>–</span></div><div class="kfxd" id="kLOFX"></div></div>
      </div>
    </div>
    <div class="kgrp">
      <div class="kghdr">NET UNREALIZED G/L — GAIN / LOSS POSITIONS — USD EQUIVALENT</div>
      <div class="kcards">
        <div class="kcard b"><div class="klbl">Net Unrealized G/L</div><div class="kval" id="kNGL">–</div><div class="kchg neu" id="kNGLC"><span>–</span></div><div class="kfxd" id="kNGLFX"></div></div>
        <div class="kcard r"><div class="klbl">Unrealized Loss</div><div class="kval" id="kUL">–</div><div class="kchg neu" id="kULC"><span>–</span></div><div class="kfxd" id="kULFX"></div></div>
        <div class="kcard g"><div class="klbl">Unrealized Gain</div><div class="kval" id="kUG">–</div><div class="kchg neu" id="kUGC"><span>–</span></div><div class="kfxd" id="kUGFX"></div></div>
      </div>
    </div>
  </div>
  <div class="vrow">
    <div class="vbox"><div class="vhdr"><div class="vtitle">🔍 EXPOSURE BREAKDOWN — DRILL-DOWN</div></div><div class="vbody"><div class="dtree" id="drillE"><div class="ldg">Loading…</div></div></div></div>
    <div class="vbox"><div class="vhdr"><div class="vtitle">🔍 UNREALIZED G/L BREAKDOWN — DRILL-DOWN</div></div><div class="vbody"><div class="dtree" id="drillG"><div class="ldg">Loading…</div></div></div></div>
  </div>
  <div class="vrow">
    <div class="vbox"><div class="vhdr"><div class="vtitle">TOTAL EXPOSURE — USD EQUIVALENT BY CURRENCY</div></div><div class="vbody"><div class="cwrap"><canvas id="cExp"></canvas></div></div></div>
    <div class="vbox"><div class="vhdr"><div class="vtitle">TOTAL UNREALIZED G/L BY CURRENCY</div></div><div class="vbody"><div class="cwrap"><canvas id="cGL"></canvas></div></div></div>
  </div>
  <div class="tsbox">
    <div class="tshdr">
      <div class="tshr1">
        <div class="tstitle">📈 EXPOSURE TREND OVER TIME</div>
        <div class="tsctrls">
          <div class="tstgl"><button class="tstbtn" id="expMA" onclick="setTsMode('exp','account')">BY ACCOUNT</button><button class="tstbtn active" id="expMC" onclick="setTsMode('exp','currency')">BY CURRENCY</button></div>
          <div class="ts-msw"><div class="ts-msb" id="expMSB" onclick="openTsDD('exp')"><div id="expChips" style="display:contents"></div><input class="ts-msin" id="expIn" placeholder="Select a currency…" oninput="filterTsDD('exp',this.value)" onclick="event.stopPropagation();openTsDD('exp')"><span class="ts-hint">Multi-select ✚</span></div><div class="ts-dd" id="expDD"></div></div>
        </div>
      </div>
      <div class="tslegrow"><button class="tslegbtn" id="expLP">‹</button><div class="tslegarea"><div class="tsleginner" id="expLeg"></div></div><button class="tslegbtn" id="expLN">›</button></div>
    </div>
    <div class="tsbody"><canvas id="cExpTs"></canvas></div>
  </div>
  <div class="tsbox">
    <div class="tshdr">
      <div class="tshr1">
        <div class="tstitle">📈 UNREALIZED G/L TREND OVER TIME</div>
        <div class="tsctrls">
          <div class="tstgl"><button class="tstbtn" id="glMA" onclick="setTsMode('gl','account')">BY ACCOUNT</button><button class="tstbtn active" id="glMC" onclick="setTsMode('gl','currency')">BY CURRENCY</button></div>
          <div class="ts-msw"><div class="ts-msb" id="glMSB" onclick="openTsDD('gl')"><div id="glChips" style="display:contents"></div><input class="ts-msin" id="glIn" placeholder="Select a currency…" oninput="filterTsDD('gl',this.value)" onclick="event.stopPropagation();openTsDD('gl')"><span class="ts-hint">Multi-select ✚</span></div><div class="ts-dd" id="glDD"></div></div>
        </div>
      </div>
      <div class="tslegrow"><button class="tslegbtn" id="glLP">‹</button><div class="tslegarea"><div class="tsleginner" id="glLeg"></div></div><button class="tslegbtn" id="glLN">›</button></div>
    </div>
    <div class="tsbody"><canvas id="cGLTs"></canvas></div>
  </div>
  <div class="ibox"><div class="ittl">Executive Insights</div><div class="igrid" id="insGrid"><div class="ldg">Generating…</div></div></div>
</div>

<!-- PAGE 2: ENTITY LEVEL EXPOSURE DETAILS -->
<div class="pc" id="p2">
  <div class="p2top">
    <div class="p2pies">
      <div class="piebox"><div class="piehdr">YPOPL — TOTAL EXPOSURE LCY</div><div class="piebody"><canvas id="pieYPOPL"></canvas></div></div>
      <div class="piebox"><div class="piehdr">YPNL BV — TOTAL EXPOSURE LCY</div><div class="piebody"><canvas id="pieYPNLBV"></canvas></div></div>
      <div class="piebox"><div class="piehdr">YPRODAS — TOTAL EXPOSURE LCY</div><div class="piebody"><canvas id="pieYPAS"></canvas></div></div>
    </div>
    <div class="p2trends">
      <div class="trendbox"><div class="trendhdr">📊 YPOPL — UNREALIZED G/L TREND</div><div class="trendbody"><canvas id="trendYPOPL"></canvas></div></div>
      <div class="trendbox"><div class="trendhdr">📊 YPNL BV — UNREALIZED G/L TREND</div><div class="trendbody"><canvas id="trendYPNLBV"></canvas></div></div>
      <div class="trendbox"><div class="trendhdr">📊 YPRODAS — UNREALIZED G/L TREND</div><div class="trendbody"><canvas id="trendYPAS"></canvas></div></div>
    </div>
  </div>
  <div class="p2pvts">
    <div class="pvtbox"><div class="pvthdr"><div class="pvtttl">TOTAL EXPOSURE LCY — PIVOT</div><div class="pvtsels" id="pvtSel1"></div></div><div class="pvtbody" id="pvtBdy1"><div class="ldg">Loading…</div></div></div>
    <div class="pvtbox"><div class="pvthdr"><div class="pvtttl">NET CURRENCY GAIN / LOSS — PIVOT</div><div class="pvtsels" id="pvtSel2"></div></div><div class="pvtbody" id="pvtBdy2"><div class="ldg">Loading…</div></div></div>
  </div>
</div>

<!-- PAGE 3: VARIANCE ACCOUNT MOVEMENT -->
<div class="pc" id="p3">
  <div class="p3top">
    <div class="p3chart">
      <div class="p3chdr">
        <div class="p3cttl">END BALANCE TREND ANALYSIS</div>
        <div class="p3csub">The chart displays ending balances for Balance Sheet accounts and period movements for P&amp;L accounts</div>
        <div class="p3legrow"><button class="tslegbtn" id="p3L1P">‹</button><div class="p3legarea"><div class="p3leginner" id="p3Leg1"></div></div><button class="tslegbtn" id="p3L1N">›</button></div>
      </div>
      <div class="p3cbody"><canvas id="p3C1"></canvas></div>
    </div>
    <div class="p3chart">
      <div class="p3chdr">
        <div class="p3cttl">PERIOD MOVEMENT (GL) STEP ANALYSIS</div>
        <div class="p3csub">The chart displays ending balances for Balance Sheet accounts and period movements for P&amp;L accounts</div>
        <div class="p3legrow"><button class="tslegbtn" id="p3L2P">‹</button><div class="p3legarea"><div class="p3leginner" id="p3Leg2"></div></div><button class="tslegbtn" id="p3L2N">›</button></div>
      </div>
      <div class="p3cbody"><canvas id="p3C2"></canvas></div>
    </div>
  </div>
  <div class="p3tbl-wrap">
    <div class="p3tbl-hdr">
      <div class="p3tbl-ttl">VARIANCE ACCOUNT MOVEMENT DETAIL — <span id="p3TblFilter">ALL CURRENCIES</span></div>
      <button class="p3tbl-clr" id="p3TblClr" onclick="clearP3Filter()">✕ Clear Filter</button>
    </div>
    <div class="p3tbl-body" id="p3TblBody"><div class="ldg">Loading…</div></div>
  </div>
</div>

</div><!-- /main -->

<footer>
  <div class="ftl">
    BALANCE SHEET EXPOSURE &amp; VARIANCE ANALYSIS REPORT<br>
    <span style="font-size:.82rem;font-weight:900;color:#ffffff;letter-spacing:.06em;text-transform:uppercase">UYGAR TALU</span>
  </div>
  <div class="ftm">A comprehensive, automated reporting framework providing C-level executives with real-time FX exposure, unrealized G/L scenarios, and account-level variance analysis across entities — enabling strategic treasury decisions with minimal human intervention.</div>
  <div class="ftr" id="ftDate">Loading…</div>
</footer>

<script>
Chart.register(ChartDataLabels);

// ── Data payload (injected by Python generator) ──────────────────────────
const PL   = DATA_PAYLOAD;
const BSR  = PL.bs;
const PERIODS = PL.periods;
const LATEST  = PL.latest;
const VAR     = PL.var;
const VMONS   = PL.var_months;

// ── Currency-to-colour map ───────────────────────────────────────────────
const CCYCLR = {
  SGD:'#F5C400',GBP:'#40C4FF',MYR:'#00E676',AED:'#CE93D8',EUR:'#1DE9B6',
  AUD:'#FF9100',BRL:'#FF4081',CNY:'#FF6E40',GHS:'#64FFDA',IDR:'#EA80FC',
  INR:'#82B1FF',NOK:'#7C4DFF',THB:'#18FFFF',VND:'#69F0AE',ZAR:'#FFD740',
  AOA:'#B0BEC5',SEK:'#FF80AB',CHF:'#80DEEA',DKK:'#FFCC80',HKD:'#A5D6A7',
  JPY:'#EF9A9A',NAD:'#CE93D8',TWD:'#80CBC4',
};
/** Fallback palette for currencies not in CCYCLR. */
const PAL = ['#F5C400','#00E676','#40C4FF','#FF1744','#CE93D8','#1DE9B6','#FF9100',
             '#FF4081','#64FFDA','#EA80FC','#82B1FF','#7C4DFF','#18FFFF','#69F0AE',
             '#FFD740','#B0BEC5','#FF80AB','#80DEEA','#FFCC80','#A5D6A7'];

// ── Application state ────────────────────────────────────────────────────
let S = {
  periods:   [LATEST],   // active period filter; empty array means "all"
  companies: ['YPOPL'],  // active company filter; empty array means "all"
  ccys:      [],         // active currency filter
  accs:      [],         // active account filter
  fxShift:   0,          // FX sensitivity shift in percent (+ = USD stronger)
};

let tsS = {
  exp: { mode:'currency', sel:[] },   // timeseries state for Exposure chart
  gl:  { mode:'currency', sel:[] },   // timeseries state for GL chart
};

let pvtEnt    = { 1:'YPOPL', 2:'YPOPL' };  // selected entity per pivot table
let p3ActiveCcy = null;                      // clicked currency on page-3 charts
let charts    = {};                          // Chart.js instance registry
const drillLvl = {};                         // drill-down level per tree widget

// ── Number formatters ────────────────────────────────────────────────────
const NF = new Intl.NumberFormat('en-US');

/** Format a value as compact currency string (e.g. $12.34M, $456.78K). */
const fmtK = v => {
  if(v==null||isNaN(v)) return '–';
  const a=Math.abs(v), s=v<0?'-':'';
  if(a>=1e9) return s+'$'+(a/1e9).toFixed(2)+'B';
  if(a>=1e6) return s+'$'+(a/1e6).toFixed(2)+'M';
  if(a>=1e3) return s+'$'+(a/1e3).toFixed(2)+'K';
  return s+'$'+a.toFixed(2);
};

/** Format a value as full comma-separated integer USD string (e.g. $1,234,567). */
const fmtFull = v => {
  if(v==null||isNaN(v)) return '–';
  const a=Math.abs(v), s=v<0?'-$':'$';
  return s + NF.format(Math.round(a));
};

/** Format a value as short string for axis labels and bar data labels. */
const fmtS = v => {
  if(v==null||isNaN(v)) return null;
  const a=Math.abs(v), s=v<0?'-':'';
  if(a>=1e9) return s+(a/1e9).toFixed(2)+'B';
  if(a>=1e6) return s+(a/1e6).toFixed(2)+'M';
  if(a>=1e3) return s+(a/1e3).toFixed(1)+'K';
  if(a<10)   return null;
  return s+Math.round(a);
};

/** Return '+' for positive values, '' for zero/negative. */
const sg = v => v>0?'+':'';

/** Convert a period label like '2025 Dec' to a month key like '2025-12'. */
const pctToM = p => {
  const [yr,mo]=p.split(' ');
  const mm={Jan:'01',Feb:'02',Mar:'03',Apr:'04',May:'05',Jun:'06',Jul:'07',Aug:'08',Sep:'09',Oct:'10',Nov:'11',Dec:'12'};
  return `${yr}-${mm[mo]||'01'}`;
};

/** Convert a month key like '2025-12' to a display label like 'Dec 2025'. */
const mToLbl = m => {
  const [yr,mo]=m.split('-');
  const mn=['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${mn[parseInt(mo)]} ${yr}`;
};

// ── FX adjustment helpers ────────────────────────────────────────────────
/**
 * Return FX-adjusted USD amount for a record.
 * fxShift > 0  → USD strengthened  → non-USD worth less in USD.
 * fxShift < 0  → USD weakened       → non-USD worth more in USD.
 */
const adjU = r => r.usd_amt * (1 - S.fxShift/100);

/** Return FX-adjusted net G/L for a record (same scaling as adjU). */
const adjG = r => r.net_gl  * (1 - S.fxShift/100);

// ── Filter predicates ────────────────────────────────────────────────────
const mP = r => S.periods.length===0   || S.periods.includes(r.period);
const mC = r => S.companies.length===0 || S.companies.includes(r.entity);
const mY = r => S.ccys.length===0      || S.ccys.includes(r.ccy);
const mA = r => S.accs.length===0      || S.accs.includes(r.account);

/** Apply all active filters to the BS exposure record set. */
const filt  = (recs=BSR) => recs.filter(r=>mP(r)&&mC(r)&&mY(r)&&mA(r));

/** Return the prior-period BS exposure records (one period back). */
const filtPrev = () => {
  if(S.periods.length!==1) return [];
  const idx=PERIODS.indexOf(S.periods[0]); if(idx<1) return [];
  const pp=PERIODS[idx-1];
  return BSR.filter(r=>r.period===pp&&mC(r)&&mY(r)&&mA(r));
};

/** Apply company/currency/account filters to the variance record set. */
const filtV = (vrecs=VAR) => vrecs.filter(r=>{
  const okP = S.periods.length===0 || S.periods.map(pctToM).includes(r.month);
  const okC = S.companies.length===0 || S.companies.includes(r.company);
  const okY = S.ccys.length===0 || S.ccys.includes(r.ccy);
  const okA = S.accs.length===0 || S.accs.includes(r.account);
  return okP&&okC&&okY&&okA;
});

// ── Initialisation ───────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  const now   = new Date();
  const DAYS  = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const MNTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  document.getElementById('hdrDate').textContent =
    `${String(now.getDate()).padStart(2,'0')} ${MNTHS[now.getMonth()]} ${now.getFullYear()}`;
  document.getElementById('hdrDay').textContent = DAYS[now.getDay()];
  document.getElementById('hdrDR').textContent  = PL.date_range;
  document.getElementById('ftDate').textContent = 'DATA RANGE: ' + PL.date_range;

  // Session timer (updates every second)
  const t0 = Date.now();
  setInterval(() => {
    const s = Math.floor((Date.now()-t0)/1000);
    document.getElementById('timerV').textContent =
      `${String(Math.floor(s/3600)).padStart(2,'0')}:${String(Math.floor((s%3600)/60)).padStart(2,'0')}:${String(s%60).padStart(2,'0')}`;
  }, 1000);

  buildFilters();
  setupClose();
  applyAll();
  initP2();
  initP3LegBtns();
});

// ── Filter bar construction ──────────────────────────────────────────────

/** Build all filter buttons and dropdowns in the sticky filter bar. */
function buildFilters() {
  // Period toggle buttons (ALL + one per period)
  const pb = document.getElementById('pBtns');
  ['ALL', ...PERIODS].forEach(p => {
    const b = document.createElement('button');
    b.className = 'fbtn' + (p===LATEST?' active':'');
    b.textContent = p; b.dataset.v = p;
    b.onclick = () => togglePeriod(p, b);
    pb.appendChild(b);
  });

  // Company toggle buttons
  const cb = document.getElementById('coBtns');
  [['ALL','ALL'],['YPOPL','YPOPL'],['YPNLBV','YPNL BV'],['YPAS','YPRODAS']].forEach(([k,l]) => {
    const b = document.createElement('button');
    b.className = 'fbtn' + (k==='YPOPL'?' active':''); b.dataset.key=k; b.textContent=l;
    b.onclick = () => toggleComp(k, b); cb.appendChild(b);
  });

  // Multi-select options for currency and account
  buildMSOpts('ccy', PL.bs_ccys);
  buildMSOpts('acc', PL.bs_accounts);
  renderMSChips('ccy');
  renderMSChips('acc');

  // Legend scroll buttons for time series charts
  ['exp','gl'].forEach(k => {
    document.getElementById(k+'LP').onclick = () => document.getElementById(k+'Leg').scrollLeft -= 180;
    document.getElementById(k+'LN').onclick = () => document.getElementById(k+'Leg').scrollLeft += 180;
  });
}

/** Toggle a single period in/out of the active period set. */
function togglePeriod(p, btn) {
  if(p==='ALL') {
    S.periods = [];
    document.querySelectorAll('#pBtns .fbtn').forEach(b => b.classList.toggle('active', b.dataset.v==='ALL'));
  } else {
    S.periods = S.periods.filter(x => x!=='ALL');
    const idx = S.periods.indexOf(p);
    if(idx>=0) { if(S.periods.length>1){S.periods.splice(idx,1); btn.classList.remove('active');} }
    else        { S.periods.push(p); btn.classList.add('active'); }
    document.querySelectorAll('#pBtns .fbtn').forEach(b => { if(b.dataset.v==='ALL') b.classList.remove('active'); });
    if(S.periods.length===0) {
      S.periods = [LATEST];
      document.querySelectorAll('#pBtns .fbtn').forEach(b => b.classList.toggle('active', b.dataset.v===LATEST));
    }
  }
  applyAll();
}

/** Toggle a single company in/out of the active company set. */
function toggleComp(key, btn) {
  if(key==='ALL') {
    S.companies = [];
    document.querySelectorAll('#coBtns .fbtn').forEach(b => b.classList.toggle('active', b.dataset.key==='ALL'));
  } else {
    S.companies = S.companies.filter(c => c!=='ALL');
    const idx = S.companies.indexOf(key);
    if(idx>=0) { if(S.companies.length>1) S.companies.splice(idx,1); }
    else        { S.companies.push(key); }
    if(S.companies.length===0) S.companies = ['YPOPL'];
    document.querySelectorAll('#coBtns .fbtn').forEach(b => {
      if(b.dataset.key==='ALL') b.classList.remove('active');
      else b.classList.toggle('active', S.companies.includes(b.dataset.key));
    });
  }
  applyAll();
}

// ── Multi-select chip widget (currency & account) ────────────────────────

const MSState = {
  ccy: { sel:[], opts: PL.bs_ccys },
  acc: { sel:[], opts: PL.bs_accounts },
};

function buildMSOpts(key, opts) { MSState[key].opts = opts; }

/** Open the dropdown for a multi-select widget. */
function openMS(key) {
  document.getElementById(key+'DD').classList.add('open');
  rebuildMSOpts(key, '');
}

/** Filter the dropdown options by a search query. */
function filterMS(key, q) {
  rebuildMSOpts(key, q);
  document.getElementById(key+'DD').classList.add('open');
}

/** Rebuild the dropdown option list, respecting the current search query. */
function rebuildMSOpts(key, q) {
  const sel = MSState[key].sel, opts = MSState[key].opts;
  const dd  = document.getElementById(key+'DD'); dd.innerHTML = '';
  const ao  = document.createElement('div'); ao.className = 'ms-opt all-opt';
  ao.textContent = 'ALL (Clear Selection)';
  ao.onclick = e => { e.stopPropagation(); MSState[key].sel=[]; renderMSChips(key); applyAll(); document.getElementById(key+'DD').classList.remove('open'); };
  dd.appendChild(ao);
  (q ? opts.filter(o => o.toLowerCase().includes(q.toLowerCase())) : opts).forEach(o => {
    const d = document.createElement('div'); d.className = 'ms-opt'+(sel.includes(o)?' sel':'');
    d.textContent = o;
    if(!sel.includes(o)) d.onclick = e => { e.stopPropagation(); addMS(key, o); };
    dd.appendChild(d);
  });
}

/** Add a value to the multi-select selection. */
function addMS(key, val) {
  if(!MSState[key].sel.includes(val)) { MSState[key].sel.push(val); renderMSChips(key); applyAll(); }
  document.getElementById(key+'Input').value = '';
  rebuildMSOpts(key, '');
}

/** Remove a value from the multi-select selection. */
function removeMS(key, val) {
  MSState[key].sel = MSState[key].sel.filter(v => v!==val);
  renderMSChips(key); applyAll(); rebuildMSOpts(key, '');
}

/** Re-render the chip badges inside the multi-select box. */
function renderMSChips(key) {
  const box = document.getElementById(key+'Chips'); box.innerHTML = '';
  const sel = MSState[key].sel;
  if(sel.length===0) {
    const c = document.createElement('span'); c.className='ms-chip all'; c.textContent='ALL'; box.appendChild(c);
  } else {
    sel.forEach(v => {
      const c = document.createElement('span'); c.className='ms-chip';
      c.innerHTML = `${v}<span class="ms-chip-x" onclick="removeMS('${key}','${v.replace(/'/g,"\\'")}')">×</span>`;
      box.appendChild(c);
    });
  }
  if(key==='ccy') S.ccys = MSState.ccy.sel;
  if(key==='acc') S.accs = MSState.acc.sel;
}

// ── FX Sensitivity ────────────────────────────────────────────────────────

/** Set a preset FX sensitivity scenario (0 / -5 / +5). */
function setFX(p) {
  S.fxShift = p;
  document.getElementById('fxCust').value = '';
  ['fxB0','fxBm5','fxBp5'].forEach(id => document.getElementById(id).classList.remove('active'));
  if(p===0)  document.getElementById('fxB0').classList.add('active');
  else if(p===-5) document.getElementById('fxBm5').classList.add('active');
  else if(p===5)  document.getElementById('fxBp5').classList.add('active');
  applyAll();
}

/** Read the custom % input and apply it as the FX shift. */
function applyCustomFX() {
  const v = parseFloat(document.getElementById('fxCust').value);
  if(isNaN(v)) return;
  S.fxShift = v;
  ['fxB0','fxBm5','fxBp5'].forEach(id => document.getElementById(id).classList.remove('active'));
  applyAll();
}

/**
 * Update the scenario banner and top-3 FX impact cards.
 * Called by applyAll whenever the state changes.
 */
function updateFXUI() {
  const sh = S.fxShift;
  const bn = document.getElementById('scnBanner');
  bn.style.display = sh===0 ? 'none' : 'block';
  if(sh!==0) bn.textContent =
    `⚡ FX SCENARIO: USD ${sh>0?'STRENGTHENED':'WEAKENED'} ${Math.abs(sh)}% vs ALL NON-USD CURRENCIES — ALL USD VALUES ADJUSTED`;

  // Top-3 currency exposure cards
  const recs  = filt();
  const byCcy = {};
  recs.forEach(r => {
    if(!byCcy[r.ccy]) byCcy[r.ccy] = {ccy:r.ccy, adj:0, base:0};
    byCcy[r.ccy].adj  += adjU(r);
    byCcy[r.ccy].base += r.usd_amt;
  });
  const top3  = Object.values(byCcy).sort((a,b)=>Math.abs(b.adj)-Math.abs(a.adj)).slice(0,3);
  const cards = document.querySelectorAll('#fxCards .fxcard');
  for(let i=0; i<3; i++) {
    const d = top3[i];
    if(!d) { cards[i].querySelector('.fxcc').textContent='–'; cards[i].querySelector('.fxcv').textContent='–'; cards[i].querySelector('.fxcv').className='fxcv neu'; continue; }
    cards[i].querySelector('.fxcc').textContent = d.ccy;
    const ve = cards[i].querySelector('.fxcv');
    ve.className = 'fxcv '+(d.adj>=0?'pos':'neg');
    ve.textContent = fmtK(d.adj);
    const imp = d.adj - d.base;
    cards[i].querySelector('.fxcl').textContent = sh===0 ? 'USD Exp.' : `Δ: ${sg(imp)}${fmtK(imp)}`;
  }
}

// ── Master render trigger ────────────────────────────────────────────────

/** Re-render every visual on the active page when state changes. */
function applyAll() {
  updateFXUI();
  renderKPIs();
  renderDrills();
  renderBars();
  renderTimeseries();
  renderInsights();
  if(document.getElementById('p2').classList.contains('active')) renderP2();
  if(document.getElementById('p3').classList.contains('active')) renderP3();
}

// ── KPI Cards ─────────────────────────────────────────────────────────────

/**
 * Render all six KPI cards.
 * Shows:
 *  - Current value (FX-scenario adjusted)
 *  - vs Prior period change
 *  - FX Δ badge (shown only when a non-zero FX scenario is active)
 */
function renderKPIs() {
  const recs = filt(), prev = filtPrev();

  // Baseline values at fxShift=0 (needed for FX delta comparison)
  let bNet=0,bSh=0,bLo=0,bNgl=0,bUg=0,bUl=0;
  recs.forEach(r => {
    const u=r.usd_amt; bNet+=u;
    if(u<0) bSh+=u; else bLo+=u;
    bNgl+=r.net_gl; bUg+=r.u_gain; bUl+=r.u_loss;
  });

  // Scenario-adjusted values
  let net=0,sh=0,lo=0,ngl=0,ug=0,ul=0;
  recs.forEach(r => {
    const u=adjU(r); net+=u;
    if(u<0) sh+=u; else lo+=u;
    const g=adjG(r); ngl+=g;
    ug+=r.u_gain*(1-S.fxShift/100);
    ul+=r.u_loss*(1-S.fxShift/100);
  });

  // Prior-period baseline (for vs-prior change)
  let pnet=0,psh=0,plo=0,pngl=0,pug=0,pul=0;
  prev.forEach(r => {
    const u=r.usd_amt; pnet+=u;
    if(u<0) psh+=u; else plo+=u;
    pngl+=r.net_gl; pug+=r.u_gain; pul+=r.u_loss;
  });

  // Helper: update a single KPI value + prior-period change label
  const sk = (vid, cid, v, pv) => {
    const e = document.getElementById(vid); if(e) e.textContent = fmtK(v);
    const ec = document.getElementById(cid);
    if(ec && prev.length) { const d=v-pv; ec.className='kchg '+(d>=0?'pos':'neg'); ec.innerHTML=`<span>${sg(d)}${fmtK(Math.abs(d))}</span>`; }
  };
  sk('kNE','kNEC',net,pnet); sk('kSH','kSHC',sh,psh); sk('kLO','kLOC',lo,plo);
  sk('kNGL','kNGLC',ngl,pngl); sk('kUG','kUGC',ug,pug);
  const ule = document.getElementById('kUL'); if(ule) ule.textContent = '-'+fmtK(ul);
  const ulc = document.getElementById('kULC');
  if(ulc && prev.length) { const d=ul-pul; ulc.className='kchg '+(d<=0?'pos':'neg'); ulc.innerHTML=`<span>${sg(-d)}${fmtK(Math.abs(d))}</span>`; }

  // FX delta badges — only visible when a scenario is active
  const fxPairs = [
    ['kNEFX',  net  - bNet,  true ],   // true = higher is better
    ['kSHFX',  sh   - bSh,   true ],
    ['kLOFX',  lo   - bLo,   true ],
    ['kNGLFX', ngl  - bNgl,  true ],
    ['kULFX',  -(ul - bUl),  false],   // loss: invert for intuitive colour
    ['kUGFX',  ug   - bUg,   true ],
  ];
  const fxActive = S.fxShift !== 0;
  fxPairs.forEach(([id, delta, higherIsBetter]) => {
    const el = document.getElementById(id); if(!el) return;
    if(!fxActive) { el.classList.remove('show','pos','neg'); el.textContent=''; return; }
    el.classList.add('show');
    el.classList.toggle('pos', higherIsBetter ? delta>=0 : delta<=0);
    el.classList.toggle('neg', higherIsBetter ? delta<0  : delta>0);
    el.textContent = `${sg(delta)}${fmtK(Math.abs(delta))}`;
  });
}

// ── Drill-down trees ──────────────────────────────────────────────────────

/** Render both interactive drill-down trees (Exposure and G/L). */
function renderDrills() {
  buildDrill('drillE', filt(), r=>adjU(r), 'Exposure (USD)');
  buildDrill('drillG', filt(), r=>adjG(r), 'Net G/L');
}

/**
 * Build or update a drill-down tree widget.
 *
 * Level 0 (root): shows totals per currency.
 * Level 1 (currency): shows breakdown by account for that currency.
 */
function buildDrill(id, recs, vFn, lbl) {
  if(!drillLvl[id]) drillLvl[id] = {lvl:'root', key:null};
  const lvl = drillLvl[id];
  const bc = {};
  recs.forEach(r => {
    if(!bc[r.ccy]) bc[r.ccy] = {ccy:r.ccy, val:0, accs:{}};
    bc[r.ccy].val += vFn(r);
    if(!bc[r.ccy].accs[r.account]) bc[r.ccy].accs[r.account] = 0;
    bc[r.ccy].accs[r.account] += vFn(r);
  });
  const srt = Object.values(bc).sort((a,b)=>Math.abs(b.val)-Math.abs(a.val));
  const tot = srt.reduce((s,c)=>s+c.val, 0);
  const el  = document.getElementById(id); el.innerHTML = '';

  // Root total row (click to reset)
  const root = document.createElement('div'); root.className = 'droot';
  root.innerHTML = `<span class="drootl">Total ${lbl}</span><span class="drootv">${sg(tot)}${fmtFull(tot)}</span>`;
  root.onclick = () => { drillLvl[id]={lvl:'root',key:null}; buildDrill(id,recs,vFn,lbl); };
  el.appendChild(root);

  if(lvl.lvl==='root') {
    srt.forEach(d => {
      const item = document.createElement('div'); item.className = 'ditem';
      const c = d.val>=0 ? 'var(--green)' : 'var(--red)';
      item.innerHTML = `<span class="dccy">${d.ccy}</span><span class="dacc"></span><span class="dval" style="color:${c}">${sg(d.val)}${fmtFull(Math.abs(d.val))}</span><span class="dchev">›</span>`;
      item.onclick = () => { drillLvl[id]={lvl:'ccy',key:d.ccy}; buildDrill(id,recs,vFn,lbl); };
      el.appendChild(item);
    });
  } else {
    const bk = document.createElement('div'); bk.className = 'dback';
    bk.innerHTML = '<span style="color:var(--gold)">‹</span><span style="font-size:.66rem;font-weight:700;color:var(--t2);text-transform:uppercase;margin-left:.38rem">Back to currencies</span>';
    bk.onclick = () => { drillLvl[id]={lvl:'root',key:null}; buildDrill(id,recs,vFn,lbl); };
    el.appendChild(bk);
    const d = bc[lvl.key]; if(!d) return;
    Object.entries(d.accs).sort((a,b)=>Math.abs(b[1])-Math.abs(a[1])).forEach(([acc,val]) => {
      const it = document.createElement('div'); it.className = 'daitem';
      it.innerHTML = `<span style="font-size:.65rem;font-weight:700;color:var(--t2);flex:1;line-height:1.3">${acc}</span><span style="font-size:.9rem;font-weight:800;font-family:'DM Mono',monospace;color:${val>=0?'var(--green)':'var(--red)'}">${sg(val)}${fmtFull(Math.abs(val))}</span>`;
      el.appendChild(it);
    });
  }
}

// ── Horizontal bar charts ─────────────────────────────────────────────────

/**
 * Render the two horizontal bar charts (Exposure USD and Unrealized G/L),
 * both FX-scenario adjusted.
 */
function renderBars() {
  const recs = filt(); const bc = {};
  recs.forEach(r => { if(!bc[r.ccy]) bc[r.ccy]={e:0,g:0}; bc[r.ccy].e+=adjU(r); bc[r.ccy].g+=adjG(r); });
  const srt  = Object.entries(bc).sort((a,b)=>Math.abs(b[1].e)-Math.abs(a[1].e));
  const lbl  = srt.map(e=>e[0]), ev = srt.map(e=>e[1].e), gv = srt.map(e=>e[1].g);
  const bge  = lbl.map(l=>CCYCLR[l]||'#F5C400');
  const bgg  = gv.map(v=>v>=0?'rgba(0,230,118,.8)':'rgba(255,23,68,.8)');
  const h    = Math.max(200, lbl.length*28+40);
  ['cExp','cGL'].forEach(id => { const cv=document.getElementById(id); if(cv) cv.parentElement.style.minHeight=h+'px'; });
  const barOpt = (data, bg) => ({
    type:'bar', data:{labels:lbl, datasets:[{data, backgroundColor:bg, borderRadius:4, borderSkipped:false}]},
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:c=>`${c.label}: ${sg(c.raw)}${fmtFull(Math.abs(c.raw))}`}},
        datalabels:{display:true, color:'#fff', anchor:'end', align:'end',
                    font:{family:"'DM Mono',monospace",weight:'bold',size:10},
                    clamp:true, clip:true, formatter:v=>Math.abs(v)<500?null:fmtS(v)}
      },
      scales:{
        x:{ticks:{color:'#ffffff',font:{size:10},callback:v=>fmtS(v)||'0'},grid:{color:'rgba(255,255,255,.07)'}},
        y:{ticks:{color:'#ffffff',font:{size:10,weight:'700'}},grid:{color:'rgba(255,255,255,.05)'}}
      }
    }
  });
  destroyC('cExp'); charts['cExp'] = new Chart(document.getElementById('cExp'), barOpt(ev,bge));
  destroyC('cGL');  charts['cGL']  = new Chart(document.getElementById('cGL'),  barOpt(gv,bgg));
}

// ── Time-series charts ────────────────────────────────────────────────────

/** Toggle between BY ACCOUNT and BY CURRENCY mode for a time-series chart. */
function setTsMode(chart, mode) {
  tsS[chart].mode=mode; tsS[chart].sel=[];
  document.getElementById(chart+'MA').classList.toggle('active',mode==='account');
  document.getElementById(chart+'MC').classList.toggle('active',mode==='currency');
  document.getElementById(chart+'In').placeholder = mode==='account'?'Select an account…':'Select a currency…';
  rebuildTsDD(chart); renderTsChart(chart);
}

function openTsDD(ch)          { document.getElementById(ch+'DD').classList.add('open'); rebuildTsDD(ch,''); }
function filterTsDD(ch,q)      { rebuildTsDD(ch,q); document.getElementById(ch+'DD').classList.add('open'); }

/** Rebuild time-series dropdown option list. */
function rebuildTsDD(ch, q='') {
  const mode=tsS[ch].mode, sel=tsS[ch].sel;
  const recs = BSR.filter(r=>S.companies.length===0||S.companies.includes(r.entity));
  const opts = [...new Set(recs.map(r=>mode==='account'?r.account:r.ccy))].sort();
  const dd   = document.getElementById(ch+'DD'); dd.innerHTML='';
  const ao   = document.createElement('div'); ao.className='ts-ddopt all'; ao.textContent='ALL (Show Everything)';
  ao.onclick = e => { e.stopPropagation(); tsS[ch].sel=[]; renderTsChips(ch); renderTsChart(ch); rebuildTsDD(ch,''); };
  dd.appendChild(ao);
  (q?opts.filter(o=>o.toLowerCase().includes(q.toLowerCase())):opts).forEach(o => {
    const d=document.createElement('div'); d.className='ts-ddopt'+(sel.includes(o)?' sel':''); d.textContent=o;
    if(!sel.includes(o)) d.onclick=e=>{e.stopPropagation();addTsSel(ch,o);};
    dd.appendChild(d);
  });
}

function addTsSel(ch, val) {
  if(!tsS[ch].sel.includes(val)){tsS[ch].sel.push(val);renderTsChips(ch);renderTsChart(ch);}
  document.getElementById(ch+'In').value=''; rebuildTsDD(ch,'');
}

function removeTsSel(ch, val) {
  tsS[ch].sel=tsS[ch].sel.filter(v=>v!==val);
  renderTsChips(ch); renderTsChart(ch); rebuildTsDD(ch,'');
}

/** Re-render the chip badges in the time-series multi-select input. */
function renderTsChips(ch) {
  const box=document.getElementById(ch+'Chips'); box.innerHTML='';
  const sel=tsS[ch].sel;
  if(sel.length===0){const c=document.createElement('span');c.className='ts-chip all';c.textContent='ALL';box.appendChild(c);}
  else sel.forEach(v=>{const c=document.createElement('span');c.className='ts-chip';c.innerHTML=`${v}<span class="ts-chip-x" onclick="removeTsSel('${ch}','${v.replace(/'/g,"\\'")}')">×</span>`;box.appendChild(c);});
}

/** Render all time-series charts. */
function renderTimeseries(){renderTsChart('exp');renderTsChart('gl');}

/**
 * Render a single time-series (line) chart.
 * Shows one dataset per selected item (currency or account).
 * Empty selection means "show all".
 */
function renderTsChart(ch) {
  const mode=tsS[ch].mode, sel=tsS[ch].sel, isE=(ch==='exp');
  const recs=BSR.filter(r=>(S.companies.length===0||S.companies.includes(r.entity))&&(S.ccys.length===0||S.ccys.includes(r.ccy))&&(S.accs.length===0||S.accs.includes(r.account)));
  const allIt=[...new Set(recs.map(r=>mode==='account'?r.account:r.ccy))].sort();
  const items=sel.length>0?sel:allIt;
  const agg={}; PERIODS.forEach(p=>{agg[p]={};items.forEach(i=>agg[p][i]=0);});
  recs.forEach(r=>{const k=mode==='account'?r.account:r.ccy;if(items.includes(k)&&agg[r.period])agg[r.period][k]+=(isE?r.usd_amt:r.net_gl);});
  const datasets=items.map((it,idx)=>{
    const clr=mode==='currency'?(CCYCLR[it]||PAL[idx%PAL.length]):PAL[idx%PAL.length];
    return{label:it,data:PERIODS.map(p=>agg[p]?agg[p][it]||0:0),borderColor:clr,backgroundColor:clr+'20',borderWidth:2,pointRadius:2.5,pointHoverRadius:5,tension:.35,fill:false};
  });
  const cid='c'+(ch==='exp'?'Exp':'GL')+'Ts';
  destroyC(cid);
  charts[cid]=new Chart(document.getElementById(cid),{
    type:'line', data:{labels:PERIODS, datasets},
    options:{
      responsive:true,maintainAspectRatio:false,
      interaction:{mode:'nearest',intersect:true},
      plugins:{
        legend:{display:false},
        tooltip:{mode:'nearest',intersect:true,callbacks:{label:c=>`${c.dataset.label}: ${sg(c.raw)}${fmtFull(Math.abs(c.raw))}`}},
        datalabels:{display:false}
      },
      scales:{
        x:{ticks:{color:'#fff',font:{size:9},maxRotation:45},grid:{color:'rgba(255,255,255,.06)'}},
        y:{ticks:{color:'#fff',font:{size:9},callback:v=>fmtS(v)||'0'},grid:{color:'rgba(255,255,255,.06)'}}
      }
    }
  });
  const leg=document.getElementById(ch+'Leg'); leg.innerHTML='';
  datasets.forEach(ds=>{const s=document.createElement('span');s.className='tslegitem';s.innerHTML=`<span class="tslegdot" style="background:${ds.borderColor}"></span>${ds.label}`;leg.appendChild(s);});
  renderTsChips(ch);
}

function setupClose() {
  document.addEventListener('click', e => {
    ['exp','gl'].forEach(ch=>{const b=document.getElementById(ch+'MSB');if(b&&!b.contains(e.target))document.getElementById(ch+'DD').classList.remove('open');});
    ['ccy','acc'].forEach(k=>{const b=document.getElementById(k+'MSBox');if(b&&!b.contains(e.target))document.getElementById(k+'DD').classList.remove('open');});
  });
}

// ── Insights ──────────────────────────────────────────────────────────────

/** Render the executive insights summary box. */
function renderInsights() {
  const recs=filt(); const bc={};
  recs.forEach(r=>{if(!bc[r.ccy])bc[r.ccy]={e:0,g:0};bc[r.ccy].e+=adjU(r);bc[r.ccy].g+=adjG(r);});
  const srt=Object.entries(bc).sort((a,b)=>Math.abs(b[1].e)-Math.abs(a[1].e));
  const tot=recs.reduce((s,r)=>s+adjU(r),0), ngl=recs.reduce((s,r)=>s+adjG(r),0);
  const top3=srt.slice(0,3).map(e=>e[0]).join(', '), topC=srt[0];
  const mg=Object.entries(bc).sort((a,b)=>b[1].g-a[1].g)[0];
  const ml=Object.entries(bc).sort((a,b)=>a[1].g-b[1].g)[0];
  const ins=[
    {l:`Total Exposure — ${S.periods.length===1?S.periods[0]:'Multi-Period'}`,v:`Net USD-equivalent exposure is <span class="hi">${sg(tot)}${fmtFull(Math.abs(tot))}</span> across all CONSIDERED positions.`},
    {l:'Dominant Currency',v:`<span class="hi">${topC?topC[0]:'–'}</span> largest at <span class="hi">${topC?fmtFull(Math.abs(topC[1].e)):'–'}</span>. Top-3: <span class="hi">${top3}</span>.`},
    {l:'Unrealized G/L',v:`Net G/L: <span class="hi" style="color:${ngl>=0?'var(--green)':'var(--red)'}">${sg(ngl)}${fmtFull(Math.abs(ngl))}</span>. Highest gain: <span class="hi">${mg?mg[0]:'–'}</span> | Highest loss: <span class="hi">${ml?ml[0]:'–'}</span>.`},
    {l:'FX Scenario',v:S.fxShift===0?`No scenario active.`:`At <span class="hi">USD ${sg(S.fxShift)}${Math.abs(S.fxShift)}%</span>: exposure adjusts to <span class="hi">${fmtFull(tot)}</span>.`},
  ];
  const g=document.getElementById('insGrid'); g.innerHTML='';
  ins.forEach(i=>{const d=document.createElement('div');d.className='iitem';d.innerHTML=`<div class="illbl">${i.l}</div><div class="ilval">${i.v}</div>`;g.appendChild(d);});
}

// ── Page navigation ───────────────────────────────────────────────────────

/** Switch the visible page and re-render it. */
function switchPage(id) {
  document.querySelectorAll('.pc').forEach(p=>p.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  document.querySelectorAll('.pnav-btn').forEach((b,i)=>b.classList.toggle('active',['p1','p2','p3'][i]===id));
  if(id==='p2') renderP2();
  if(id==='p3') renderP3();
}

// ── Page 2: Entity Level Exposure Details ────────────────────────────────

function initP2() {
  buildPvtSels('pvtSel1', 1, 'YPOPL');
  buildPvtSels('pvtSel2', 2, 'YPOPL');
}

/** Build entity selector buttons for a pivot table. */
function buildPvtSels(cid, tbl, def) {
  const c=document.getElementById(cid); c.innerHTML='';
  [['YPOPL','YPOPL'],['YPNLBV','YPNL BV'],['YPAS','YPRODAS']].forEach(([k,l])=>{
    const b=document.createElement('button'); b.className='pvtsel'+(k===def?' active':''); b.dataset.key=k; b.textContent=l;
    b.onclick=()=>{c.querySelectorAll('.pvtsel').forEach(x=>x.classList.remove('active'));b.classList.add('active');pvtEnt[tbl]=k;renderPvt(tbl);};
    c.appendChild(b);
  });
}

function renderP2() { renderPies(); renderTrends(); renderPvt(1); renderPvt(2); }

/** Render the three pie charts (one per entity) on page 2. */
function renderPies() {
  [['YPOPL','pieYPOPL'],['YPNLBV','pieYPNLBV'],['YPAS','pieYPAS']].forEach(([entity,cid])=>{
    const recs=BSR.filter(r=>r.entity===entity&&(S.periods.length===0||S.periods.includes(r.period))&&mY(r)&&mA(r));
    const bc={}; recs.forEach(r=>{if(!bc[r.ccy])bc[r.ccy]=0;bc[r.ccy]+=Math.abs(r.ccy_amt);});
    const entries=Object.entries(bc).filter(e=>e[1]>0).sort((a,b)=>b[1]-a[1]);
    const lbls=entries.map(e=>e[0]),data=entries.map(e=>e[1]);
    const clrs=lbls.map((l,i)=>CCYCLR[l]||PAL[i%PAL.length]);
    destroyC(cid);
    charts[cid]=new Chart(document.getElementById(cid),{
      type:'pie',
      data:{labels:lbls,datasets:[{data,backgroundColor:clrs,borderColor:'#050609',borderWidth:2,hoverBorderWidth:3}]},
      options:{
        responsive:true,maintainAspectRatio:false,layout:{padding:{right:40}},
        plugins:{
          legend:{display:true,position:'right',labels:{color:'#fff',font:{size:9,weight:'700'},boxWidth:10,padding:4,
            generateLabels:ch=>{const d=ch.data,tot=d.datasets[0].data.reduce((a,b)=>a+b,0);return d.labels.map((l,i)=>({text:`${l} ${((d.datasets[0].data[i]/tot)*100).toFixed(1)}%`,fillStyle:d.datasets[0].backgroundColor[i],strokeStyle:'transparent',lineWidth:0,index:i,fontColor:'#fff'}));}}},
          tooltip:{callbacks:{label:c=>`${c.label}: ${(c.dataset.data[c.dataIndex]/c.dataset.data.reduce((a,b)=>a+b,0)*100).toFixed(1)}%`}},
          datalabels:{
            display:ctx=>{const tot=ctx.dataset.data.reduce((a,b)=>a+b,0);return ctx.dataset.data[ctx.dataIndex]/tot>0.08;},
            color:'#000',font:{weight:'900',size:8.5},
            formatter:(v,ctx)=>{const tot=ctx.dataset.data.reduce((a,b)=>a+b,0);return `${ctx.chart.data.labels[ctx.dataIndex]}\\n${fmtS(v)||'<1K'}\\n${(v/tot*100).toFixed(1)}%`;}
          }
        }
      }
    });
  });
}

/** Render the three stacked trend charts (gain/loss/net) for each entity. */
function renderTrends() {
  [['YPOPL','trendYPOPL'],['YPNLBV','trendYPNLBV'],['YPAS','trendYPAS']].forEach(([entity,cid])=>{
    const agg={}; PERIODS.forEach(p=>agg[p]={gain:0,loss:0,net:0});
    BSR.filter(r=>r.entity===entity&&mY(r)&&mA(r)).forEach(r=>{if(agg[r.period]){agg[r.period].gain+=r.u_gain;agg[r.period].loss+=r.u_loss;agg[r.period].net+=r.net_gl;}});
    destroyC(cid);
    charts[cid]=new Chart(document.getElementById(cid),{
      type:'line',
      data:{labels:PERIODS,datasets:[
        {label:'Unrealized Gain',data:PERIODS.map(p=>agg[p].gain),borderColor:'#00E676',backgroundColor:'rgba(0,230,118,.18)',borderWidth:2.5,fill:true,tension:.35,pointRadius:2},
        {label:'Unrealized Loss',data:PERIODS.map(p=>-agg[p].loss),borderColor:'#FF1744',backgroundColor:'rgba(255,23,68,.18)',borderWidth:2.5,fill:true,tension:.35,pointRadius:2},
        {label:'Net G/L',data:PERIODS.map(p=>agg[p].net),borderColor:'#3730A3',backgroundColor:'rgba(55,48,163,.22)',borderWidth:3.5,fill:true,tension:.35,pointRadius:3,pointBackgroundColor:'#3730A3'},
      ]},
      options:{
        responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
        plugins:{
          legend:{labels:{color:'#fff',font:{size:9,weight:'700'},boxWidth:12,padding:8}},
          tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${sg(c.raw)}${fmtFull(Math.abs(c.raw))}`}},
          datalabels:{display:ctx=>ctx.datasetIndex===2&&Math.abs(ctx.dataset.data[ctx.dataIndex])>1000,color:'#40C4FF',font:{size:8,weight:'bold',family:"'DM Mono',monospace"},formatter:v=>fmtS(v),anchor:'end',align:'top',offset:-2,clamp:true,clip:true}
        },
        scales:{
          x:{ticks:{color:'#fff',font:{size:8},maxRotation:45},grid:{color:'rgba(255,255,255,.06)'}},
          y:{ticks:{color:'#fff',font:{size:8},callback:v=>fmtS(v)||'0'},grid:{color:'rgba(255,255,255,.06)'}}
        }
      }
    });
  });
}

/** Render a single pivot table for the given table index (1 or 2). */
function renderPvt(tbl) {
  const entity=pvtEnt[tbl], vf=tbl===1?'ccy_amt':'net_gl', bid='pvtBdy'+tbl;
  const recs=BSR.filter(r=>r.entity===entity&&(S.periods.length===0||S.periods.includes(r.period))&&mY(r)&&mA(r));
  if(!recs.length){document.getElementById(bid).innerHTML='<div class="ldg" style="min-height:80px;font-size:.78rem">No data for selected filters.</div>';return;}
  const ccys=[...new Set(recs.map(r=>r.ccy))].sort();
  const rows={}; recs.forEach(r=>{if(!rows[r.account])rows[r.account]={account:r.account,vals:{}};if(!rows[r.account].vals[r.ccy])rows[r.account].vals[r.ccy]=0;rows[r.account].vals[r.ccy]+=(vf==='ccy_amt'?r.ccy_amt:r.net_gl);});
  const rl=Object.values(rows).sort((a,b)=>a.account.localeCompare(b.account));
  const tots={}; ccys.forEach(c=>tots[c]=0);
  rl.forEach(r=>ccys.forEach(c=>tots[c]+=(r.vals[c]||0)));
  let h=`<table class="pvttbl"><thead><tr><th>ACCOUNT</th>`;
  ccys.forEach(c=>h+=`<th>${c}</th>`); h+=`<th>TOTAL</th></tr></thead><tbody>`;
  rl.forEach(r=>{
    const rt=ccys.reduce((s,c)=>s+(r.vals[c]||0),0);
    if(!ccys.some(c=>r.vals[c]))return;
    h+=`<tr><td>${r.account}</td>`;
    ccys.forEach(c=>{const v=r.vals[c]||null;h+=`<td class="${v>0?'pos':v<0?'neg':''}">${v!=null?(fmtS(v)||'0'):'–'}</td>`;});
    h+=`<td class="${rt>0?'pos':rt<0?'neg':''}">${fmtS(rt)||'0'}</td></tr>`;
  });
  h+=`<tr class="totr"><td>TOTAL</td>`;let gt=0;
  ccys.forEach(c=>{h+=`<td class="${tots[c]>0?'pos':tots[c]<0?'neg':''}">${fmtS(tots[c])||'0'}</td>`;gt+=tots[c];});
  h+=`<td class="${gt>0?'pos':gt<0?'neg':''}">${fmtS(gt)||'0'}</td></tr>`;
  h+='</tbody></table>';
  document.getElementById(bid).innerHTML=h;
}

// ── Page 3: Variance Account Movement ────────────────────────────────────

/** Initialise legend scroll buttons for both page-3 charts. */
function initP3LegBtns() {
  [['p3L1P','p3Leg1',-1],['p3L1N','p3Leg1',1],['p3L2P','p3Leg2',-1],['p3L2N','p3Leg2',1]].forEach(([id,area,dir])=>{
    document.getElementById(id).onclick=()=>document.getElementById(area).scrollLeft+=dir*180;
  });
}

/** Render both page-3 charts and the detail table. */
function renderP3() {
  const vBS = filtV(VAR.filter(r=>r.acc_type==='Balance Sheet'&&r.ccy&&(r.is_ana==='CONSIDERED'||r.is_ana==='NOT LABELED')));
  const vMV = filtV(VAR.filter(r=>r.ccy&&(r.is_ana==='CONSIDERED'||r.is_ana==='NOT LABELED')));
  const months = VMONS.map(m=>mToLbl(m));
  renderP3Chart('p3C1','p3Leg1', vBS, months, 'closing',  false);
  renderP3Chart('p3C2','p3Leg2', vMV, months, 'movement', true);
  renderP3Table();
}

/**
 * Render a single page-3 line chart.
 * Automatically assigns outlier currencies (>8× median) to a secondary Y axis
 * to prevent scale compression.
 *
 * @param {string}  cid     - Canvas element ID
 * @param {string}  legId   - Legend container element ID
 * @param {Array}   vrecs   - Filtered variance records
 * @param {Array}   months  - Display labels for the X axis
 * @param {string}  field   - 'closing' or 'movement'
 * @param {boolean} stepped - If true, use stepped (before) line style
 */
function renderP3Chart(cid, legId, vrecs, months, field, stepped) {
  const byCcy={};
  vrecs.forEach(r=>{
    if(!byCcy[r.ccy])byCcy[r.ccy]={};
    if(!byCcy[r.ccy][r.month])byCcy[r.ccy][r.month]=0;
    byCcy[r.ccy][r.month]+=(field==='closing'?r.closing:r.movement);
  });
  const ccys=Object.keys(byCcy).sort();
  const vmons=VMONS;

  // Detect outlier currencies using max-absolute-value vs median heuristic
  const maxAbs={};
  ccys.forEach(ccy=>{ maxAbs[ccy]=Math.max(...vmons.map(m=>Math.abs(byCcy[ccy][m]||0))); });
  const allMaxVals=Object.values(maxAbs).filter(v=>v>0).sort((a,b)=>a-b);
  const med=allMaxVals[Math.floor(allMaxVals.length/2)]||1;
  const OUTLIER_THRESHOLD=8;
  const isOutlier=ccy=>maxAbs[ccy]>med*OUTLIER_THRESHOLD;
  const hasOutlier=ccys.some(isOutlier);

  const datasets=ccys.map((ccy,idx)=>{
    const clr=CCYCLR[ccy]||PAL[idx%PAL.length];
    const isActive=p3ActiveCcy===null||p3ActiveCcy===ccy;
    return{
      label:ccy,
      data:vmons.map(m=>byCcy[ccy][m]!==undefined?byCcy[ccy][m]:0),
      borderColor:clr, backgroundColor:clr+'18',
      borderWidth:isActive?2.5:1, pointRadius:isActive?3:1.5,
      tension:stepped?0:.35, fill:false,
      stepped:stepped?'before':false,
      spanGaps:true,
      yAxisID:hasOutlier&&isOutlier(ccy)?'y2':'y1',
    };
  });

  // Custom plugin: draw currency labels at the end of each visible line
  const endLabelPlugin={
    id:'endLbl_'+cid,
    afterDatasetsDraw(chart){
      const ctx=chart.ctx, lbls=[];
      chart.data.datasets.forEach((d,i)=>{
        const isActive=p3ActiveCcy===null||p3ActiveCcy===d.label;
        if(!isActive) return;
        const meta=chart.getDatasetMeta(i);
        const pts=meta.data.filter(p=>p&&!isNaN(p.y)&&p.x!=null);
        if(!pts.length) return;
        const last=pts[pts.length-1];
        lbls.push({x:last.x+5,y:last.y,text:d.label,color:d.borderColor,orig:last.y});
      });
      lbls.sort((a,b)=>a.orig-b.orig);
      // Anti-collision: push overlapping labels apart vertically
      const minGap=12;
      for(let i=1;i<lbls.length;i++){if(lbls[i].y-lbls[i-1].y<minGap)lbls[i].y=lbls[i-1].y+minGap;}
      lbls.forEach(l=>{ctx.save();ctx.fillStyle=l.color;ctx.font='bold 8.5px "DM Sans"';ctx.fillText(l.text,l.x,l.y+4);ctx.restore();});
    }
  };
  Chart.register(endLabelPlugin);

  const scalesConfig={
    x:{ticks:{color:'#fff',font:{size:9},maxRotation:45},grid:{color:'rgba(255,255,255,.06)'}},
    y1:{type:'linear',position:'left',ticks:{color:'#fff',font:{size:9},callback:v=>fmtS(v)||'0'},grid:{color:'rgba(255,255,255,.06)'},title:{display:true,text:'Primary (LCY)',color:'rgba(255,255,255,.55)',font:{size:8}}},
  };
  if(hasOutlier){
    scalesConfig.y2={type:'linear',position:'right',ticks:{color:'rgba(245,196,0,.85)',font:{size:9},callback:v=>fmtS(v)||'0'},grid:{drawOnChartArea:false},title:{display:true,text:'Secondary (LCY)',color:'rgba(245,196,0,.7)',font:{size:8}}};
  }

  destroyC(cid);
  charts[cid]=new Chart(document.getElementById(cid),{
    type:'line', data:{labels:months,datasets},
    options:{
      responsive:true, maintainAspectRatio:false, spanGaps:true,
      layout:{padding:{right:55}},
      interaction:{mode:'nearest',intersect:true},
      onClick:(evt,elements)=>{
        if(elements.length>0){
          const clickedCcy=datasets[elements[0].datasetIndex].label;
          p3ActiveCcy=(p3ActiveCcy===clickedCcy?null:clickedCcy);
          renderP3();
        }
      },
      plugins:{
        legend:{display:false},
        tooltip:{mode:'nearest',intersect:true,callbacks:{
          label:c=>{const outlier=hasOutlier&&isOutlier(c.dataset.label);return `${c.dataset.label}${outlier?' ★':''}: ${c.raw!=null?fmtFull(c.raw):'–'}`;},
          afterBody:ctx=>{if(!hasOutlier)return[];const out=ccys.filter(isOutlier);if(out.length)return[`★ Secondary axis: ${out.join(', ')}`];return[];}
        }},
        datalabels:{display:false}
      },
      scales:scalesConfig,
    }
  });

  // Build scrollable legend with axis indicator for outliers
  const leg=document.getElementById(legId); leg.innerHTML='';
  datasets.forEach(ds=>{
    const outlier=hasOutlier&&isOutlier(ds.label);
    const s=document.createElement('span'); s.className='tslegitem';
    s.style.opacity=p3ActiveCcy&&p3ActiveCcy!==ds.label?'0.4':'1';
    s.innerHTML=`<span class="tslegdot" style="background:${ds.borderColor}"></span>${ds.label}${outlier?'<span style="color:var(--gold);font-size:.58rem;margin-left:.15rem">★</span>':''}`;
    leg.appendChild(s);
  });
  if(hasOutlier){
    const note=document.createElement('span'); note.className='tslegitem';
    note.style.color='rgba(245,196,0,.65)'; note.style.fontSize='.6rem'; note.style.marginLeft='.5rem';
    note.textContent='★ = secondary axis (right)';
    leg.appendChild(note);
  }
}

/** Clear the active currency filter on page 3. */
function clearP3Filter() {
  p3ActiveCcy = null;
  document.getElementById('p3TblFilter').textContent = 'ALL CURRENCIES';
  document.getElementById('p3TblClr').classList.remove('show');
  renderP3();
}

/** Render the variance detail table on page 3. */
function renderP3Table() {
  const vAll    = filtV(VAR.filter(r=>r.ccy&&(r.is_ana==='CONSIDERED'||r.is_ana==='NOT LABELED')));
  const filtered = p3ActiveCcy ? vAll.filter(r=>r.ccy===p3ActiveCcy) : vAll;
  document.getElementById('p3TblFilter').textContent = p3ActiveCcy ? `${p3ActiveCcy} ONLY` : 'ALL CURRENCIES';
  document.getElementById('p3TblClr').classList.toggle('show', p3ActiveCcy!==null);
  const sorted=[...filtered].sort((a,b)=>a.month.localeCompare(b.month)||a.account.localeCompare(b.account));
  if(!sorted.length){document.getElementById('p3TblBody').innerHTML='<div class="ldg" style="min-height:80px;font-size:.78rem">No data for selected filters.</div>';return;}
  let h=`<table class="dtbl"><thead><tr><th>ACCOUNT</th><th>ACCOUNT TYPE</th><th>MONTH</th><th>CURRENCY</th><th>OPENING (t-1)</th><th>MOVEMENT (GL)</th><th>CLOSING (t)</th><th>VARIANCE %</th><th>MAIN DRIVER</th></tr></thead><tbody>`;
  sorted.forEach(r=>{
    const hl=p3ActiveCcy&&r.ccy===p3ActiveCcy?' highlighted':'';
    h+=`<tr class="${hl}"><td>${r.account}</td><td>${r.acc_type}</td><td>${mToLbl(r.month)}</td><td><span style="color:${CCYCLR[r.ccy]||'var(--gold)'}">▪</span> ${r.ccy}</td><td class="${r.opening>0?'pos':r.opening<0?'neg':''}">${r.opening?fmtFull(r.opening):'–'}</td><td class="${r.movement>0?'pos':r.movement<0?'neg':''}">${r.movement?fmtFull(r.movement):'–'}</td><td class="${r.closing>0?'pos':r.closing<0?'neg':''}">${r.closing?fmtFull(r.closing):'–'}</td><td class="${r.var_pct>0?'pos':r.var_pct<0?'neg':''}">${r.var_pct?(r.var_pct*100).toFixed(1)+'%':'–'}</td><td style="max-width:220px;text-overflow:ellipsis;overflow:hidden;white-space:nowrap" title="${r.driver||''}">${r.driver||'–'}</td></tr>`;
  });
  h+='</tbody></table>';
  document.getElementById('p3TblBody').innerHTML=h;
}

// ── Utility ───────────────────────────────────────────────────────────────

/** Destroy a Chart.js instance and remove it from the registry. */
function destroyC(id) { if(charts[id]){charts[id].destroy();delete charts[id];} }
</script>
</body>
</html>"""

    return HTML.replace('DATA_PAYLOAD', PJ)


def write_and_open(html: str, out_path: Path) -> None:
    """
    Write the HTML string to disk and open it in the default browser.

    Parameters
    ----------
    html     : str   - Complete HTML content
    out_path : Path  - Destination file path
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html, encoding='utf-8')
    open_in_browser(out_path)


# ── CLI entry point ──────────────────────────────────────────────────────────

def main() -> None:
    """
    Interactive command-line entry point.

    Workflow
    --------
    1. Print welcome banner
    2. Wait for user to press Enter
    3. Ask for the reporting period (year + month)
    4. Validate the maintenance gate silently
    5. Ask for the path to the master Excel file
    6. Load, process, and build the dashboard
    7. Save to Desktop and open in browser
    """
    # ── Welcome banner ──────────────────────────────────────────────────────
    banner = r"""
╔══════════════════════════════════════════════════════════════════════╗
║          YINSON PRODUCTION — TREASURY REPORTING TEAM                ║
║     Balance Sheet Exposure & Variance Analysis Dashboard            ║
║                    Generator  v4.0                                  ╠══
║                                                                     ║
║  This tool reads the master Excel workbook produced by the          ║
║  Treasury Reporting process and generates a fully interactive       ║
║  HTML dashboard covering:                                           ║
║                                                                     ║
║    • Executive Summary    (KPIs, drill-downs, FX sensitivity)       ║
║    • Entity Level Details (pie charts, G/L trends, pivot tables)    ║
║    • Variance Movement    (step charts, account detail table)       ║
║                                                                     ║
║  The output file is saved to your Desktop and opened automatically. ║
║  Author: Uygar Talu  |  Treasury Reporting Team                     ║
╚══════════════════════════════════════════════════════════════════════╝
"""
    print(banner)
    input("  Press ENTER to create the Balance Sheet Exposure & Variance Analysis Report…")

    # ── Step 1: Reporting period ─────────────────────────────────────────────
    print()
    while True:
        raw_period = input("  Which year and month are you updating the report for?\n"
                           "  (e.g. 2025 December, Dec 2025, 2025-12)  → ").strip()
        if not raw_period:
            print("  ⚠  Please enter a valid year and month.")
            continue
        try:
            year, month, label = parse_reporting_date(raw_period)
            print(f"\n  ✔  Reporting period identified: {_MONTH_SHORT[month]} {year}\n")
            break
        except ValueError as exc:
            print(f"  ⚠  Could not parse '{raw_period}': {exc}  —  please try again.")

    # ── Maintenance gate (silent) ────────────────────────────────────────────
    _check_maintenance_gate(year, month)

    # ── Step 2: Master data file path ────────────────────────────────────────
    print()
    while True:
        raw_path = input("  Please provide the file path to the master data Excel workbook:\n  → ").strip()
        if not raw_path:
            print("  ⚠  File path cannot be empty.")
            continue
        # Remove surrounding quotes that some terminals/explorers add
        raw_path = raw_path.strip('"').strip("'")
        xlsx_path = Path(raw_path)
        if not xlsx_path.exists():
            print(f"  ⚠  File not found: {xlsx_path}  —  please check the path and try again.")
            continue
        if xlsx_path.suffix.lower() not in ('.xlsx', '.xlsm', '.xls'):
            print(f"  ⚠  Expected an Excel file (.xlsx / .xlsm), got: {xlsx_path.suffix}")
            continue
        print(f"\n  ✔  File located: {xlsx_path.name}\n")
        break

    # ── Step 3: Load data ────────────────────────────────────────────────────
    print("  ⏳  Loading and processing data…")
    try:
        payload = load_data(str(xlsx_path))
    except Exception as exc:
        print(f"\n  ❌  Error reading Excel file: {exc}")
        sys.exit(1)
    print(f"  ✔  Data loaded  ({len(payload['bs'])} exposure records, "
          f"{len(payload['var'])} variance records)\n")

    # ── Step 4: Build HTML ───────────────────────────────────────────────────
    print("  ⏳  Building dashboard…")
    period_display = raw_period.strip()
    html = build_html(payload, period_display)

    # ── Step 5: Save to Desktop ──────────────────────────────────────────────
    filename  = f"BALANCE_SHEET_EXPOSURE_VARIANCE_ANALYSIS_{label}.html"
    out_path  = get_desktop_path() / filename
    write_and_open(html, out_path)

    size_kb = out_path.stat().st_size // 1024
    print(f"  ✔  Dashboard saved  →  {out_path}  ({size_kb} KB)")
    print(f"  ✔  Opening in your default browser…\n")
    print("  ✅  Done!  Share the HTML file with your team — no server required.\n")


if __name__ == '__main__':
    main()
