#!/usr/bin/env python3
"""
export_strava.py
Gebruik: python export_strava.py [pad_naar_excel] [output_json]

Standaard:
  excel:  strava_analyse.xlsx
  output: strava_data.json

Zet dit script in dezelfde map als je Excel bestand.
Run het elke keer als je Excel is bijgewerkt, of automatiseer via Task Scheduler / cron.
"""

import sys
import json
import numpy as np
import pandas as pd
import openpyxl
from pathlib import Path

EXCEL_PATH = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("strava_analyse.xlsx")
OUTPUT_PATH = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("strava_data.json")

def tempo_to_decimal(t):
    try:
        parts = str(t).split(":")
        return int(parts[0]) + int(parts[1]) / 60
    except:
        return None

def min_to_time(m):
    if not m or (isinstance(m, float) and np.isnan(m)) or m == 0:
        return "—"
    h, rem = divmod(int(m), 60)
    s = int((m % 1) * 60)
    return f"{h}:{rem:02d}:{s:02d}" if h else f"{rem}:{s:02d}"

def clean(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    return str(v).strip()

def nan_safe_int(v):
    try:
        f = float(v)
        return int(f) if not np.isnan(f) else 0
    except:
        return 0

print(f"[>>] Excel inladen: {EXCEL_PATH}")

# ── 1. Hyperlinks via openpyxl ──────────────────────────────────────────────
wb = openpyxl.load_workbook(EXCEL_PATH)
ws = wb["Trainingsdata"]
header_row = 11
cols = {cell.value: cell.column - 1 for cell in ws[header_row] if cell.value}

link_map = {}
for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
    seg = str(row[cols.get("Segment", 4)].value or "")
    if "Totale Run" not in seg:
        continue
    datum = str(row[cols.get("Datum", 1)].value or "").strip()
    naam = str(row[cols.get("Run Naam", 2)].value or "").strip()
    link_cell = row[cols.get("Link", 3)]
    url = link_cell.hyperlink.target if link_cell.hyperlink else ""
    key = f"{datum}|{naam}"
    if key not in link_map:
        link_map[key] = []
    if url:
        link_map[key].append(url)

# ── 2. Data via pandas ───────────────────────────────────────────────────────
df = pd.read_excel(EXCEL_PATH, sheet_name="Trainingsdata", header=10)
df["Afstand (km)"] = pd.to_numeric(df["Afstand (km)"], errors="coerce")
df["Gem HR"] = pd.to_numeric(df["Gem HR"], errors="coerce")
df["Min HR"] = pd.to_numeric(df["Min HR"], errors="coerce")
df["Max HR"] = pd.to_numeric(df["Max HR"], errors="coerce")
df["Run Naam"] = df["Run Naam"].ffill()
df["Datum"] = df["Datum"].ffill()
df["Is_totaal"] = df["Segment"].str.strip() == "🏃\u200d♂️ Totale Run"
df["Run_afstand"] = np.where(df["Is_totaal"], df["Afstand (km)"], np.nan)
df["Run_afstand"] = df["Run_afstand"].ffill()

# ── 3. Rondes per run (voor HR@5:00 berekening) ──────────────────────────────
rondes = df[df["Segment"].str.strip().str.startswith("Ronde", na=False)].copy()
rondes["Ronde_nr"] = rondes["Segment"].str.extract(r"(\d+)").astype(float)
rondes["Tempo_dec"] = rondes["Tempo (min/km)"].apply(tempo_to_decimal)

SLOPE = -9.71
DRIFT = 0.29

hr500_map = {}
for (datum, naam), group in rondes.groupby(["Datum", "Run Naam"]):
    run_afstand = group["Run_afstand"].iloc[0]
    if run_afstand <= 8:
        continue
    g_all = group[group["Ronde_nr"] >= 3].copy()
    if any(g_all["Afstand (km)"].fillna(1) < 0.4) and (g_all["Afstand (km)"] < 0.4).sum() >= 3:
        continue
    g = g_all[
        (g_all["Tempo_dec"] >= 4.0) &
        (g_all["Tempo_dec"] <= 6.0) &
        g_all["Gem HR"].notna()
    ]
    if len(g) < 3:
        continue
    gem_tempo = g["Tempo_dec"].mean()
    gem_hr = g["Gem HR"].mean()
    midpoint_km = (3 + run_afstand) / 2
    drift_corr = -(midpoint_km - 3) * DRIFT
    hr_500 = gem_hr + drift_corr + SLOPE * (5.0 - gem_tempo)
    if 100 < hr_500 < 185:
        hr500_map[f"{datum}|{naam}"] = round(hr_500, 1)

# ── 4. Alle runs bouwen ───────────────────────────────────────────────────────
totaal = df[df["Is_totaal"] & (df["Afstand (km)"] > 0.5)].copy()
totaal["Tempo_dec"] = totaal["Tempo (min/km)"].apply(tempo_to_decimal)
totaal["Tijd_min"] = totaal["Tempo_dec"] * totaal["Afstand (km)"]
totaal["Tijd_str"] = totaal["Tijd_min"].apply(min_to_time)

alle_runs = []
for _, r in totaal.sort_values("Datum", ascending=False).iterrows():
    datum = clean(r["Datum"])
    naam = clean(r["Run Naam"])
    key = f"{datum}|{naam}"
    urls = link_map.get(key, [])
    url = urls.pop(0) if urls else ""
    link_map[key] = urls
    afstand = float(r["Afstand (km)"]) if not np.isnan(r["Afstand (km)"]) else 0
    alle_runs.append({
        "datum": datum,
        "datum_sort": datum.split(",")[0].strip(),
        "naam": naam,
        "afstand": round(afstand, 2),
        "tempo": clean(r["Tempo (min/km)"]),
        "tempo_dec": round(r["Tempo_dec"], 3) if r["Tempo_dec"] else 0,
        "tijd": r["Tijd_str"],
        "min_hr": nan_safe_int(r["Min HR"]),
        "gem_hr": nan_safe_int(r["Gem HR"]),
        "max_hr": nan_safe_int(r["Max HR"]),
        "zone": clean(r["HR Zone"]),
        "trainingsplan": clean(r["Trainingsplan"]),
        "strava_url": url,
        "hr_500": hr500_map.get(key, None),
    })

# ── 5. Duurlopen & wedstrijden categoriseren ─────────────────────────────────
def is_wedstrijd(r):
    naam = r["naam"].lower()
    plan = r["trainingsplan"].lower()
    kw = ["marathon rotterdam", "haas", "silvester", "schoorl", "cpc 21", "vaartspel spanderswoud"]
    return any(k in naam or k in plan for k in kw)

def is_duurloop(r):
    return (not is_wedstrijd(r) and r["afstand"] >= 8 and
            ("Z1" in r["zone"] or "Z2" in r["zone"] or "duurloop" in r["trainingsplan"].lower()))

duurlopen = [r for r in alle_runs if is_duurloop(r)]
hr500_list = [{"datum": r["datum_sort"], "hr500": r["hr_500"], "afstand": r["afstand"],
               "gem_hr": r["gem_hr"], "gem_tempo": r["tempo_dec"]}
              for r in duurlopen if r["hr_500"]]

# ── 6. Wedstrijd events ────────────────────────────────────────────────────────
events = [
    {"naam": "Marathon Rotterdam", "datum": "12 april 2026", "afstand": 42.2,
     "runs": [r for r in alle_runs if "marathon rotterdam" in r["naam"].lower() or "marathondag" in r["trainingsplan"].lower()]},
    {"naam": "Vaartspel Spanderswoud", "datum": "4 april 2026", "afstand": 11.13,
     "runs": [r for r in alle_runs if "vaartspel spanderswoud" in r["trainingsplan"].lower() and "2026-04-04" in r["datum_sort"]]},
    {"naam": "CPC 21,1 km", "datum": "15 maart 2026", "afstand": 21.27,
     "runs": [r for r in alle_runs if "cpc 21" in r["trainingsplan"].lower()]},
    {"naam": "Schoorl 10 km", "datum": "8 februari 2026", "afstand": 10.02,
     "runs": [r for r in alle_runs if "schoorl" in r["trainingsplan"].lower()]},
    {"naam": "Silvestercross", "datum": "31 december 2025", "afstand": 7.28,
     "runs": [r for r in alle_runs if "silvestercross" in r["trainingsplan"].lower()]},
    {"naam": "Haas Ultramarathon", "datum": "19 oktober 2025", "afstand": 42.41,
     "runs": [r for r in alle_runs if "haas" in r["naam"].lower()]},
]
for e in events:
    race = max((r for r in e["runs"] if r["gem_hr"] > 100), key=lambda x: x["gem_hr"], default=None)
    if race:
        e["gem_hr"] = race["gem_hr"]
        e["max_hr"] = race["max_hr"]
        e["tempo"] = race["tempo"]
        e["tijd"] = race["tijd"]
        e["strava_url"] = race["strava_url"]

# ── 7. PR's ────────────────────────────────────────────────────────────────────
prs = [
    {"afstand": "1,5 km", "tijd": "4:22", "tempo": "2:55", "datum": ""},
    {"afstand": "5 km", "tijd": "16:53", "tempo": "3:23", "datum": ""},
    {"afstand": "10 km", "tijd": "32:12", "tempo": "3:13", "datum": ""},
    {"afstand": "15 km", "tijd": "49:58", "tempo": "3:20", "datum": ""},
    {"afstand": "16,1 km", "tijd": "55:43", "tempo": "3:27", "datum": ""},
    {"afstand": "21,1 km", "tijd": "1:11:15", "tempo": "3:22", "datum": ""},
    {"afstand": "42,2 km", "tijd": "2:36:07", "tempo": "3:41", "datum": "Rotterdam 2026"},
]

# ── 8. Exporteer JSON ─────────────────────────────────────────────────────────
output = {
    "gegenereerd_op": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
    "alle": alle_runs,
    "duurlopen": duurlopen,
    "hr500": hr500_list,
    "events": events,
    "prs": prs,
}

with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"[OK] Klaar!")
print(f"   Runs:       {len(alle_runs)}")
print(f"   Duurlopen:  {len(duurlopen)}")
print(f"   HR@5:00:    {len(hr500_list)}")
print(f"   Output:     {OUTPUT_PATH}")