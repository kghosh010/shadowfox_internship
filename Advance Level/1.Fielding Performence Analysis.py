import random
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

WEIGHTS = {"CP":1, "GT":1, "C":3, "DC":-3, "ST":3, "RO":3, "MRO":-2, "DH":2}

cols = ["Match No.","Innings","Teams","Player Name","BallCount","Position","Pick","Throw","Runs","Overcount","Venue","Stadium"]

PICK_NAMES = {"Y":"Clean Pick","N":"Fumble","C":"Catch",
              "DC":"Dropped Catch","S":"Stumping"}
THROW_NAMES = {"Y":"Good Throw","N":"Bad Throw","DH":"Direct Hit","RO":"Run Out","MR":"Missed Run Out"}

FIELDERS = [
    ("Rilee Russouw","Short mid wicket"),("Phil Salt","Wicket keeper"),
    ("Yash Dhull","Covers"),("Axar Patel","Point"),
    ("Lalit Yadav","Cover point"),("Aman Khan","Long off"),
    ("Kuldeep Yadav","Fine leg"),("Mitchell Marsh","Mid on"),
    ("Anrich Nortje","Bowler"),("Khaleel Ahmed","Third man"),
    ("David Warner","Deep square leg"),
]

def build_demo(seed=2367, match="IPL2367",team="Delhi Capitals",
               venue="Delhi",stadium="Arun Jaitley Stadium"):
    rng = random.Random(seed)
    rows = []
    for over in range(1,21):
        for ball in range(1,7):
            ball_count = round((over - 1) + ball / 10,1)
            row = dict(zip(cols, [match,1,team,"",ball_count,"","","","",over,venue,stadium]))
            r = rng.random()
            if r < 0.30:  
                rows.append(row)
                continue
            name, pos = rng.choice(FIELDERS)
            row["Player Name"], row["Position"]=name,pos
            if pos == "Wicket keeper" and r > 0.97:
                row["Pick"], row["Runs"] = "S", 0
            elif r < 0.34:
                row["Pick"], row["Runs"] = "C", 0
            elif r < 0.37:
                row["Pick"], row["Runs"] = "DC", -rng.choice([2, 4, 6])
            elif r < 0.45:
                row["Pick"], row["Runs"] = "N", -rng.choice([1, 2])
            else:
                row["Pick"] = "Y"
                row["Runs"] = rng.choice([0, 0, 1, 1, 2])
                t = rng.random()
                if t < 0.45:
                    row["Throw"] = "Y"
                elif t < 0.52:
                    row["Throw"] = "N"
                    row["Runs"] = -1
                elif t < 0.60:
                    row["Throw"] = "DH"
                elif t < 0.64:
                    row["Throw"] = "RO"
                    row["Runs"] = 1
                elif t < 0.70:
                    row["Throw"] = "MR"
            rows.append(row)
    return pd.DataFrame(rows, columns=cols)

def score_players(df):
    d = df[df["Player Name"].astype(str).str.strip() != ""].copy()
    d["Pick"] = d["Pick"].astype(str).str.strip().str.upper().replace({"NAN": ""})
    d["Throw"] = d["Throw"].astype(str).str.strip().str.upper().replace({"NAN": ""})
    d["Runs"] = pd.to_numeric(d["Runs"], errors="coerce").fillna(0)

    g = d.groupby("Player Name")
    out = pd.DataFrame({
        "Clean Picks (CP)": g["Pick"].apply(lambda s: (s == "Y").sum()),
        "Good Throws (GT)": g["Throw"].apply(lambda s: (s == "Y").sum()),
        "Catches (C)": g["Pick"].apply(lambda s: (s == "C").sum()),
        "Dropped Catches (DC)": g["Pick"].apply(lambda s: (s == "DC").sum()),
        "Stumpings (ST)": g["Pick"].apply(lambda s: (s == "S").sum()),
        "Run Outs (RO)": g["Throw"].apply(lambda s: (s == "RO").sum()),
        "Missed Run Outs (MRO)": g["Throw"].apply(lambda s: (s == "MR").sum()),
        "Direct Hits (DH)": g["Throw"].apply(lambda s: (s == "DH").sum()),
        "Runs Saved (RS)": g["Runs"].sum(),
    })
    w = WEIGHTS
    out["Performance Score (PS)"] = (
        out["Clean Picks (CP)"] * w["CP"] + out["Good Throws (GT)"] * w["GT"]
        + out["Catches (C)"] * w["C"] + out["Dropped Catches (DC)"] * w["DC"]
        + out["Stumpings (ST)"] * w["ST"] + out["Run Outs (RO)"] * w["RO"]
        + out["Missed Run Outs (MRO)"] * w["MRO"] + out["Direct Hits (DH)"] * w["DH"]
        + out["Runs Saved (RS)"]
    )
    return out.sort_values("Performance Score (PS)", ascending=False).reset_index()

# 3.Excel output (matrix uses live COUNTIFS/SUMIFS + a PS formula)

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")

def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def autowidth(ws):
    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 2, 10), 28)

def write_workbook(df, players, path):
    wb = Workbook()

    ws = wb.active
    ws.title = "Fielding Data"
    ws.append(cols)
    for r in df.itertuples(index=False):
        ws.append(["" if pd.isna(v) else v for v in r])
    style_header(ws, 1, len(cols))
    ws.freeze_panes = "A2"
    autowidth(ws)
    last = ws.max_row

    wsw = wb.create_sheet("Weights")
    wsw.append(["Action", "Weight"])
    for k, v in WEIGHTS.items():
        wsw.append([k, v])
    wsw.append(["RS", "added as is"])
    style_header(wsw, 1, 2)
    autowidth(wsw)
    wrow = {k: i + 2 for i, k in enumerate(WEIGHTS)}  # weight cell rows

    wsm = wb.create_sheet("Performance Matrix")
    heads = ["Player Name", "Clean Picks (CP)", "Good Throws (GT)", "Catches (C)",
             "Dropped Catches (DC)", "Stumpings (ST)", "Run Outs (RO)",
             "Missed Run Outs (MRO)", "Direct Hits (DH)", "Runs Saved (RS)",
             "Performance Score (PS)"]
    wsm.append(heads)
    style_header(wsm, 1, len(heads))
    D = "'Fielding Data'!"
    rng = lambda col: f"{D}${col}$2:${col}${last}"
    for i, name in enumerate(players["Player Name"], start=2):
        wsm.cell(i, 1, name)
        wsm.cell(i, 2, f'=COUNTIFS({rng("D")},$A{i},{rng("G")},"Y")')
        wsm.cell(i, 3, f'=COUNTIFS({rng("D")},$A{i},{rng("H")},"Y")')
        wsm.cell(i, 4, f'=COUNTIFS({rng("D")},$A{i},{rng("G")},"C")')
        wsm.cell(i, 5, f'=COUNTIFS({rng("D")},$A{i},{rng("G")},"DC")')
        wsm.cell(i, 6, f'=COUNTIFS({rng("D")},$A{i},{rng("G")},"S")')
        wsm.cell(i, 7, f'=COUNTIFS({rng("D")},$A{i},{rng("H")},"RO")')
        wsm.cell(i, 8, f'=COUNTIFS({rng("D")},$A{i},{rng("H")},"MR")')
        wsm.cell(i, 9, f'=COUNTIFS({rng("D")},$A{i},{rng("H")},"DH")')
        wsm.cell(i, 10, f'=SUMIFS({rng("I")},{rng("D")},$A{i})')
        wsm.cell(i, 11, (f"=B{i}*Weights!$B${wrow['CP']}+C{i}*Weights!$B${wrow['GT']}"
                         f"+D{i}*Weights!$B${wrow['C']}+E{i}*Weights!$B${wrow['DC']}"
                         f"+F{i}*Weights!$B${wrow['ST']}+G{i}*Weights!$B${wrow['RO']}"
                         f"+H{i}*Weights!$B${wrow['MRO']}+I{i}*Weights!$B${wrow['DH']}+J{i}"))
    wsm.freeze_panes = "B2"
    autowidth(wsm)

    wsl = wb.create_sheet("Legend")
    legend = [["Pick code", "Meaning"]] + [[k, v] for k, v in PICK_NAMES.items()]
    legend += [[], ["Throw code", "Meaning"]] + [[k, v] for k, v in THROW_NAMES.items()]
    legend += [[], ["Runs", "+ runs saved, - runs conceded"]]
    for r in legend:
        wsl.append(r)
    style_header(wsl, 1, 2)
    style_header(wsl, 8, 2)
    autowidth(wsl)

    wb.save(path)

def main():
    if len(sys.argv) > 1:
        df = pd.read_excel(sys.argv[1], header=None)
        # locate header row of the sample-style sheet
        hdr = df.index[df.eq("Match No.").any(axis=1)][0]
        df.columns = df.iloc[hdr]
        df = df.iloc[hdr + 1:]
        df = df[[c for c in cols if c in df.columns]].dropna(how="all")
        df = df[df["Match No."].astype(str).str.startswith("IPL")] \
            if df["Match No."].astype(str).str.startswith("IPL").any() else df
    else:
        df = build_demo()

    players = score_players(df)
    print(players.to_string(index=False))
    out = "fielding_analysis_output.xlsx"
    write_workbook(df, players, out)
    print(f"\nSaved -> {out}")

if __name__ == "__main__":
    main()