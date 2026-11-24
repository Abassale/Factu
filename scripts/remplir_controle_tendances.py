"""Remplit le fichier de contrôle des tendances (lot xx vierge) pour le lot 12 EPN.

Usage : python3 -I build_lot12.py BLANK JDV_N ESSAI MODELE OUT [JDV_N1 CONTROLE_N1 [COMMENTAIRES.json]]
"""
import sys, copy, datetime, collections, json
import openpyxl

blank, jdv_n_path, essai_path, model_path, out = sys.argv[1:6]
jdv_n1_path = sys.argv[6] if len(sys.argv) > 6 else None
ctrl_n1_path = sys.argv[7] if len(sys.argv) > 7 else None
comments = json.load(open(sys.argv[8], encoding="utf-8")) if len(sys.argv) > 8 else {}
TARIFS_2024 = {"E50EPN": {0.15896}, "E60EPN": {1.45}, "A60EPN": {2.96}}

# Tarifs déjà validés dans le modèle (lot 06) : prix 2025 / 2026
TARIFS_2025 = {"E50EPN": {0.16532}, "E60EPN": {1.51}, "A60EPN": {3.08}}
TARIFS_2026 = {"E50EPN": {0.17192, 0.35759, 0.55833, 0.75449, 1.59332, 2.11345, 2.4386},
               "E60EPN": {1.57}, "A60EPN": {3.2}, "T20A": {0.123}, "T30A": {0.11},
               "T40A": {0.34}, "T50A": {0.11}}


def to_date(v):
    if isinstance(v, str):
        try:
            return datetime.datetime.strptime(v.strip(), "%d/%m/%Y")
        except ValueError:
            return v
    return v


def read_jdv(path, sheet=None):
    """Renvoie des lignes au format du modèle (colonnes A..P)."""
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    hdr = [str(h).strip().upper().replace(" ", "_") if h else "" for h in next(it)]
    idx = {h: i for i, h in enumerate(hdr)}
    def g(r, *names):
        for n in names:
            if n in idx:
                return r[idx[n]]
        return None
    rows = []
    for r in it:
        if r is None or all(v is None for v in r):
            continue
        ref = g(r, "PDS", "REF_ABONNEMENT")
        ref = str(ref).lstrip("'") if ref is not None else None
        rows.append([
            g(r, "FACTURE"), g(r, "CONTRAT"), ref, g(r, "REFAQUA"), g(r, "TITULAIRE"),
            g(r, "CODE_ARTICLE"), g(r, "NOM_ARTICLE"), to_date(g(r, "DÉBUT", "DEBUT")), to_date(g(r, "FIN")),
            g(r, "PU_HT", "PUHT"), g(r, "VOLBASE"), g(r, "VOLFACT"),
            g(r, "MNT_HT", "MNTHT"), g(r, "MNTMTTC"), g(r, "MNTMTVA"), g(r, "PÉRIODE", "PERIODE"),
        ])
    return rows


def cs(dst, src):
    """Copie le style d'une cellule d'un autre classeur."""
    dst.font = copy.copy(src.font); dst.fill = copy.copy(src.fill)
    dst.border = copy.copy(src.border); dst.alignment = copy.copy(src.alignment)
    dst.number_format = src.number_format
    return dst


def num(x):
    try:
        return float(str(x).replace(",", "."))
    except (TypeError, ValueError):
        return 0.0


jdv_n = read_jdv(jdv_n_path)
jdv_n1 = read_jdv(jdv_n1_path, "JDV FINAL") if jdv_n1_path else None

wb = openpyxl.load_workbook(blank)
model = openpyxl.load_workbook(model_path)

# ---------- JDV N ----------
ws = wb["JDV N"]
ref_row = model["JDV N"][2]
for i, row in enumerate(jdv_n, start=2):
    for j, v in enumerate(row, start=1):
        c = ws.cell(i, j, v)
        if j == 3:
            c.number_format = "@"
        elif j in (8, 9):
            c.number_format = "mm-dd-yy"

# ---------- JDV N-1 ----------
if jdv_n1:
    ws = wb["JDV N-1"]
    for i, row in enumerate(jdv_n1, start=2):
        for j, v in enumerate(row, start=1):
            c = ws.cell(i, j, v)
            if j == 3:
                c.number_format = "@"
            elif j in (8, 9):
                c.number_format = "mm-dd-yy"

# ---------- Tarifs (TCD : code > PU HT, somme MNT HT) ----------
ws = wb["Tarifs"]
mws = model["Tarifs"]
piv = collections.OrderedDict()
for r in sorted(jdv_n, key=lambda r: (r[5], num(r[9]))):
    piv.setdefault(r[5], collections.OrderedDict())
    piv[r[5]][num(r[9])] = piv[r[5]].get(num(r[9]), 0) + num(r[12])
ws["A3"], ws["B3"], ws["C3"], ws["D3"] = "Étiquettes de lignes", "Somme de MNT HT", 2025, 2026
for a in ("A3", "B3", "C3", "D3"):
    cs(ws[a], mws[a])
ok_style = mws["D11"]
row = 4
non_classes = []
total = 0
for code, pus in piv.items():
    cs(ws.cell(row, 1, code), mws["A4"])
    cs(ws.cell(row, 2, sum(pus.values())), mws["B4"])
    row += 1
    for pu, mnt in pus.items():
        cs(ws.cell(row, 1, pu), mws["A5"])
        cs(ws.cell(row, 2, mnt), mws["B5"])
        if pu in TARIFS_2025.get(code, ()):
            cs(ws.cell(row, 3, "ok"), ok_style)
        elif pu in TARIFS_2026.get(code, ()):
            cs(ws.cell(row, 4, "ok"), ok_style)
        elif pu in TARIFS_2024.get(code, ()):
            cs(ws.cell(row, 5, "tarif 2024 (rattrapage)"), ok_style)
        else:
            non_classes.append((code, pu, round(mnt, 2)))
        row += 1
    total += sum(pus.values())
cs(ws.cell(row, 1, "Total général"), mws["A29"])
cs(ws.cell(row, 2, total), mws["B29"])
for col in "ABCD":
    ws.column_dimensions[col].width = mws.column_dimensions[col].width

# ---------- Concordance VOL ----------
ws = wb["Concordance VOL"]
mws = model["Concordance VOL"]
agg = collections.OrderedDict()
for r in sorted(jdv_n, key=lambda r: r[5]):
    a = agg.setdefault(r[5], [0, 0, 0])
    a[0] += num(r[11]); a[1] += num(r[12]); a[2] += num(r[13])
hdr = ["Étiquettes de lignes", "Somme de VOLFACT", "Somme de MNT HT", "Somme de MNTMTTC"]
for j, h in enumerate(hdr, start=1):
    cs(ws.cell(3, j, h), mws.cell(3, j))
row = 4
for code, (v, ht, ttc) in agg.items():
    for j, val in enumerate([code, v, ht, ttc], start=1):
        cs(ws.cell(row, j, val), mws.cell(5, j))
    row += 1
tot = [sum(a[k] for a in agg.values()) for k in range(3)]
for j, val in enumerate(["Total général"] + tot, start=1):
    cs(ws.cell(row, j, val), mws.cell(12, j))
for col in "ABCD":
    ws.column_dimensions[col].width = mws.column_dimensions[col].width

# ---------- TES + Etudes EP / ASST (besoin du JDV N-1) ----------
if jdv_n1:
    ws = wb["TES"]
    i = 2
    for src, year in ((jdv_n, 2026), (jdv_n1, 2025)):
        for r in src:
            vals = r[:11] + ([r[11], None] if year == 2026 else [None, r[11]]) + r[12:16]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(i, j, v)
                if j == 3:
                    c.number_format = "@"
                elif j in (8, 9):
                    c.number_format = "mm-dd-yy"
            i += 1

    def etude(sheet, codes):
        w = wb[sheet]
        v25 = collections.defaultdict(float); v26 = collections.defaultdict(float)
        for r in jdv_n1:
            if r[5] in codes:
                v25[r[2]] += num(r[11])
        for r in jdv_n:
            if r[5] in codes:
                v26[r[2]] += num(r[11])
        refs = sorted(set(v25) | set(v26))
        f_style = w["D2"]._style
        for k, ref in enumerate(refs, start=2):
            w.cell(k, 1, ref).number_format = "@"
            w.cell(k, 2, v25[ref] if ref in v25 else None)
            w.cell(k, 3, v26[ref] if ref in v26 else None)
            c = w.cell(k, 4, f'=IF(C{k}="","sortant",IF(B{k}="","entrant",IF(AND(B{k}<>0,C{k}<>0),C{k}-B{k},"rien")))')
            c._style = copy.copy(f_style)
        # même formule que le modèle : variation hors effet entrants / sortants
        w["K5"] = "=K4-G4"
        w["L5"] = "=(((K3-G4)-K2)/K2)"
        return refs
    etude("Etude EP", {"E60EPN", "E60C"})
    etude("Etude ASST", {"A60EPN", "A60C"})

# ---------- CONTROLE CA ----------
ws = wb["CONTROLE CA"]
e = openpyxl.load_workbook(essai_path, read_only=True)
it = e["Listing Client"].iter_rows(values_only=True)
h = next(it)
iref, iville = h.index("REF_ABONNEMENT"), h.index("VILLE_PDI")
ville = {str(r[iref]): r[iville] for r in it if r[0] is not None}
refs_n = {r[2] for r in jdv_n}
communes = sorted({ville[x] for x in refs_n if x in ville})
it = e["Population Totale"].iter_rows(values_only=True)
next(it)
population = sum(1 for r in it if r[0] is not None)
it = e["Factures tous états"].iter_rows(values_only=True)
h = next(it)
roles = collections.Counter((r[0], r[1]) for r in it if r[0] is not None)
(exercice, role), _ = roles.most_common(1)[0]
it = e["Factures tous états"].iter_rows(values_only=True)
next(it)
date_calc = min(r[5] for r in it if r[0] is not None)

ws["R2"] = 12
ws["F4"] = ", ".join(communes)
ws["K28"] = datetime.datetime(date_calc.year, date_calc.month, date_calc.day)
ws["M28"] = len({r[0] for r in jdv_n})
ws["O28"] = f"{exercice}-{role}"
ws["Q28"] = population
ews = openpyxl.load_workbook(essai_path, data_only=True)["Medianes"]
ws["F34"], ws["H34"] = ews["C1"].value, ews["D1"].value   # abonnement EP
ws["F35"], ws["H35"] = ews["G1"].value, ews["H1"].value   # consommation EP

# ---------- Facturation N-1 (contrôle de l'an dernier + JDV N-1) ----------
if ctrl_n1_path:
    prev = openpyxl.load_workbook(ctrl_n1_path, read_only=True)["CONTROLE CA"]
    pv = {}
    for row_ in prev.iter_rows(min_row=26, max_row=36):
        for c in row_:
            if getattr(c, "coordinate", None):
                pv[c.coordinate] = c.value
    ws["K13"] = pv["K28"]                      # date de facturation N-1
    ws["M13"] = pv["M28"]                      # nb de factures émises N-1
    ws["O13"] = pv.get("M30") or 0             # dont mensus
    ws["Q13"] = pv["O28"]                      # train N-1
    ws["P14"] = pv["Q28"]                      # population N-1
    ws["F21"], ws["H21"] = pv["F34"], pv["H34"]
    ws["F22"], ws["H22"] = pv["F35"], pv["H35"]
if jdv_n1:
    def s_(codes, k):
        return sum(num(r[k]) for r in jdv_n1 if r[5] in codes)
    lignes = {60: ["E50E", "E50EPN"], 61: ["E60C", "E60EPN"], 62: ["P30A"], 63: ["P40A"],
              64: ["T40A"], 65: ["T20A"], 66: ["T30A"],
              79: ["A50C"], 80: ["A60C", "A60EPN"], 81: ["P50A"], 82: ["T50A"]}
    for r_, codes in lignes.items():
        ws[f"L{r_}"] = round(s_(codes, 13), 2)
        vol = s_(codes, 11)
        if r_ == 60:
            vol -= sum(num(r[11]) for r in jdv_n1 if r[5] == "E50E" and num(r[9]) == 0)
        ws[f"O{r_}"] = vol
for cell, txt in comments.items():
    ws[cell] = txt

wb.save(out)
print("lignes JDV N", len(jdv_n), "factures", ws["M28"].value, "population", population,
      "train", ws["O28"].value, "date", ws["K28"].value.date(), "communes", len(communes))
print("tarifs non classés", non_classes)
