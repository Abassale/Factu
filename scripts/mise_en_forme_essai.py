"""Met l'essai 2026 du lot 12 dans la présentation de l'essai 2025.

Usage : python3 -I mise_en_forme_essai.py ESSAI_2026 JDV_FINAL_2026 SORTIE
"""
import sys, datetime, collections
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

src, jdv_path, out = sys.argv[1:4]

FONT = "Aptos Narrow"
thin = Side(style="thin")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
DATE = "mm-dd-yy"


def fill(color=None, theme=None, tint=0.0, indexed=None):
    if indexed is not None:
        c = Color(indexed=indexed)
    elif theme is not None:
        c = Color(theme=theme, tint=tint)
    else:
        c = Color(rgb=color)
    return PatternFill(fill_type="solid", fgColor=c, bgColor=Color(indexed=64))


GREEN = fill("FF92D050")
YELLOW = fill("FFFFFF00")
GREY = fill(theme=2, tint=-0.0999786370433668)


def clean(v):
    if isinstance(v, str):
        return v.lstrip("'").strip()
    return v


def pds0(v):
    try:
        return int(str(v).lstrip("'").strip())
    except (TypeError, ValueError):
        return v


wb = openpyxl.load_workbook(src)

# ---------- données de référence ----------
lst = wb["Listing Client"]
lh = [c.value for c in lst[1]]
li = {h: i for i, h in enumerate(lh)}
listing = {}
for row in lst.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        continue
    listing[str(row[li["REF_ABONNEMENT"]]).lstrip("'")] = row

jw = openpyxl.load_workbook(jdv_path, read_only=True)
it = jw.worksheets[0].iter_rows(values_only=True)
jh = next(it)
ji = jh.index("REF_ABONNEMENT")
factures_finales = {str(r[ji]).lstrip("'") for r in it if r and r[0] is not None}

fte = wb["Factures tous états"]
role = fte["B2"].value
exercice = fte["A2"].value
train = f"{exercice}-{role}"
date_fin = max(r[5] for r in fte.iter_rows(min_row=2, values_only=True) if r[0] is not None)


def set_font(ws, size=11):
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None or c.has_style:
                f = copy_font(c.font, size)
                c.font = f


def copy_font(f, size):
    return Font(name=FONT, sz=size, b=f.b, i=f.i, u=f.u, strike=f.strike, color=f.color)


# ---------- 04.04 (ex « CTRL Factu ») ----------
ctrl = wb["CTRL Factu"]
ch = [c.value for c in ctrl[3]]
rows = [r for r in ctrl.iter_rows(min_row=4, values_only=True) if r[0] is not None]
ano_cols = list(range(ch.index("Nb Ano") + 1, len(ch)))          # colonnes d'anomalies 2026
ws = wb.create_sheet("04.04")
ws.sheet_view.showGridLines = False
ws["A1"] = "DETAIL DES CONTROLES FACTURES"
ws["A1"].font = Font(name=FONT, sz=18, b=True)
ws.row_dimensions[1].height = 23.4
ws["A2"] = f"Fin de traitement à {date_fin:%d/%m/%Y %H:%M}"
ws["A3"] = f"Train:{train} EPN LOT 12"

entetes = ["Ref. PDL", "Ref. PDS", "PDS 0", "Code Facture", "A SORTIR DU TRAIN", "Mode Facturation",
           "Offre Type", "Client", "Adresse", "Ville", "HT", "TVA", "TTC", "DELTA DATE DEBUT",
           "DELTA DATE FIN", "commentaire Abo", "Période Abo EP. Debut", "Période Abo EP. Fin",
           "commentaire date conso", "Période Conso EP. Debut", "Période Conso EP. Fin",
           "Période Conso Asst. Debut", "Période Conso Asst. Fin", "commentaire Conso",
           "Volume_eau_fact", "Volume_asst_fact", "DELTA VOLUME", "surconso",
           "METHODE DERNIERE RELEVE ", "CONSO DERNIERE RELEVE", "METHODE RELEVE 1",
           "CONSO RELEVE 1", "CONSO RELEVE 2", "Debut PDS", "Fin PDS", "Nb Ano"]
entetes += [ch[i] for i in ano_cols]
gris = {"DELTA DATE DEBUT", "DELTA DATE FIN", "DELTA VOLUME", "surconso", "METHODE DERNIERE RELEVE ",
        "CONSO DERNIERE RELEVE", "METHODE RELEVE 1", "CONSO RELEVE 1", "CONSO RELEVE 2"}
jaunes = {"commentaire Abo", "commentaire date conso", "commentaire Conso"}
for j, h in enumerate(entetes, start=1):
    c = ws.cell(6, j, h)
    c.font = Font(name=FONT, sz=11)
    c.border = BORDER
    c.alignment = Alignment(wrap_text=True)
    if h == "A SORTIR DU TRAIN":
        c.font = Font(name=FONT, sz=11, color="FFFF0000"); c.fill = YELLOW
    elif h in gris:
        c.fill = GREY
    elif h in jaunes:
        c.fill = YELLOW
ws.row_dimensions[6].height = 28.8

idx = {h: i for i, h in enumerate(ch)}
L = li  # colonnes du listing
first, last = 7, 7 + len(rows) - 1
for k, r in enumerate(rows):
    n = first + k
    eau_seule = r[idx["Offre Type"]] == "EVREUX - EAU"
    ref = clean(r[idx["Ref. PDS"]])
    lrow = listing.get(ref)
    vals = [clean(r[idx["Ref. PDL"]]), ref, pds0(ref), clean(r[idx["Code Facture"]]), None,
            r[idx["Mode Facturation"]], r[idx["Offre Type"]], clean(r[idx["Client"]]),
            clean(r[idx["Adresse"]]), r[idx["Ville"]], r[idx["HT"]], r[idx["TVA"]], r[idx["TTC"]],
            f"=Q{n}-T{n}" if eau_seule else f'=IF(AND(Q{n}=T{n},Q{n}=V{n},T{n}=V{n}),Q{n}-T{n},"problème date début")',
            f"=R{n}-U{n}" if eau_seule else f'=IF(AND(R{n}=U{n},R{n}=W{n},U{n}=W{n}),R{n}-U{n},"problème date fin")',
            None, r[idx["Période Abo EP. Debut"]], r[idx["Période Abo EP. Fin"]], None,
            r[idx["Période Conso EP. Debut"]], r[idx["Période Conso EP. Fin"]],
            r[idx["Période Conso Asst. Debut"]], r[idx["Période Conso Asst. Fin"]], None,
            r[idx["Volume_eau fact"]], r[idx["Volume_asst fact"]],
            "offre eau seule" if eau_seule else f"=Y{n}-Z{n}", None,
            lrow[L["METHODE_RELEVE"]] if lrow else None, lrow[L["CONSO_DERNIER_RELEVE"]] if lrow else None,
            lrow[L["METHODE_RELEVE_1"]] if lrow else None, lrow[L["CONSO_RELEVE_1"]] if lrow else None,
            lrow[L["CONSO_RELEVE_2"]] if lrow else None,
            r[idx["Debut PDS"]], r[idx["Fin PDS"]], r[idx["Nb Ano"]]]
    vals += [r[i] for i in ano_cols]
    for j, v in enumerate(vals, start=1):
        c = ws.cell(n, j, v)
        c.border = BORDER
        if isinstance(v, (datetime.datetime, datetime.date)):
            c.number_format = DATE
ncol = len(entetes)
lastc = get_column_letter(ncol)
ws.auto_filter.ref = f"A6:{lastc}{last}"
ws.freeze_panes = "C7"

# bloc des médianes (P3:W5), comme en 2025
med = [("MEDIANE OFFRE EAU + ASST", 3, fill(theme=7, tint=0.7999816888943144), "'POUR MEDIANE OFFRE EAU +ASST'",
        {"Q": "A1", "R": "B1", "T": "C1", "U": "D1", "V": "E1", "W": "F1"}),
       ("MEDIANE OFFRE EAU", 4, fill(theme=8, tint=0.5999938962981048), "'POUR MEDIANE OFFRE EAU'",
        {"Q": "A1", "R": "B1", "T": "C1", "U": "D1", "V": None, "W": None})]
for label, rr, f, sheet, cells in med:
    c = ws[f"P{rr}"]; c.value = label; c.fill = f; c.alignment = Alignment(wrap_text=True)
    for col, ref in cells.items():
        c = ws[f"{col}{rr}"]
        c.value = f"={sheet}!{ref}" if ref else None
        c.fill = f; c.number_format = DATE
f = fill(theme=9, tint=0.3999755851924192)
ws["P5"] = "MEDIANE GENERAL"; ws["P5"].fill = f; ws["P5"].alignment = Alignment(wrap_text=True)
for col in "QRTUVW":
    c = ws[f"{col}5"]; c.value = f"=MEDIAN({col}{first}:{col}{last})"; c.fill = f; c.number_format = DATE

largeurs = {"A": 11.0, "B": 15.6, "C": 13.3, "D": 14.3, "E": 24.0, "F": 19.3, "G": 18.9, "H": 45.7,
            "J": 39.3, "K": 9.0, "L": 7.0, "M": 11.7, "N": 20.6, "O": 18.4, "P": 25.6, "Q": 22.7,
            "R": 20.3, "S": 25.0, "T": 25.0, "U": 22.4, "V": 26.6, "W": 24.1, "X": 24.1, "Y": 18.4,
            "Z": 18.9, "AA": 18.9, "AB": 18.9, "AC": 18.9, "AH": 11.9, "AI": 7.6, "AJ": 7.4}
for k, w in largeurs.items():
    ws.column_dimensions[k].width = w
for j in range(37, ncol + 1):
    ws.column_dimensions[get_column_letter(j)].width = 31.1

# ---------- POUR MEDIANE OFFRE EAU +ASST / EAU ----------
def mediane(name, offre, cols, labels):
    w = wb.create_sheet(name)
    data = [r for r in rows if r[idx["Offre Type"]] == offre]
    end = 2 + len(data)
    for j, lab in enumerate(labels, start=1):
        cl = get_column_letter(j)
        c = w.cell(1, j, f"=MEDIAN({cl}3:{cl}{end})"); c.number_format = DATE
        c = w.cell(2, j, lab); c.border = BORDER; c.alignment = Alignment(wrap_text=True)
        w.column_dimensions[cl].width = 13
    w.row_dimensions[2].height = 43.2
    for k, r in enumerate(data, start=3):
        for j, col in enumerate(cols, start=1):
            c = w.cell(k, j, r[idx[col]]); c.border = BORDER; c.number_format = DATE
    return len(data)

lab6 = ["Période Abo EP. Debut", "Période Abo EP. Fin", "Période Conso EP. Debut",
        "Période Conso EP. Fin", "Période Conso Asst. Debut", "Période Conso Asst. Fin"]
n_ea = mediane("POUR MEDIANE OFFRE EAU +ASST", "EVREUX - EAU + ASST", lab6, lab6)
n_e = mediane("POUR MEDIANE OFFRE EAU", "EVREUX - EAU", lab6[:4], lab6[:4])

# ---------- NON FACT APRES CALCUL / NON FACT FINAL ----------
NF_HDR = ["PDS", "PDS 0", "Titulaire", "Details", "Commentaire agent", "Début PDS", "Fin PDS",
          "Groupe de facturation", "Type d'offre", "Ville PDL"]
NF_W = [10.0, 13.7, 27.4, 11.9, 16.7, 10.1, 9.0, 18.1, 17.3, 22.9, 37.6]


def non_fact(name, items, last_label, marquer):
    w = wb.create_sheet(name)
    hdr = NF_HDR + [last_label]
    for j, h in enumerate(hdr, start=1):
        c = w.cell(1, j, h)
        c.font = Font(name=FONT, sz=10, b=True)
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        w.column_dimensions[get_column_letter(j)].width = NF_W[j - 1]
    w.row_dimensions[1].height = 27.6
    for k, (ref, titulaire, debut, fin) in enumerate(items, start=2):
        lrow = listing.get(ref)
        ok = marquer and ref in factures_finales
        vals = [ref, pds0(ref), titulaire, "Non facturé(s)", None, debut, fin,
                lrow[L["GROUPE_FACTURATION"]] if lrow else "EPN Masse Lot 12",
                lrow[L["TYPE_OFFRE"]] if lrow else None, lrow[L["VILLE_PDI"]] if lrow else None,
                "facture ok" if ok else None]
        for j, v in enumerate(vals, start=1):
            c = w.cell(k, j, v)
            c.font = Font(name=FONT, sz=11 if j in (5, 11) else 10)
            if j < 11:
                c.alignment = Alignment(wrap_text=True)
            if isinstance(v, (datetime.datetime, datetime.date)):
                c.number_format = DATE
            if ok:
                c.fill = GREEN
    w.auto_filter.ref = f"A1:{get_column_letter(len(hdr))}{max(2, len(items) + 1)}"
    return w

asf = wb["Abonnements sans facture"]
sans = [(clean(r[0]), clean(r[3]), r[6], r[7]) for r in asf.iter_rows(min_row=2, values_only=True) if r[0]]
non_fact("NON FACT APRES CALCUL", sans, "COMMENTAIRE", True)

pop = wb["Population Totale"]
popl = [(clean(r[0]), clean(r[3]), r[6], r[7]) for r in pop.iter_rows(min_row=2, values_only=True) if r[0]]
finals = [p for p in popl if p[0] not in factures_finales]
non_fact("NON FACT FINAL", finals, "justification ", False)

# ---------- ANO AVANT CALCUL (ex « PDI en ano ») ----------
ano = wb["PDI en ano"]
ano.title = "ANO AVANT CALCUL"
set_font(ano)
n_ano = ano.max_row
for j, h in ((11, "A TRAITER"), (12, "JUSTIFICATION")):
    c = ano.cell(1, j, h)
    c.font = Font(name="Calibri", sz=11, b=True)
    c.fill = fill(indexed=22); c.border = BORDER
ano.auto_filter.ref = f"A1:L{n_ano}"
for k, w in {"A": 17.0, "E": 12.0, "F": 55.0, "G": 12.0, "H": 13.0, "I": 30.0, "J": 9.1, "K": 12.1, "L": 37.0}.items():
    ano.column_dimensions[k].width = w

# ---------- autres onglets renommés ----------
tcd = wb["Population"]
tcd.title = "TCD LISTING"
for k, w in {"A": 28.9, "B": 24.4, "C": 12.6, "D": 4.0, "E": 12.6}.items():
    tcd.column_dimensions[k].width = w

lst.title = "LISTING ACTIFS"
lst.freeze_panes = "AE2"          # comme 2025 : en-tête + colonnes jusqu'au mode de paiement
for p in tcd._pivots:
    p.cache.cacheSource.worksheetSource.sheet = "LISTING ACTIFS"

pop.title = "POPULATION TRAIN"
set_font(pop)
fte.title = "FACTURES TOUS ETATS"
set_font(fte)
m2a = wb["Déchargement M2A"]
m2a.title = "DECHARGEMENT M2A"

# ---------- onglets remplacés ----------
for name in ("CTRL Factu", "Medianes", "Abonnements sans facture"):
    del wb[name]

ordre = ["TCD LISTING", "ANO AVANT CALCUL", "NON FACT APRES CALCUL", "POUR MEDIANE OFFRE EAU +ASST",
         "POUR MEDIANE OFFRE EAU", "04.04", "LISTING ACTIFS", "POPULATION TRAIN", "NON FACT FINAL",
         "FACTURES TOUS ETATS", "DECHARGEMENT M2A"]
wb._sheets = [wb[n] for n in ordre]
wb.active = 0
for w in wb.worksheets:
    w.sheet_view.tabSelected = (w.title == ordre[0])

wb.save(out)
print("04.04 lignes", len(rows), "| médianes EAU+ASST", n_ea, "EAU", n_e,
      "| non fact après calcul", len(sans), "dont facturés au final", sum(1 for s in sans if s[0] in factures_finales),
      "| non fact final", len(finals), "| train", train, "| fin", date_fin)
