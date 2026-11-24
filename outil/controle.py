"""Remplissage du fichier « Contrôle des tendances » à partir des exports."""
import collections, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill
from lecture import num, ref, to_date, lignes_jdv

DATE = "mm-dd-yy"
OK_FONT = Font(name="Calibri", sz=11, color=openpyxl.styles.colors.Color(theme=9))
VERT = PatternFill(fill_type="solid", fgColor="FF00B050")

# Lignes du tableau EAU POTABLE / ASSAINISSEMENT de l'onglet CONTROLE CA
LIGNES = {60: ["E50E", "E50EPN"], 61: ["E60C", "E60EPN"], 62: ["P30A"], 63: ["P40A"],
          64: ["T40A"], 65: ["T20A"], 66: ["T30A"],
          79: ["A50C"], 80: ["A60C", "A60EPN"], 81: ["P50A"], 82: ["T50A"]}
COMMENTAIRE = {60: 68, 61: 69, 62: 70, 63: 71, 64: 72, 65: 73, 66: 74, 79: 84, 80: 85, 81: 86, 82: 87}
CODES_EP = {"E60EPN", "E60C"}
CODES_ASST = {"A60EPN", "A60C"}


# ---------- petites aides de mise en forme française ----------
def fr(x, dec=0):
    s = f"{x:,.{dec}f}".replace(",", "\u00a0").replace(".", ",")
    return s


def frs(x, dec=0):
    return ("+" if x > 0 else "") + fr(x, dec)


def pct(x):
    return ("+" if x > 0 else "") + f"{x * 100:.1f}".replace(".", ",") + " %"


def pu_txt(p):
    return f"{p:g}".replace(".", ",")


def ecrire_jdv(ws, lignes):
    for i, row in enumerate(lignes, start=2):
        for j, v in enumerate(row, start=1):
            c = ws.cell(i, j, v)
            if j == 3:
                c.number_format = "@"
            elif j in (8, 9):
                c.number_format = DATE


def somme(lignes, codes, k):
    return sum(num(r[k]) for r in lignes if r[5] in codes)


def pu_dominant(lignes, codes):
    c = collections.Counter()
    for r in lignes:
        if r[5] in codes and num(r[9]):
            c[num(r[9])] += max(num(r[11]), 1)
    return c.most_common(1)[0][0] if c else None


def jours_abo(lignes):
    v = somme(lignes, {"E50E", "E50EPN"}, 11)
    v -= sum(num(r[11]) for r in lignes if r[5] == "E50E" and num(r[9]) == 0)
    return v


def entrants_sortants(n, n1, codes):
    v26, v25 = collections.defaultdict(float), collections.defaultdict(float)
    for r in n:
        if r[5] in codes:
            v26[r[2]] += num(r[11])
    for r in n1:
        if r[5] in codes:
            v25[r[2]] += num(r[11])
    ent = sum(v for k, v in v26.items() if k not in v25)
    sor = sum(v for k, v in v25.items() if k not in v26)
    return ent, sor, v25, v26


def commentaires(n, n1, an, nb_n, nb_n1, es):
    """Pré-rédige les commentaires « VALIDATION DES EVOLUTIONS » à partir des chiffres."""
    out = {}
    for ligne, codes in LIGNES.items():
        cell = f"H{COMMENTAIRE[ligne]}"
        ca, ca1 = somme(n, codes, 13), somme(n1, codes, 13)
        vol, vol1 = somme(n, codes, 11), somme(n1, codes, 11)
        p, p1 = pu_dominant(n, codes), pu_dominant(n1, codes)
        if not ca and not ca1:
            out[cell] = "N/A"
            continue
        if not ca1:
            out[cell] = f"Nouvelle ligne facturée en {an}"
            continue
        if not ca:
            out[cell] = "N'est plus facturé"
            continue
        tarif = ""
        unite = "€/jour" if ligne in (60, 79) else "€/m³"
        if p and p1 and abs(p - p1) > 1e-9:
            sens = "Hausse" if p > p1 else "Baisse"
            fort = "Forte " + sens.lower() if (p / p1 > 2 or p1 / p > 2) else sens
            tarif = f"{fort} du tarif en {an} ({pu_txt(p1)} → {pu_txt(p)} {unite})"
        if ligne == 60:
            j, j1 = jours_abo(n), jours_abo(n1)
            df = nb_n - nb_n1
            debut = (f"{fr(abs(df))} factures en {'moins' if df < 0 else 'plus'} ({fr(nb_n)} contre {fr(nb_n1)})"
                     if df else f"Même nombre de factures ({fr(nb_n)})")
            txt = f"{debut} : {frs(j - j1)} jours d'abonnement facturés ({pct((j - j1) / j1) if j1 else 'n.c.'})"
            out[cell] = txt + (f". {tarif}" if tarif else ". Tarif inchangé")
        elif ligne in (61, 80):
            ent, sor = es["EP" if ligne == 61 else "ASST"]
            nom = "Etude EP" if ligne == 61 else "Etude ASST"
            txt = f"Volume {frs(vol - vol1)} m³ ({pct((vol - vol1) / vol1) if vol1 else 'n.c.'}), dont entrants / sortants : {frs(ent - sor)} m³ (voir {nom})"
            out[cell] = txt + (f". {tarif}" if tarif else ". Tarif inchangé")
        else:
            if tarif:
                out[cell] = tarif
            else:
                out[cell] = f"Tarif inchangé ({pu_txt(p)} {unite}) : évolution identique au volume" if p else "Tarif inchangé"
    return out


def lire_controle_n1(path):
    """Infos de la facturation précédente + son JDV (onglet « JDV N » du contrôle de l'an dernier)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["CONTROLE CA"]
    v = {}
    for row in ws.iter_rows(min_row=1, max_row=36):
        for c in row:
            if getattr(c, "coordinate", None) and c.value is not None:
                v[c.coordinate] = c.value
    from lecture import Table
    jw = wb["JDV N"]
    it = jw.iter_rows(values_only=True)
    hdr = list(next(it))
    rows = [list(r) for r in it if r and r[0] is not None]
    jdv = lignes_jdv(Table("jdv", path, "JDV N", hdr, rows))
    wb.close()
    return v, jdv


def remplir(template, out, d, log):
    """d : dictionnaire préparé par factu.py (lignes JDV, infos facturation, N-1…)."""
    n, n1 = d["jdv_n"], d.get("jdv_n1")
    an, an1 = d["annee"], d["annee"] - 1
    wb = openpyxl.load_workbook(template)

    ecrire_jdv(wb["JDV N"], n)
    if n1:
        ecrire_jdv(wb["JDV N-1"], n1)

    # --- Tarifs : TCD code > PU, somme MNT HT ; année du tarif déduite des dates facturées ---
    ws = wb["Tarifs"]
    ws["A3"], ws["B3"], ws["C3"], ws["D3"] = "Étiquettes de lignes", "Somme de MNT HT", an1, an
    piv = collections.defaultdict(lambda: collections.defaultdict(lambda: [0.0, collections.Counter()]))
    for r in n:
        e = piv[r[5]][num(r[9])]
        e[0] += num(r[12])
        fin = to_date(r[8])
        if isinstance(fin, datetime.datetime):
            e[1][fin.year] += 1
    row, total, alertes = 4, 0.0, []
    for code in sorted(piv, key=str):
        pus = piv[code]
        ws.cell(row, 1, code).alignment = openpyxl.styles.Alignment(horizontal="left")
        ws.cell(row, 2, sum(e[0] for e in pus.values()))
        row += 1
        for pu in sorted(pus):
            mnt, years = pus[pu]
            ws.cell(row, 1, pu).alignment = openpyxl.styles.Alignment(horizontal="left", indent=1)
            ws.cell(row, 2, mnt)
            if years:
                y, k = years.most_common(1)[0]
                if k < 0.9 * sum(years.values()):
                    ws.cell(row, 5, "à vérifier : périodes sur plusieurs années").font = OK_FONT
                    alertes.append(f"tarif {code} {pu_txt(pu)} : périodes sur plusieurs années")
                elif y == an:
                    ws.cell(row, 4, "ok").font = OK_FONT
                elif y == an1:
                    ws.cell(row, 3, "ok").font = OK_FONT
                else:
                    ws.cell(row, 5, f"tarif {y} (rattrapage)").font = OK_FONT
            row += 1
        total += sum(e[0] for e in pus.values())
    ws.cell(row, 1, "Total général").alignment = openpyxl.styles.Alignment(horizontal="left")
    ws.cell(row, 2, total)
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["E"].width = 34

    # --- Concordance VOL ---
    ws = wb["Concordance VOL"]
    for j, h in enumerate(["Étiquettes de lignes", "Somme de VOLFACT", "Somme de MNT HT", "Somme de MNTMTTC"], 1):
        ws.cell(3, j, h)
    agg = collections.defaultdict(lambda: [0.0, 0.0, 0.0])
    for r in n:
        a = agg[r[5]]
        a[0] += num(r[11]); a[1] += num(r[12]); a[2] += num(r[13])
    row = 4
    for code in sorted(agg, key=str):
        for j, val in enumerate([code] + agg[code], 1):
            c = ws.cell(row, j, val)
            if row == 4:
                c.fill = VERT
        row += 1
    for j, val in enumerate(["Total général"] + [sum(a[k] for a in agg.values()) for k in range(3)], 1):
        ws.cell(row, j, val)
    for col, w in zip("ABCD", (20, 18, 18, 20)):
        ws.column_dimensions[col].width = w

    # --- TES + études entrants / sortants ---
    es = {"EP": (0, 0), "ASST": (0, 0)}
    if n1:
        ws = wb["TES"]
        ws["L1"], ws["M1"] = f"VOLFACT {an}", f"VOLFACT {an1}"
        i = 2
        for src, courant in ((n, True), (n1, False)):
            for r in src:
                vals = r[:11] + ([r[11], None] if courant else [None, r[11]]) + r[12:16]
                for j, v in enumerate(vals, 1):
                    c = ws.cell(i, j, v)
                    if j == 3:
                        c.number_format = "@"
                    elif j in (8, 9):
                        c.number_format = DATE
                i += 1
        for sheet, codes, key in (("Etude EP", CODES_EP, "EP"), ("Etude ASST", CODES_ASST, "ASST")):
            w = wb[sheet]
            ent, sor, v25, v26 = entrants_sortants(n, n1, codes)
            es[key] = (ent, sor)
            w["B1"], w["C1"] = f"Somme de VOLFACT {an1}", f"Somme de VOLFACT {an}"
            w["J2"], w["J3"] = an1, an
            style = w["D2"]._style
            for k, rf in enumerate(sorted(set(v25) | set(v26)), start=2):
                w.cell(k, 1, rf).number_format = "@"
                w.cell(k, 2, v25[rf] if rf in v25 else None)
                w.cell(k, 3, v26[rf] if rf in v26 else None)
                c = w.cell(k, 4, f'=IF(C{k}="","sortant",IF(B{k}="","entrant",IF(AND(B{k}<>0,C{k}<>0),C{k}-B{k},"rien")))')
                c._style = style
            w["K5"] = "=K4-G4"                      # variation hors effet entrants / sortants
            w["L5"] = "=(((K3-G4)-K2)/K2)"

    # --- CONTROLE CA ---
    ws = wb["CONTROLE CA"]
    ws["R2"] = d["lot"]
    ws["F4"] = ", ".join(d["communes"])
    ws["K28"] = d["date_facturation"]
    ws["M28"] = d["nb_factures"]
    ws["O28"] = d["train"]
    ws["Q28"] = d["population"]
    if d["nb_mensus"]:
        ws["J29"], ws["J30"] = None, "X"
    ws["M30"] = d["nb_mensus"]
    ws["F34"], ws["H34"] = d["abo"]
    ws["F35"], ws["H35"] = d["conso"]

    p = d.get("prec")
    if p:
        ws["K13"] = p.get("K28")
        ws["M13"] = p.get("M28")
        ws["O13"] = p.get("M30") or 0
        ws["Q13"] = p.get("O28")
        ws["P14"] = p.get("Q28")
        ws["F21"], ws["H21"] = p.get("F34"), p.get("H34")
        ws["F22"], ws["H22"] = p.get("F35"), p.get("H35")
    if n1:
        for ligne, codes in LIGNES.items():
            ws[f"L{ligne}"] = round(somme(n1, codes, 13), 2)
            ws[f"O{ligne}"] = jours_abo(n1) if ligne == 60 else somme(n1, codes, 11)
        nb_n1 = p.get("M28") if p else len({r[0] for r in n1})
        for cell, txt in commentaires(n, n1, an, d["nb_factures"], nb_n1, es).items():
            ws[cell] = txt

    wb.calculation.fullCalcOnLoad = True     # Excel recalcule tout à l'ouverture
    wb.save(out)
    return alertes
