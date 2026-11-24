"""Fichier d'essai dans la présentation de l'essai 2025 (onglets 04.04, POUR MEDIANE, NON FACT…)."""
import collections, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter
from lecture import ref, to_date

POLICE = "Aptos Narrow"
_t = Side(style="thin")
BORD = Border(left=_t, right=_t, top=_t, bottom=_t)
DATE = "mm-dd-yy"


def remplissage(rgb=None, theme=None, tint=0.0, indexed=None):
    if indexed is not None:
        c = Color(indexed=indexed)
    elif theme is not None:
        c = Color(theme=theme, tint=tint)
    else:
        c = Color(rgb=rgb)
    return PatternFill(fill_type="solid", fgColor=c, bgColor=Color(indexed=64))


VERT = remplissage("FF92D050")
JAUNE = remplissage("FFFFFF00")
GRIS = remplissage(theme=2, tint=-0.0999786370433668)
F11 = Font(name=POLICE, sz=11)
F10 = Font(name=POLICE, sz=10)


def propre(v):
    return v.lstrip("'").strip() if isinstance(v, str) else v


def pds0(v):
    try:
        return int(str(v).lstrip("'").strip())
    except (TypeError, ValueError):
        return v


def copier_table(wb, nom, t, filtre=True, largeurs=None):
    """Recopie un export tel quel (valeurs), police de l'essai 2025, filtre sur l'en-tête."""
    ws = wb.create_sheet(nom)
    for j, h in enumerate(t.headers, 1):
        ws.cell(1, j, h).font = F11
    for i, r in enumerate(t.rows, 2):
        for j, v in enumerate(r, 1):
            if v is None:
                continue
            c = ws.cell(i, j, v)
            c.font = F11
            if isinstance(v, datetime.datetime):
                c.number_format = DATE
    if filtre:
        ws.auto_filter.ref = f"A1:{get_column_letter(max(1, len(t.headers)))}{len(t.rows) + 1}"
    for k, w in (largeurs or {}).items():
        ws.column_dimensions[k].width = w
    return ws


def construire(out, d, log):
    T = d["tables"]
    listing, ctrl = T.get("listing"), T.get("ctrl")
    lot, train = d["lot"], d["train"]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # index du listing par référence d'abonnement
    lst = {}
    if listing:
        for r in listing.rows:
            lst[ref(listing.get(r, "REF_ABONNEMENT", "REF_PDS"))] = r
    def L(r, *names):
        return listing.get(r, *names) if (listing and r) else None

    # ---------- TCD LISTING (offre > commune, par mode de paiement, sans les mensualisés) ----------
    ws = wb.create_sheet("TCD LISTING")
    ws["A1"], ws["B1"] = "ETAT_ABONNEMENT", "Actif"
    ws["A2"], ws["B2"] = "GROUPE_FACTURATION", d["groupe"]
    ws["A4"], ws["B4"] = "Nombre de REF_PDI", "Étiquettes de colonnes"
    if listing:
        compte = collections.defaultdict(lambda: collections.Counter())
        modes = set()
        for r in listing.rows:
            mode = L(r, "MODE_PAIEMENT")
            if mode == "Mensualisation" or str(L(r, "ETAT_ABONNEMENT", "ETAT_PDS") or "Actif") != "Actif":
                continue
            offre = L(r, "TYPE_OFFRE") or "(vide)"
            ville = L(r, "VILLE_PDI", "VILLE_PDL") or "(vide)"
            compte[(offre, ville)][mode] += 1
            modes.add(mode)
        modes = sorted(modes)
        ent = ["Étiquettes de lignes"] + modes + ["Total général"]
        for j, h in enumerate(ent, 1):
            ws.cell(5, j, h).font = Font(name=POLICE, sz=11, b=True)
        row = 6
        tot = collections.Counter()
        for offre in sorted({o for o, _ in compte}):
            sous = collections.Counter()
            for (o, v), c in compte.items():
                if o == offre:
                    sous.update(c)
            vals = [offre] + [sous[m] or None for m in modes] + [sum(sous.values())]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row, j, v); c.font = Font(name=POLICE, sz=11, b=True)
                c.fill = remplissage(theme=4, tint=0.7999816888943144)
            row += 1
            for ville in sorted(v for o, v in compte if o == offre):
                c = compte[(offre, ville)]
                vals = [ville] + [c[m] or None for m in modes] + [sum(c.values())]
                for j, v in enumerate(vals, 1):
                    cell = ws.cell(row, j, v); cell.font = F11
                    if j == 1:
                        cell.alignment = Alignment(indent=1)
                row += 1
            tot.update(sous)
        vals = ["Total général"] + [tot[m] for m in modes] + [sum(tot.values())]
        for j, v in enumerate(vals, 1):
            ws.cell(row, j, v).font = Font(name=POLICE, sz=11, b=True)
        d["population_listing"] = sum(tot.values())
    for k, w in {"A": 28.9, "B": 24.4, "C": 12.6, "D": 12.6, "E": 12.6}.items():
        ws.column_dimensions[k].width = w

    # ---------- ANO AVANT CALCUL ----------
    if T.get("anomalies"):
        a = T["anomalies"]
        ws = copier_table(wb, "ANO AVANT CALCUL", a, filtre=False,
                          largeurs={"A": 17.0, "F": 55.0, "I": 30.0})
        nc = len(a.headers)
        for j, h in ((nc + 1, "A TRAITER"), (nc + 2, "JUSTIFICATION")):
            c = ws.cell(1, j, h)
            c.font = Font(name="Calibri", sz=11, b=True); c.fill = remplissage(indexed=22); c.border = BORD
        ws.column_dimensions[get_column_letter(nc + 1)].width = 12.1
        ws.column_dimensions[get_column_letter(nc + 2)].width = 37.0
        ws.auto_filter.ref = f"A1:{get_column_letter(nc + 2)}{len(a.rows) + 1}"

    # ---------- NON FACT APRES CALCUL / NON FACT FINAL ----------
    NF = ["PDS", "PDS 0", "Titulaire", "Details", "Commentaire agent", "Début PDS", "Fin PDS",
          "Groupe de facturation", "Type d'offre", "Ville PDL"]
    NFW = [10.0, 13.7, 27.4, 11.9, 16.7, 10.1, 9.0, 18.1, 17.3, 22.9, 37.6]

    def non_fact(nom, t, refs_garder, dernier, marquer):
        w = wb.create_sheet(nom)
        for j, h in enumerate(NF + [dernier], 1):
            c = w.cell(1, j, h)
            c.font = Font(name=POLICE, sz=10, b=True)
            c.alignment = Alignment(horizontal="center", wrap_text=True)
            w.column_dimensions[get_column_letter(j)].width = NFW[j - 1]
        w.row_dimensions[1].height = 27.6
        k = 1
        for r in t.rows:
            rf = ref(t.get(r, "Référence"))
            if refs_garder is not None and rf not in refs_garder:
                continue
            k += 1
            lr = lst.get(rf)
            ok = marquer and rf in d["refs_facturees"]
            vals = [rf, pds0(rf), propre(t.get(r, "Nom abonné")), "Non facturé(s)", None,
                    to_date(t.get(r, "Démarré le")), to_date(t.get(r, "Terminé le")),
                    L(lr, "GROUPE_FACTURATION") or d["groupe"], L(lr, "TYPE_OFFRE"),
                    L(lr, "VILLE_PDI", "VILLE_PDL"), "facture ok" if ok else None]
            for j, v in enumerate(vals, 1):
                c = w.cell(k, j, v)
                c.font = F11 if j in (5, 11) else F10
                if j < 11:
                    c.alignment = Alignment(wrap_text=True)
                if isinstance(v, datetime.datetime):
                    c.number_format = DATE
                if ok:
                    c.fill = VERT
        w.auto_filter.ref = f"A1:K{max(k, 2)}"
        return k - 1

    if T.get("sans_facture"):
        non_fact("NON FACT APRES CALCUL", T["sans_facture"], None, "COMMENTAIRE", True)

    # ---------- POUR MEDIANE + 04.04 ----------
    if ctrl:
        c_rows = [r for r in ctrl.rows if ctrl.get(r, "Ref. PDS") is not None]
        g = lambda r, n: ctrl.get(r, n)
        LAB = ["Période Abo EP. Debut", "Période Abo EP. Fin", "Période Conso EP. Debut",
               "Période Conso EP. Fin", "Période Conso Asst. Debut", "Période Conso Asst. Fin"]

        def mediane(nom, offre, labels):
            w = wb.create_sheet(nom)
            data = [r for r in c_rows if g(r, "Offre Type") == offre]
            fin = 2 + max(len(data), 1)
            for j, lab in enumerate(labels, 1):
                cl = get_column_letter(j)
                c = w.cell(1, j, f"=MEDIAN({cl}3:{cl}{fin})"); c.number_format = DATE
                c = w.cell(2, j, lab); c.border = BORD; c.alignment = Alignment(wrap_text=True); c.font = F11
                w.column_dimensions[cl].width = 13
            w.row_dimensions[2].height = 43.2
            for k, r in enumerate(data, 3):
                for j, lab in enumerate(labels, 1):
                    c = w.cell(k, j, to_date(g(r, lab))); c.border = BORD; c.number_format = DATE; c.font = F11

        offres = collections.Counter(g(r, "Offre Type") for r in c_rows)
        o_ea = next((o for o in offres if o and "ASST" in str(o).upper()), None)
        o_e = next((o for o in offres if o and "ASST" not in str(o).upper()), None)
        mediane("POUR MEDIANE OFFRE EAU +ASST", o_ea, LAB)
        mediane("POUR MEDIANE OFFRE EAU", o_e, LAB[:4])

        ws = wb.create_sheet("04.04")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "DETAIL DES CONTROLES FACTURES"
        ws["A1"].font = Font(name=POLICE, sz=18, b=True)
        ws.row_dimensions[1].height = 23.4
        if d.get("fin_traitement"):
            ws["A2"] = f"Fin de traitement à {d['fin_traitement']:%d/%m/%Y %H:%M}"
        ws["A3"] = f"Train:{train} EPN LOT {lot}"
        heads = list(ctrl.headers)
        i_ano = ctrl.col("Nb Ano")
        ano_cols = list(range(i_ano + 1, len(heads))) if i_ano is not None else []
        ent = ["Ref. PDL", "Ref. PDS", "PDS 0", "Code Facture", "A SORTIR DU TRAIN", "Mode Facturation",
               "Offre Type", "Client", "Adresse", "Ville", "HT", "TVA", "TTC", "DELTA DATE DEBUT",
               "DELTA DATE FIN", "commentaire Abo", "Période Abo EP. Debut", "Période Abo EP. Fin",
               "commentaire date conso", "Période Conso EP. Debut", "Période Conso EP. Fin",
               "Période Conso Asst. Debut", "Période Conso Asst. Fin", "commentaire Conso",
               "Volume_eau_fact", "Volume_asst_fact", "DELTA VOLUME", "surconso",
               "METHODE DERNIERE RELEVE ", "CONSO DERNIERE RELEVE", "METHODE RELEVE 1",
               "CONSO RELEVE 1", "CONSO RELEVE 2", "Debut PDS", "Fin PDS", "Nb Ano"]
        ent += [heads[i] for i in ano_cols]
        gris = set(ent[13:15] + ent[26:33])
        jaunes = {"commentaire Abo", "commentaire date conso", "commentaire Conso"}
        for j, h in enumerate(ent, 1):
            c = ws.cell(6, j, h); c.font = F11; c.border = BORD; c.alignment = Alignment(wrap_text=True)
            if h == "A SORTIR DU TRAIN":
                c.font = Font(name=POLICE, sz=11, color="FFFF0000"); c.fill = JAUNE
            elif h in gris:
                c.fill = GRIS
            elif h in jaunes:
                c.fill = JAUNE
        ws.row_dimensions[6].height = 28.8
        first = 7
        for k, r in enumerate(c_rows):
            n = first + k
            eau_seule = "ASST" not in str(g(r, "Offre Type")).upper()
            rf = ref(g(r, "Ref. PDS"))
            lr = lst.get(rf)
            vals = [propre(g(r, "Ref. PDL")), rf, pds0(rf), propre(g(r, "Code Facture")), None,
                    g(r, "Mode Facturation"), g(r, "Offre Type"), propre(g(r, "Client")),
                    propre(g(r, "Adresse")), g(r, "Ville"), g(r, "HT"), g(r, "TVA"), g(r, "TTC"),
                    f"=Q{n}-T{n}" if eau_seule else f'=IF(AND(Q{n}=T{n},Q{n}=V{n},T{n}=V{n}),Q{n}-T{n},"problème date début")',
                    f"=R{n}-U{n}" if eau_seule else f'=IF(AND(R{n}=U{n},R{n}=W{n},U{n}=W{n}),R{n}-U{n},"problème date fin")',
                    None, to_date(g(r, LAB[0])), to_date(g(r, LAB[1])), None,
                    to_date(g(r, LAB[2])), to_date(g(r, LAB[3])), to_date(g(r, LAB[4])), to_date(g(r, LAB[5])), None,
                    ctrl.get(r, "Volume_eau fact", "Volume_eau_fact"), ctrl.get(r, "Volume_asst fact", "Volume_asst_fact"),
                    "offre eau seule" if eau_seule else f"=Y{n}-Z{n}", None,
                    L(lr, "METHODE_RELEVE"), L(lr, "CONSO_DERNIER_RELEVE"), L(lr, "METHODE_RELEVE_1"),
                    L(lr, "CONSO_RELEVE_1"), L(lr, "CONSO_RELEVE_2"),
                    to_date(g(r, "Debut PDS")), to_date(g(r, "Fin PDS")), g(r, "Nb Ano")]
            vals += [r[i] if i < len(r) else None for i in ano_cols]
            for j, v in enumerate(vals, 1):
                c = ws.cell(n, j, v); c.border = BORD
                if isinstance(v, datetime.datetime):
                    c.number_format = DATE
        last = first + len(c_rows) - 1
        ws.auto_filter.ref = f"A6:{get_column_letter(len(ent))}{last}"
        ws.freeze_panes = "C7"
        blocs = [("MEDIANE OFFRE EAU + ASST", 3, remplissage(theme=7, tint=0.7999816888943144), "'POUR MEDIANE OFFRE EAU +ASST'", "QRTUVW"),
                 ("MEDIANE OFFRE EAU", 4, remplissage(theme=8, tint=0.5999938962981048), "'POUR MEDIANE OFFRE EAU'", "QRTU")]
        for lab, rr, f, sh, cols in blocs:
            c = ws[f"P{rr}"]; c.value = lab; c.fill = f; c.alignment = Alignment(wrap_text=True)
            for k, col in enumerate("QRTUVW"):
                c = ws[f"{col}{rr}"]; c.fill = f; c.number_format = DATE
                if col in cols:
                    c.value = f"={sh}!{get_column_letter(k + 1)}1"
        f = remplissage(theme=9, tint=0.3999755851924192)
        ws["P5"] = "MEDIANE GENERAL"; ws["P5"].fill = f; ws["P5"].alignment = Alignment(wrap_text=True)
        for col in "QRTUVW":
            c = ws[f"{col}5"]; c.value = f"=MEDIAN({col}{first}:{col}{last})"; c.fill = f; c.number_format = DATE
        larg = {"A": 11.0, "B": 15.6, "C": 13.3, "D": 14.3, "E": 24.0, "F": 19.3, "G": 18.9, "H": 45.7,
                "I": 35.0, "J": 39.3, "K": 9.0, "L": 7.0, "M": 11.7, "N": 20.6, "O": 18.4, "P": 25.6,
                "Q": 22.7, "R": 20.3, "S": 25.0, "T": 25.0, "U": 22.4, "V": 26.6, "W": 24.1, "X": 24.1,
                "Y": 18.4, "Z": 18.9, "AA": 18.9, "AB": 18.9, "AC": 18.9, "AH": 11.9, "AI": 7.6, "AJ": 7.4}
        for k, w in larg.items():
            ws.column_dimensions[k].width = w
        for j in range(37, len(ent) + 1):
            ws.column_dimensions[get_column_letter(j)].width = 31.1

    # ---------- LISTING ACTIFS, POPULATION TRAIN, NON FACT FINAL, autres exports ----------
    if listing:
        w = copier_table(wb, f"LISTING ACTIFS {d['date_facturation']:%d-%m-%Y}"[:31], listing)
        i_mode = listing.col("MODE_PAIEMENT")
        if i_mode is not None:
            w.freeze_panes = f"{get_column_letter(i_mode + 2)}2"
    if T.get("population"):
        copier_table(wb, "POPULATION TRAIN", T["population"], largeurs={"A": 11.4, "C": 34, "D": 30})
        manquants = {ref(T["population"].get(r, "Référence")) for r in T["population"].rows} - d["refs_facturees"]
        non_fact("NON FACT FINAL", T["population"], manquants, "justification ", False)
    if T.get("factures"):
        copier_table(wb, "FACTURES TOUS ETATS", T["factures"], largeurs={"E": 15, "K": 38})
    if T.get("m2a"):
        copier_table(wb, "DECHARGEMENT M2A", T["m2a"])

    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
