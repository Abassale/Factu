"""Outil Factu : à partir des exports bruts d'un lot, produit l'essai mis en forme
et le contrôle des tendances rempli (comparé à la facturation de l'an dernier).

Usage : python factu.py [dossier_entree] [--lot N]
"""
import argparse, collections, datetime, glob, os, re, shutil, sys, time
import openpyxl

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
from lecture import scanner, choisir_jdv, annee_jdv, lignes_jdv, mediane_dates, ref, to_date, NOMS, norm
import controle, essai

RACINE = os.path.dirname(ICI)
MODELE = os.path.join(RACINE, "modeles", "CONTROLE_TENDANCES_VIERGE.xlsx")


def main():
    ap = argparse.ArgumentParser(description="Remplit l'essai et le contrôle des tendances d'un lot.")
    ap.add_argument("entree", nargs="?", default=os.path.join(RACINE, "entree"))
    ap.add_argument("--lot", type=int, help="numéro du lot (sinon lu dans le listing)")
    ap.add_argument("--sortie", default=os.path.join(RACINE, "sortie"))
    ap.add_argument("--historique", default=os.path.join(RACINE, "historique"))
    a = ap.parse_args()

    journal = []
    def log(msg=""):
        print(msg)
        journal.append(msg)

    t0 = time.time()
    log(f"Lecture du dossier : {a.entree}")
    trouves = scanner(a.entree, log)

    # --- population / abonnements sans facture : même export, on les distingue ---
    abos = trouves.pop("abonnements", [])
    for t in abos:
        nom = norm(os.path.basename(t.path) + " " + t.sheet)
        if "SANS" in nom:
            trouves["sans_facture"].append(t)
        elif "POPUL" in nom:
            trouves["population"].append(t)
    reste = [t for t in abos if t not in trouves["sans_facture"] and t not in trouves["population"]]
    for t in sorted(reste, key=lambda t: -len(t.rows)):
        (trouves["population"] if not trouves["population"] else trouves["sans_facture"]).append(t)

    # --- JDV de l'année et éventuel JDV de l'an dernier ---
    if not trouves.get("jdv"):
        log("\n!! Aucun JDV trouvé : impossible de continuer.")
        return fin(journal, a.sortie, None)
    par_annee = collections.defaultdict(list)
    for t in trouves["jdv"]:
        par_annee[annee_jdv(t)].append(t)
    an_jdv = max(par_annee)
    jdv_t = choisir_jdv(par_annee[an_jdv])
    jdv_n1_t = choisir_jdv(par_annee[an_jdv - 1]) if (an_jdv - 1) in par_annee else None

    T = {k: (v[0] if k != "jdv" else None) for k, v in trouves.items() if v}
    T["jdv"] = jdv_t
    log("\nFichiers reconnus :")
    for k in ("jdv", "listing", "ctrl", "factures", "anomalies", "population", "sans_facture", "m2a"):
        log(f"  {NOMS[k]:<48} {T[k] if T.get(k) else '— absent'}")
    if jdv_n1_t:
        log(f"  {'JDV de l’an dernier':<48} {jdv_n1_t}")

    n = lignes_jdv(jdv_t)
    refs_fact = {r[2] for r in n}
    listing = T.get("listing")

    # --- lot, année, semestre ---
    lot = a.lot
    groupe = None
    if listing:
        gc = collections.Counter(listing.get(r, "GROUPE_FACTURATION") for r in listing.rows)
        groupe = gc.most_common(1)[0][0]
        m = re.search(r"LOT\s*0*(\d+)", str(groupe).upper())
        if m and not lot:
            lot = int(m.group(1))
    if not lot:
        for t in trouves.get("jdv", []) + [x for v in trouves.values() for x in v]:
            m = re.search(r"LOT[ _-]*0*(\d+)", norm(os.path.basename(t.path)))
            if m:
                lot = int(m.group(1)); break
    if not lot:
        log("\n!! Numéro de lot introuvable : relance avec --lot N (ex. lancer.bat 12).")
        return fin(journal, a.sortie, None)
    groupe = groupe or f"EPN Masse Lot {lot}"

    fac = T.get("factures")
    date_fact, train, fin_trait = None, None, None
    if fac:
        emis = [to_date(fac.get(r, "Date émission")) for r in fac.rows]
        calc = [to_date(fac.get(r, "Date calcul")) for r in fac.rows]
        emis = [x for x in emis if isinstance(x, datetime.datetime)]
        calc = [x for x in calc if isinstance(x, datetime.datetime)]
        base = emis or calc
        if base:
            dmin = min(base)
            date_fact = datetime.datetime(dmin.year, dmin.month, dmin.day)
        if calc:
            fin_trait = max(calc)
        tr = collections.Counter((fac.get(r, "Exercice"), fac.get(r, "N° rôle")) for r in fac.rows)
        (ex, role), _ = tr.most_common(1)[0]
        train = f"{ex}-{role}"
    if not date_fact:
        fins = sorted(to_date(r[8]) for r in n if isinstance(to_date(r[8]), datetime.datetime))
        date_fact = fins[len(fins) // 2]
        log(f"\n! Pas d'export « Factures tous états » : date de facturation estimée au {date_fact:%d/%m/%Y}, à corriger.")
    annee = date_fact.year
    sem = 2 if date_fact.month >= 7 else 1
    train = train or "x"

    # --- population ---
    pop_t = T.get("population")
    nb_fact = len({r[0] for r in n})
    if pop_t:
        population = len({ref(pop_t.get(r, "Référence")) for r in pop_t.rows})
    else:
        population = nb_fact + (len(T["sans_facture"].rows) if T.get("sans_facture") else 0)
        log("! Pas d'export « Population totale » : population estimée (factures + abonnements sans facture).")

    lst = {}
    if listing:
        for r in listing.rows:
            lst[ref(listing.get(r, "REF_ABONNEMENT", "REF_PDS"))] = r
    nb_mensus = sum(1 for x in refs_fact if x in lst and listing.get(lst[x], "MODE_PAIEMENT") == "Mensualisation")
    communes = sorted({listing.get(lst[x], "VILLE_PDI", "VILLE_PDL") for x in refs_fact if x in lst} - {None})

    # --- périodes (médianes) ---
    ctrl = T.get("ctrl")
    def med(table, col):
        return mediane_dates([to_date(table.get(r, col)) for r in table.rows])
    if ctrl:
        abo = (med(ctrl, "Période Abo EP. Debut"), med(ctrl, "Période Abo EP. Fin"))
        conso = (med(ctrl, "Période Conso EP. Debut"), med(ctrl, "Période Conso EP. Fin"))
    else:
        def medj(codes, k):
            return mediane_dates([to_date(r[k]) for r in n if r[5] in codes])
        abo = (medj({"E50EPN", "E50E"}, 7), medj({"E50EPN", "E50E"}, 8))
        conso = (medj({"E60EPN", "E60C"}, 7), medj({"E60EPN", "E60C"}, 8))
        log("! Pas d'export « Contrôle factu » : périodes calculées sur le JDV.")

    # --- facturation de l'an dernier : contrôle N-1 (entrée ou historique) ---
    cand = [t.path for t in trouves.get("controle_n1", [])]
    hist_lot = os.path.join(a.historique, f"LOT_{lot:02d}")
    cand += glob.glob(os.path.join(hist_lot, "*.xlsx"))
    prec, jdv_n1, prec_path = None, None, None
    meilleurs = []
    for p in dict.fromkeys(cand):
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            ws = wb["CONTROLE CA"]
            k28, r2 = ws["K28"].value, ws["R2"].value
            wb.close()
        except Exception:
            continue
        k28 = to_date(k28)
        if not isinstance(k28, datetime.datetime) or k28 >= date_fact:
            continue
        if r2 not in (None, "") and str(r2).strip().lstrip("0") != str(lot):
            continue
        meme = (k28.year == annee - 1) and ((k28.month >= 7) == (sem == 2))
        meilleurs.append((meme, k28, p))
    if meilleurs:
        meme, k28, prec_path = max(meilleurs)
        prec, jdv_n1 = controle.lire_controle_n1(prec_path)
        log(f"\nFacturation précédente : {os.path.basename(prec_path)} (facturée le {k28:%d/%m/%Y})")
        if not meme:
            log("! Ce n'est pas la même période de l'an dernier : vérifie la comparaison.")
    elif jdv_n1_t:
        jdv_n1 = lignes_jdv(jdv_n1_t)
        log("\n! Pas de contrôle de l'an dernier : comparaison faite avec le JDV N-1 fourni, "
            "mais la date, le nombre de factures et la population N-1 sont à saisir à la main.")
    else:
        log(f"\n! Aucun contrôle de l'an dernier pour le lot {lot} (ni dans l'entrée ni dans {hist_lot}) : "
            "la partie N-1 restera vide.")

    d = dict(tables=T, jdv_n=n, jdv_n1=jdv_n1, prec=prec, lot=lot, groupe=groupe, annee=annee,
             date_facturation=date_fact, train=train, nb_factures=nb_fact, population=population,
             nb_mensus=nb_mensus, communes=communes, abo=abo, conso=conso,
             refs_facturees=refs_fact, fin_traitement=fin_trait)

    dossier = os.path.join(a.sortie, f"LOT_{lot:02d}_{annee}_S{sem}")
    os.makedirs(dossier, exist_ok=True)
    f_essai = os.path.join(dossier, f"ESSAI_EPN_LOT_{lot:02d}_{annee}.xlsx")
    f_ctrl = os.path.join(dossier, f"EVREUX_LOT_{lot:02d}_ControlesTendances_S{sem}_{annee}.xlsx")

    log(f"\nLot {lot} — facturation du {date_fact:%d/%m/%Y} (S{sem} {annee}), train {train}")
    log(f"  {nb_fact} factures, population {population}, {population - nb_fact} sans facture, {nb_mensus} mensualisé(s)")
    log(f"  période abonnement {abo[0]:%d/%m/%Y} → {abo[1]:%d/%m/%Y}, consommation {conso[0]:%d/%m/%Y} → {conso[1]:%d/%m/%Y}")

    log("\nCréation de l'essai…")
    essai.construire(f_essai, d, log)
    if d.get("population_listing") and d["population_listing"] != population:
        log(f"! Le listing compte {d['population_listing']} abonnés non mensualisés, la population totale {population} : à vérifier.")
    log("Création du contrôle des tendances (1 à 3 minutes)…")
    alertes = controle.remplir(MODELE, f_ctrl, d, log)
    for x in alertes:
        log(f"! {x}")

    os.makedirs(hist_lot, exist_ok=True)
    shutil.copy2(f_ctrl, os.path.join(hist_lot, os.path.basename(f_ctrl)))
    log(f"\nTerminé en {time.time() - t0:.0f} s. Fichiers dans : {dossier}")
    log("À faire à la main : relire les commentaires « Validation des évolutions », "
        "remplir les colonnes commentaire / justification de l'essai.")
    return fin(journal, a.sortie, dossier)


def fin(journal, sortie, dossier):
    os.makedirs(dossier or sortie, exist_ok=True)
    with open(os.path.join(dossier or sortie, "rapport.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(journal) + "\n")
    return 0 if dossier else 1


if __name__ == "__main__":
    sys.exit(main())
