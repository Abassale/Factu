"""Lecture des fichiers bruts : reconnaît chaque export par ses en-têtes de colonnes."""
import csv, datetime, os, re, statistics, collections
import openpyxl

# Signatures : un onglet est reconnu si sa ligne d'en-tête contient toutes ces colonnes.
SIGNATURES = {
    "jdv": {"FACTURE", "VOLFACT", "MNTMTTC"},
    "listing": {"CODE_SYNDICAT", "MODE_PAIEMENT", "GROUPE_FACTURATION"},
    "ctrl": {"REF. PDL", "REF. PDS", "CODE FACTURE"},
    "factures": {"N° RÔLE", "RÉF. DETTE"},
    "anomalies": {"DATE ANOMALIE", "MESSAGE", "REF. ABN"},
    "abonnements": {"RÉFÉRENCE", "NOM ABONNÉ", "DÉMARRÉ LE"},
    "m2a": {"PDI RÉFÉRENCE", "CODE TOURNÉE"},
}
NOMS = {
    "jdv": "JDV (journal de facturation)",
    "listing": "Listing clients",
    "ctrl": "Contrôle factu (détail des contrôles factures)",
    "factures": "Factures tous états",
    "anomalies": "PDI en anomalie",
    "population": "Population totale",
    "sans_facture": "Abonnements sans facture",
    "m2a": "Déchargement M2A",
    "controle_n1": "Contrôle des tendances de l'an dernier",
}


def norm(h):
    if h is None:
        return ""
    return re.sub(r"\s+", " ", str(h)).strip().upper()


class Table:
    """Un onglet reconnu : en-têtes + lignes de valeurs."""

    def __init__(self, kind, path, sheet, headers, rows):
        self.kind, self.path, self.sheet = kind, path, sheet
        self.headers = headers
        self.rows = rows
        self.idx = {}
        for i, h in enumerate(headers):
            for key in (norm(h), norm(h).replace(" ", "_")):
                self.idx.setdefault(key, i)

    def col(self, *names):
        for n in names:
            for key in (norm(n), norm(n).replace(" ", "_")):
                if key in self.idx:
                    return self.idx[key]
        return None

    def get(self, row, *names):
        i = self.col(*names)
        return row[i] if i is not None and i < len(row) else None

    def __repr__(self):
        return f"{os.path.basename(self.path)} › {self.sheet} ({len(self.rows)} lignes)"


# ---------- conversions ----------
_date_re = re.compile(r"^(\d{2})/(\d{2})/(\d{4})(?: (\d{2}):(\d{2})(?::(\d{2}))?)?$")
_num_re = re.compile(r"^-?\d+([.,]\d+)?$")


def to_date(v):
    if isinstance(v, datetime.datetime):
        return v
    if isinstance(v, datetime.date):
        return datetime.datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        m = _date_re.match(v.strip())
        if m:
            d, mo, y, hh, mm, ss = m.groups()
            return datetime.datetime(int(y), int(mo), int(d), int(hh or 0), int(mm or 0), int(ss or 0))
    return v


def from_csv(v):
    """Convertit une valeur texte de CSV, sans abîmer les références à zéros en tête."""
    if v is None:
        return None
    s = v.strip()
    if s == "":
        return None
    d = to_date(s)
    if isinstance(d, datetime.datetime):
        return d
    if _num_re.match(s) and not (s.lstrip("-").startswith("0") and len(s.lstrip("-")) > 1 and s[1:2] not in ".,"):
        s2 = s.replace(",", ".")
        return float(s2) if "." in s2 else int(s2)
    return s


def num(x):
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    try:
        return float(str(x).replace(",", ".").replace(" ", "").replace(" ", ""))
    except ValueError:
        return 0.0


def ref(v):
    """Référence d'abonnement / PDS normalisée (texte, sans apostrophe)."""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).lstrip("'").strip()


# ---------- lecture des fichiers ----------
def _find_header(rows_iter, limit=12):
    """Cherche la ligne d'en-tête parmi les premières lignes. Renvoie (kind, n°, en-têtes, lignes lues)."""
    head = []
    for i, row in enumerate(rows_iter):
        head.append(list(row))
        if i >= limit:
            break
    for i, row in enumerate(head):
        cells = {norm(v) for v in row if v is not None}
        for kind, sig in SIGNATURES.items():
            if sig <= cells:
                return kind, i, [v for v in row], head
    return None, None, None, head


def _read_xlsx(path):
    out = []
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "CONTROLE CA" in wb.sheetnames:
        out.append(Table("controle_n1", path, "CONTROLE CA", [], []))
        return out, wb
    for ws in wb.worksheets:
        it = ws.iter_rows(values_only=True)
        kind, hi, headers, head = _find_header(it)
        if not kind:
            continue
        rows = [r for r in head[hi + 1:]]
        rows += [list(r) for r in it]
        rows = [r for r in rows if r and any(v is not None and v != "" for v in r)]
        rows = [[to_date(v) for v in r] for r in rows]
        out.append(Table(kind, path, ws.title, headers, rows))
    return out, wb


def _read_csv(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    premiere = text.split("\n", 1)[0]
    sep = max(";,\t", key=premiere.count)
    rows = list(csv.reader(text.splitlines(), delimiter=sep))
    kind, hi, headers, _ = _find_header(iter(rows))
    if not kind:
        return []
    data = [[from_csv(v) for v in r] for r in rows[hi + 1:] if any(x.strip() for x in r)]
    return [Table(kind, path, os.path.basename(path), headers, data)]


def scanner(dossier, log):
    """Lit tous les fichiers du dossier d'entrée et classe les onglets reconnus."""
    trouves = collections.defaultdict(list)
    for nom in sorted(os.listdir(dossier)):
        p = os.path.join(dossier, nom)
        if nom.startswith(("~$", ".")) or not os.path.isfile(p):
            continue
        ext = nom.lower().rsplit(".", 1)[-1]
        try:
            if ext in ("xlsx", "xlsm"):
                tables, _ = _read_xlsx(p)
            elif ext in ("csv", "txt"):
                tables = _read_csv(p)
            else:
                log(f"  - ignoré (format non lu) : {nom}")
                continue
        except Exception as e:  # fichier ouvert dans Excel, corrompu…
            log(f"  ! impossible de lire {nom} : {e}")
            continue
        if not tables:
            log(f"  - aucun export reconnu dans : {nom}")
        for t in tables:
            trouves[t.kind].append(t)
    return trouves


def choisir_jdv(tables):
    """Plusieurs onglets JDV : on garde celui nommé « FINAL », sinon le plus récent / le plus gros."""
    finals = [t for t in tables if "FINAL" in norm(t.sheet)]
    if finals:
        return max(finals, key=lambda t: len(t.rows))
    return max(tables, key=lambda t: (annee_jdv(t), len(t.rows)))


def annee_jdv(t):
    fins = [to_date(t.get(r, "FIN")) for r in t.rows]
    fins = [d for d in fins if isinstance(d, datetime.datetime)]
    if not fins:
        return 0
    return sorted(fins)[len(fins) // 2].year


def lignes_jdv(t):
    """Lignes du JDV au format du fichier de contrôle (colonnes A à P)."""
    out = []
    for r in t.rows:
        fact = t.get(r, "FACTURE")
        if fact is None:
            continue
        out.append([
            fact, t.get(r, "CONTRAT"), ref(t.get(r, "PDS", "REF_ABONNEMENT", "REF ABONNEMENT")),
            t.get(r, "REFAQUA"), t.get(r, "TITULAIRE"),
            t.get(r, "CODE_ARTICLE", "CODE ARTICLE"), t.get(r, "NOM_ARTICLE", "NOM ARTICLE"),
            to_date(t.get(r, "DÉBUT", "DEBUT")), to_date(t.get(r, "FIN")),
            t.get(r, "PUHT", "PU HT", "PU_HT"), t.get(r, "VOLBASE"), t.get(r, "VOLFACT"),
            t.get(r, "MNTHT", "MNT HT", "MNT_HT"), t.get(r, "MNTMTTC"), t.get(r, "MNTMTVA"),
            t.get(r, "PÉRIODE", "PERIODE"),
        ])
    return out


def mediane_dates(vals):
    vals = sorted(v for v in vals if isinstance(v, datetime.datetime))
    if not vals:
        return None
    n = len(vals)
    if n % 2:
        return vals[n // 2]
    a, b = vals[n // 2 - 1], vals[n // 2]
    return a + (b - a) / 2
