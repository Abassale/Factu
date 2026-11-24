/* Factu — cœur de l'outil (navigateur et Node).
 * À partir des exports bruts d'un lot : essai mis en forme (présentation 2025)
 * + contrôle des tendances rempli et comparé à la facturation de l'an dernier.
 */
(function (root) {
  "use strict";
  const Factu = {};

  // ------------------------------------------------------------------ outils
  const SIGNATURES = {
    jdv: ["FACTURE", "VOLFACT", "MNTMTTC"],
    listing: ["CODE_SYNDICAT", "MODE_PAIEMENT", "GROUPE_FACTURATION"],
    ctrl: ["REF. PDL", "REF. PDS", "CODE FACTURE"],
    factures: ["N° RÔLE", "RÉF. DETTE"],
    anomalies: ["DATE ANOMALIE", "MESSAGE", "REF. ABN"],
    abonnements: ["RÉFÉRENCE", "NOM ABONNÉ", "DÉMARRÉ LE"],
    m2a: ["PDI RÉFÉRENCE", "CODE TOURNÉE"],
  };
  const NOMS = {
    jdv: "JDV (journal de facturation)",
    listing: "Listing clients",
    ctrl: "Contrôle factu",
    factures: "Factures tous états",
    anomalies: "PDI en anomalie",
    population: "Population totale",
    sans_facture: "Abonnements sans facture",
    m2a: "Déchargement M2A",
  };
  Factu.NOMS = NOMS;

  const norm = (h) => (h == null ? "" : String(h).replace(/\s+/g, " ").trim().toUpperCase());
  const isDate = (v) => v instanceof Date && !isNaN(v);
  const DATE_RE = /^(\d{2})\/(\d{2})\/(\d{4})(?: (\d{2}):(\d{2})(?::(\d{2}))?)?$/;
  function toDate(v) {
    if (isDate(v)) return v;
    if (typeof v === "string") {
      const m = DATE_RE.exec(v.trim());
      if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)));
    }
    return v;
  }
  function num(x) {
    if (x == null || x === "") return 0;
    if (typeof x === "number") return x;
    const n = parseFloat(String(x).replace(/[ \s]/g, "").replace(",", "."));
    return isNaN(n) ? 0 : n;
  }
  function ref(v) {
    if (v == null) return null;
    return String(v).replace(/^'+/, "").trim();
  }
  const propre = (v) => (typeof v === "string" ? v.replace(/^'+/, "").trim() : v);
  function pds0(v) {
    const s = String(v == null ? "" : v).replace(/^'+/, "").trim();
    return /^\d+$/.test(s) ? Number(s) : v;
  }
  const dmy = (d) => `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()}`;
  const hm = (d) => `${String(d.getUTCHours()).padStart(2, "0")}:${String(d.getUTCMinutes()).padStart(2, "0")}`;
  const jour = (d) => new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  function mediane(vals) {
    const v = vals.filter(isDate).map((d) => d.getTime()).sort((a, b) => a - b);
    if (!v.length) return null;
    const n = v.length;
    return new Date(n % 2 ? v[(n - 1) / 2] : (v[n / 2 - 1] + v[n / 2]) / 2);
  }
  function counter() { return new Map(); }
  function inc(m, k, n = 1) { m.set(k, (m.get(k) || 0) + n); }
  function mostCommon(m) { let best = null, bv = -Infinity; for (const [k, v] of m) if (v > bv) { best = k; bv = v; } return best; }

  // valeur d'une cellule ExcelJS -> valeur simple
  function val(v) {
    if (v == null) return null;
    if (isDate(v)) return v;
    if (typeof v === "object") {
      if ("result" in v) return val(v.result);
      if (v.richText) return v.richText.map((t) => t.text).join("");
      if ("text" in v) return v.text;
      if (v.error || v.formula || v.sharedFormula) return null;
    }
    return v;
  }

  // ------------------------------------------------------------------ tables
  class Table {
    constructor(kind, name, sheet, headers, rows) {
      this.kind = kind; this.name = name; this.sheet = sheet; this.headers = headers; this.rows = rows;
      this.idx = new Map();
      headers.forEach((h, i) => {
        const k = norm(h);
        if (!this.idx.has(k)) this.idx.set(k, i);
        const k2 = k.replace(/ /g, "_");
        if (!this.idx.has(k2)) this.idx.set(k2, i);
      });
    }
    col(...names) {
      for (const n of names) {
        const k = norm(n);
        if (this.idx.has(k)) return this.idx.get(k);
        if (this.idx.has(k.replace(/ /g, "_"))) return this.idx.get(k.replace(/ /g, "_"));
      }
      return null;
    }
    get(row, ...names) { const i = this.col(...names); return i == null || !row ? null : (row[i] === undefined ? null : row[i]); }
    toString() { return `${this.name} › ${this.sheet} (${this.rows.length} lignes)`; }
  }

  function findHeader(rows) {
    for (let i = 0; i < Math.min(rows.length, 13); i++) {
      const cells = new Set((rows[i] || []).filter((v) => v != null).map(norm));
      for (const [kind, sig] of Object.entries(SIGNATURES)) {
        if (sig.every((s) => cells.has(s))) return { kind, i };
      }
    }
    return null;
  }

  function sheetRows(ws) {
    const out = [];
    ws.eachRow({ includeEmpty: true }, (row, n) => {
      const vals = row.values; // index 1..
      const arr = [];
      for (let j = 1; j < vals.length; j++) arr.push(val(vals[j]));
      out[n - 1] = arr;
    });
    for (let i = 0; i < out.length; i++) if (!out[i]) out[i] = [];
    return out;
  }

  function csvParse(text) {
    const first = text.split("\n", 1)[0];
    const sep = [";", ",", "\t"].sort((a, b) => first.split(b).length - first.split(a).length)[0];
    const rows = []; let row = [], cur = "", q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; }
        else cur += c;
      } else if (c === '"') q = true;
      else if (c === sep) { row.push(cur); cur = ""; }
      else if (c === "\n" || c === "\r") {
        if (c === "\r" && text[i + 1] === "\n") i++;
        row.push(cur); rows.push(row); row = []; cur = "";
      } else cur += c;
    }
    if (cur !== "" || row.length) { row.push(cur); rows.push(row); }
    return rows;
  }
  const NUM_RE = /^-?\d+([.,]\d+)?$/;
  function fromCsv(v) {
    if (v == null) return null;
    const s = v.trim();
    if (s === "") return null;
    const d = toDate(s);
    if (isDate(d)) return d;
    const u = s.replace(/^-/, "");
    if (NUM_RE.test(s) && !(u.length > 1 && u[0] === "0" && !/[.,]/.test(u[1]))) return parseFloat(s.replace(",", "."));
    return s;
  }
  function decode(buf) {
    const bytes = new Uint8Array(buf);
    try { return new TextDecoder("utf-8", { fatal: true }).decode(bytes).replace(/^﻿/, ""); }
    catch (e) { return new TextDecoder("windows-1252").decode(bytes); }
  }

  /** Lit un fichier (xlsx ou csv) et renvoie les exports reconnus. */
  async function lireFichier(ExcelJS, file, log) {
    const name = file.name;
    const ext = name.toLowerCase().split(".").pop();
    const out = [];
    if (ext === "csv" || ext === "txt") {
      const rows = csvParse(decode(file.buffer));
      const h = findHeader(rows);
      if (h) out.push(new Table(h.kind, name, name, rows[h.i], rows.slice(h.i + 1).filter((r) => r.some((x) => x.trim() !== "")).map((r) => r.map(fromCsv))));
      return out;
    }
    if (ext !== "xlsx" && ext !== "xlsm") { log(`  - ignoré (format non lu) : ${name}`); return out; }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(file.buffer);
    const ctrl = wb.getWorksheet("CONTROLE CA");
    if (ctrl) {
      // contrôle des tendances d'une année précédente
      const v = {};
      for (let r = 1; r <= 36; r++) for (let c = 1; c <= 28; c++) {
        const cell = ctrl.getCell(r, c); const x = val(cell.value);
        if (x != null) v[cell.address] = x;
      }
      const jw = wb.getWorksheet("JDV N");
      let jdv = [];
      if (jw) {
        const rows = sheetRows(jw);
        jdv = lignesJdv(new Table("jdv", name, "JDV N", rows[0] || [], rows.slice(1).filter((r) => r[0] != null)));
      }
      out.push({ kind: "controle_n1", name, v, jdv, toString() { return name; } });
      return out;
    }
    for (const ws of wb.worksheets) {
      const rows = sheetRows(ws);
      const h = findHeader(rows);
      if (!h) continue;
      const data = rows.slice(h.i + 1).filter((r) => r.some((x) => x != null && x !== "")).map((r) => r.map(toDate));
      out.push(new Table(h.kind, name, ws.name, rows[h.i], data));
    }
    return out;
  }

  function anneeJdv(t) {
    const f = t.rows.map((r) => toDate(t.get(r, "FIN"))).filter(isDate).map((d) => d.getTime()).sort((a, b) => a - b);
    return f.length ? new Date(f[Math.floor(f.length / 2)]).getUTCFullYear() : 0;
  }
  function choisirJdv(tables) {
    const finals = tables.filter((t) => norm(t.sheet).includes("FINAL"));
    const pool = finals.length ? finals : tables;
    return pool.slice().sort((a, b) => (anneeJdv(b) - anneeJdv(a)) || (b.rows.length - a.rows.length))[0];
  }
  function lignesJdv(t) {
    const out = [];
    for (const r of t.rows) {
      const fact = t.get(r, "FACTURE");
      if (fact == null) continue;
      out.push([
        fact, t.get(r, "CONTRAT"), ref(t.get(r, "PDS", "REF_ABONNEMENT", "REF ABONNEMENT")),
        t.get(r, "REFAQUA"), t.get(r, "TITULAIRE"),
        t.get(r, "CODE_ARTICLE", "CODE ARTICLE"), t.get(r, "NOM_ARTICLE", "NOM ARTICLE"),
        toDate(t.get(r, "DÉBUT", "DEBUT")), toDate(t.get(r, "FIN")),
        t.get(r, "PUHT", "PU HT", "PU_HT"), t.get(r, "VOLBASE"), t.get(r, "VOLFACT"),
        t.get(r, "MNTHT", "MNT HT", "MNT_HT"), t.get(r, "MNTMTTC"), t.get(r, "MNTMTVA"),
        t.get(r, "PÉRIODE", "PERIODE"),
      ]);
    }
    return out;
  }

  // ------------------------------------------------------------------ styles
  const THIN = { style: "thin" };
  const BORD = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  const DATE = "mm-dd-yy";
  const solid = (color) => ({ type: "pattern", pattern: "solid", fgColor: color, bgColor: { indexed: 64 } });
  const VERT = solid({ argb: "FF92D050" });
  const JAUNE = solid({ argb: "FFFFFF00" });
  const GRIS = solid({ theme: 2, tint: -0.0999786370433668 });
  const POLICE = "Aptos Narrow";
  const F11 = { name: POLICE, size: 11 };
  const F10 = { name: POLICE, size: 10 };
  const colL = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

  function set(ws, r, c, v, opts) {
    const cell = ws.getCell(r, c);
    cell.value = v === undefined ? null : v;
    if (opts) {
      if (opts.font) cell.font = opts.font;
      if (opts.fill) cell.fill = opts.fill;
      if (opts.border) cell.border = opts.border;
      if (opts.alignment) cell.alignment = opts.alignment;
      if (opts.numFmt) cell.numFmt = opts.numFmt;
    }
    if (isDate(v) && !(opts && opts.numFmt)) cell.numFmt = DATE;
    return cell;
  }

  // ------------------------------------------------------------------ essai
  function copierTable(wb, nom, t, largeurs) {
    const ws = wb.addWorksheet(nom.slice(0, 31));
    t.headers.forEach((h, j) => set(ws, 1, j + 1, h, { font: F11 }));
    t.rows.forEach((r, i) => r.forEach((v, j) => { if (v != null && v !== "") set(ws, i + 2, j + 1, v, { font: F11 }); }));
    ws.autoFilter = `A1:${colL(Math.max(1, t.headers.length))}${t.rows.length + 1}`;
    for (const [k, w] of Object.entries(largeurs || {})) ws.getColumn(k).width = w;
    return ws;
  }

  function construireEssai(ExcelJS, d) {
    const T = d.tables, listing = T.listing, ctrl = T.ctrl;
    const wb = new ExcelJS.Workbook();
    wb.calcProperties.fullCalcOnLoad = true;
    const lst = new Map();
    if (listing) for (const r of listing.rows) lst.set(ref(listing.get(r, "REF_ABONNEMENT", "REF_PDS")), r);
    const L = (r, ...n) => (listing && r ? listing.get(r, ...n) : null);

    // TCD LISTING
    let ws = wb.addWorksheet("TCD LISTING");
    set(ws, 1, 1, "ETAT_ABONNEMENT"); set(ws, 1, 2, "Actif");
    set(ws, 2, 1, "GROUPE_FACTURATION"); set(ws, 2, 2, d.groupe);
    set(ws, 4, 1, "Nombre de REF_PDI"); set(ws, 4, 2, "Étiquettes de colonnes");
    if (listing) {
      const compte = new Map(); const modes = new Set();
      for (const r of listing.rows) {
        const mode = L(r, "MODE_PAIEMENT");
        const etat = L(r, "ETAT_ABONNEMENT", "ETAT_PDS") || "Actif";
        if (mode === "Mensualisation" || String(etat) !== "Actif") continue;
        const offre = L(r, "TYPE_OFFRE") || "(vide)", ville = L(r, "VILLE_PDI", "VILLE_PDL") || "(vide)";
        const key = offre + "\u0000" + ville;
        if (!compte.has(key)) compte.set(key, counter());
        inc(compte.get(key), mode); modes.add(mode);
      }
      const ms = [...modes].sort();
      ["Étiquettes de lignes", ...ms, "Total général"].forEach((h, j) => set(ws, 5, j + 1, h, { font: { ...F11, bold: true } }));
      let row = 6; const tot = counter();
      const offres = [...new Set([...compte.keys()].map((k) => k.split("\u0000")[0]))].sort();
      for (const offre of offres) {
        const sous = counter();
        const villes = [...compte.keys()].filter((k) => k.split("\u0000")[0] === offre).map((k) => k.split("\u0000")[1]).sort();
        for (const v of villes) for (const [m, n] of compte.get(offre + "\u0000" + v)) inc(sous, m, n);
        const ts = [...sous.values()].reduce((a, b) => a + b, 0);
        [offre, ...ms.map((m) => sous.get(m) || null), ts].forEach((v, j) =>
          set(ws, row, j + 1, v, { font: { ...F11, bold: true }, fill: solid({ theme: 4, tint: 0.7999816888943144 }) }));
        row++;
        for (const v of villes) {
          const c = compte.get(offre + "\u0000" + v);
          const tv = [...c.values()].reduce((a, b) => a + b, 0);
          [v, ...ms.map((m) => c.get(m) || null), tv].forEach((x, j) =>
            set(ws, row, j + 1, x, { font: F11, alignment: j === 0 ? { indent: 1 } : undefined }));
          row++;
        }
        for (const [m, n] of sous) inc(tot, m, n);
      }
      const tt = [...tot.values()].reduce((a, b) => a + b, 0);
      ["Total général", ...ms.map((m) => tot.get(m) || 0), tt].forEach((v, j) => set(ws, row, j + 1, v, { font: { ...F11, bold: true } }));
      d.populationListing = tt;
    }
    for (const [k, w] of Object.entries({ A: 28.9, B: 24.4, C: 12.6, D: 12.6, E: 12.6 })) ws.getColumn(k).width = w;

    // ANO AVANT CALCUL
    if (T.anomalies) {
      const a = T.anomalies;
      ws = copierTable(wb, "ANO AVANT CALCUL", a, { A: 17, F: 55, I: 30 });
      const nc = a.headers.length;
      [["A TRAITER", 12.1], ["JUSTIFICATION", 37]].forEach(([h, w], k) => {
        set(ws, 1, nc + 1 + k, h, { font: { name: "Calibri", size: 11, bold: true }, fill: solid({ indexed: 22 }), border: BORD });
        ws.getColumn(nc + 1 + k).width = w;
      });
      ws.autoFilter = `A1:${colL(nc + 2)}${a.rows.length + 1}`;
    }

    // NON FACT
    const NF = ["PDS", "PDS 0", "Titulaire", "Details", "Commentaire agent", "Début PDS", "Fin PDS",
      "Groupe de facturation", "Type d'offre", "Ville PDL"];
    const NFW = [10.0, 13.7, 27.4, 11.9, 16.7, 10.1, 9.0, 18.1, 17.3, 22.9, 37.6];
    function nonFact(nom, t, garder, dernier, marquer) {
      const w = wb.addWorksheet(nom);
      [...NF, dernier].forEach((h, j) => {
        set(w, 1, j + 1, h, { font: { ...F10, bold: true }, alignment: { horizontal: "center", wrapText: true } });
        w.getColumn(j + 1).width = NFW[j];
      });
      w.getRow(1).height = 27.6;
      let k = 1;
      for (const r of t.rows) {
        const rf = ref(t.get(r, "Référence"));
        if (garder && !garder.has(rf)) continue;
        k++;
        const lr = lst.get(rf);
        const ok = marquer && d.refsFacturees.has(rf);
        const vals = [rf, pds0(rf), propre(t.get(r, "Nom abonné")), "Non facturé(s)", null,
          toDate(t.get(r, "Démarré le")), toDate(t.get(r, "Terminé le")),
          L(lr, "GROUPE_FACTURATION") || d.groupe, L(lr, "TYPE_OFFRE"), L(lr, "VILLE_PDI", "VILLE_PDL"),
          ok ? "facture ok" : null];
        vals.forEach((v, j) => set(w, k, j + 1, v, {
          font: (j === 4 || j === 10) ? F11 : F10,
          alignment: j < 10 ? { wrapText: true } : undefined,
          fill: ok ? VERT : undefined,
        }));
      }
      w.autoFilter = `A1:K${Math.max(k, 2)}`;
      return k - 1;
    }
    if (T.sans_facture) nonFact("NON FACT APRES CALCUL", T.sans_facture, null, "COMMENTAIRE", true);

    // POUR MEDIANE + 04.04
    if (ctrl) {
      const cr = ctrl.rows.filter((r) => ctrl.get(r, "Ref. PDS") != null);
      const g = (r, n) => ctrl.get(r, n);
      const LAB = ["Période Abo EP. Debut", "Période Abo EP. Fin", "Période Conso EP. Debut",
        "Période Conso EP. Fin", "Période Conso Asst. Debut", "Période Conso Asst. Fin"];
      const offres = [...new Set(cr.map((r) => g(r, "Offre Type")))].filter(Boolean);
      const oEA = offres.find((o) => String(o).toUpperCase().includes("ASST"));
      const oE = offres.find((o) => !String(o).toUpperCase().includes("ASST"));
      const med = (nom, offre, labels) => {
        const w = wb.addWorksheet(nom);
        const data = cr.filter((r) => g(r, "Offre Type") === offre);
        const fin = 2 + Math.max(data.length, 1);
        labels.forEach((lab, j) => {
          const cl = colL(j + 1);
          set(w, 1, j + 1, { formula: `MEDIAN(${cl}3:${cl}${fin})` }, { numFmt: DATE });
          set(w, 2, j + 1, lab, { border: BORD, alignment: { wrapText: true }, font: F11 });
          w.getColumn(j + 1).width = 13;
        });
        w.getRow(2).height = 43.2;
        data.forEach((r, k) => labels.forEach((lab, j) => set(w, k + 3, j + 1, toDate(g(r, lab)) || null, { border: BORD, numFmt: DATE, font: F11 })));
      };
      med("POUR MEDIANE OFFRE EAU +ASST", oEA, LAB);
      med("POUR MEDIANE OFFRE EAU", oE, LAB.slice(0, 4));

      ws = wb.addWorksheet("04.04", { views: [{ state: "frozen", xSplit: 2, ySplit: 6, showGridLines: false }] });
      set(ws, 1, 1, "DETAIL DES CONTROLES FACTURES", { font: { name: POLICE, size: 18, bold: true } });
      ws.getRow(1).height = 23.4;
      if (d.finTraitement) set(ws, 2, 1, `Fin de traitement à ${dmy(d.finTraitement)} ${hm(d.finTraitement)}`);
      set(ws, 3, 1, `Train:${d.train} EPN LOT ${d.lot}`);
      const iAno = ctrl.col("Nb Ano");
      const anoCols = [];
      if (iAno != null) for (let i = iAno + 1; i < ctrl.headers.length; i++) if (ctrl.headers[i] != null) anoCols.push(i);
      const ent = ["Ref. PDL", "Ref. PDS", "PDS 0", "Code Facture", "A SORTIR DU TRAIN", "Mode Facturation",
        "Offre Type", "Client", "Adresse", "Ville", "HT", "TVA", "TTC", "DELTA DATE DEBUT",
        "DELTA DATE FIN", "commentaire Abo", ...LAB.slice(0, 2), "commentaire date conso", ...LAB.slice(2),
        "commentaire Conso", "Volume_eau_fact", "Volume_asst_fact", "DELTA VOLUME", "surconso",
        "METHODE DERNIERE RELEVE ", "CONSO DERNIERE RELEVE", "METHODE RELEVE 1",
        "CONSO RELEVE 1", "CONSO RELEVE 2", "Debut PDS", "Fin PDS", "Nb Ano", ...anoCols.map((i) => ctrl.headers[i])];
      const gris = new Set([...ent.slice(13, 15), ...ent.slice(26, 33)]);
      const jaunes = new Set(["commentaire Abo", "commentaire date conso", "commentaire Conso"]);
      ent.forEach((h, j) => {
        const o = { font: F11, border: BORD, alignment: { wrapText: true } };
        if (h === "A SORTIR DU TRAIN") { o.font = { ...F11, color: { argb: "FFFF0000" } }; o.fill = JAUNE; }
        else if (gris.has(h)) o.fill = GRIS;
        else if (jaunes.has(h)) o.fill = JAUNE;
        set(ws, 6, j + 1, h, o);
      });
      ws.getRow(6).height = 28.8;
      const first = 7;
      cr.forEach((r, k) => {
        const n = first + k;
        const eau = !String(g(r, "Offre Type")).toUpperCase().includes("ASST");
        const rf = ref(g(r, "Ref. PDS"));
        const lr = lst.get(rf);
        const f = (s) => ({ formula: s });
        const vals = [propre(g(r, "Ref. PDL")), rf, pds0(rf), propre(g(r, "Code Facture")), null,
          g(r, "Mode Facturation"), g(r, "Offre Type"), propre(g(r, "Client")), propre(g(r, "Adresse")), g(r, "Ville"),
          g(r, "HT"), g(r, "TVA"), g(r, "TTC"),
          eau ? f(`Q${n}-T${n}`) : f(`IF(AND(Q${n}=T${n},Q${n}=V${n},T${n}=V${n}),Q${n}-T${n},"problème date début")`),
          eau ? f(`R${n}-U${n}`) : f(`IF(AND(R${n}=U${n},R${n}=W${n},U${n}=W${n}),R${n}-U${n},"problème date fin")`),
          null, toDate(g(r, LAB[0])), toDate(g(r, LAB[1])), null,
          toDate(g(r, LAB[2])), toDate(g(r, LAB[3])), toDate(g(r, LAB[4])), toDate(g(r, LAB[5])), null,
          ctrl.get(r, "Volume_eau fact", "Volume_eau_fact"), ctrl.get(r, "Volume_asst fact", "Volume_asst_fact"),
          eau ? "offre eau seule" : f(`Y${n}-Z${n}`), null,
          L(lr, "METHODE_RELEVE"), L(lr, "CONSO_DERNIER_RELEVE"), L(lr, "METHODE_RELEVE_1"),
          L(lr, "CONSO_RELEVE_1"), L(lr, "CONSO_RELEVE_2"),
          toDate(g(r, "Debut PDS")), toDate(g(r, "Fin PDS")), g(r, "Nb Ano"),
          ...anoCols.map((i) => (r[i] === undefined ? null : r[i]))];
        vals.forEach((v, j) => set(ws, n, j + 1, v, { border: BORD }));
      });
      const last = first + cr.length - 1;
      ws.autoFilter = `A6:${colL(ent.length)}${last}`;
      const blocs = [
        ["MEDIANE OFFRE EAU + ASST", 3, solid({ theme: 7, tint: 0.7999816888943144 }), "'POUR MEDIANE OFFRE EAU +ASST'", "QRTUVW"],
        ["MEDIANE OFFRE EAU", 4, solid({ theme: 8, tint: 0.5999938962981048 }), "'POUR MEDIANE OFFRE EAU'", "QRTU"],
      ];
      for (const [lab, rr, fl, sh, cols] of blocs) {
        set(ws, rr, 16, lab, { fill: fl, alignment: { wrapText: true } });
        "QRTUVW".split("").forEach((c, k) => {
          const cell = ws.getCell(`${c}${rr}`);
          cell.fill = fl; cell.numFmt = DATE;
          if (cols.includes(c)) cell.value = { formula: `${sh}!${colL(k + 1)}1` };
        });
      }
      const f5 = solid({ theme: 9, tint: 0.3999755851924192 });
      set(ws, 5, 16, "MEDIANE GENERAL", { fill: f5, alignment: { wrapText: true } });
      for (const c of "QRTUVW") { const cell = ws.getCell(`${c}5`); cell.value = { formula: `MEDIAN(${c}${first}:${c}${last})` }; cell.fill = f5; cell.numFmt = DATE; }
      const larg = { A: 11, B: 15.6, C: 13.3, D: 14.3, E: 24, F: 19.3, G: 18.9, H: 45.7, I: 35, J: 39.3, K: 9, L: 7, M: 11.7,
        N: 20.6, O: 18.4, P: 25.6, Q: 22.7, R: 20.3, S: 25, T: 25, U: 22.4, V: 26.6, W: 24.1, X: 24.1, Y: 18.4, Z: 18.9,
        AA: 18.9, AB: 18.9, AC: 18.9, AH: 11.9, AI: 7.6, AJ: 7.4 };
      for (const [k, w] of Object.entries(larg)) ws.getColumn(k).width = w;
      for (let j = 37; j <= ent.length; j++) ws.getColumn(j).width = 31.1;
    }

    if (listing) {
      const df = d.dateFacturation;
      ws = copierTable(wb, `LISTING ACTIFS ${dmy(df).replace(/\//g, "-")}`, listing);
      const im = listing.col("MODE_PAIEMENT");
      if (im != null) ws.views = [{ state: "frozen", xSplit: im + 1, ySplit: 1 }];
    }
    if (T.population) {
      copierTable(wb, "POPULATION TRAIN", T.population, { A: 11.4, C: 34, D: 30 });
      const manquants = new Set(T.population.rows.map((r) => ref(T.population.get(r, "Référence"))).filter((x) => !d.refsFacturees.has(x)));
      nonFact("NON FACT FINAL", T.population, manquants, "justification ", false);
    }
    if (T.factures) copierTable(wb, "FACTURES TOUS ETATS", T.factures, { E: 15, K: 38 });
    if (T.m2a) copierTable(wb, "DECHARGEMENT M2A", T.m2a);
    return wb;
  }

  // ------------------------------------------------------------------ contrôle des tendances
  const LIGNES = { 60: ["E50E", "E50EPN"], 61: ["E60C", "E60EPN"], 62: ["P30A"], 63: ["P40A"], 64: ["T40A"], 65: ["T20A"],
    66: ["T30A"], 79: ["A50C"], 80: ["A60C", "A60EPN"], 81: ["P50A"], 82: ["T50A"] };
  const COMMENTAIRE = { 60: 68, 61: 69, 62: 70, 63: 71, 64: 72, 65: 73, 66: 74, 79: 84, 80: 85, 81: 86, 82: 87 };
  const CODES_EP = ["E60EPN", "E60C"], CODES_ASST = ["A60EPN", "A60C"];

  const fr = (x, dec = 0) => {
    const s = Math.abs(x).toFixed(dec).split(".");
    s[0] = s[0].replace(/\B(?=(\d{3})+(?!\d))/g, " ");
    return (x < 0 ? "-" : "") + s.join(",");
  };
  const frs = (x) => (x > 0 ? "+" : "") + fr(x);
  const pct = (x) => (x > 0 ? "+" : "") + (x * 100).toFixed(1).replace(".", ",") + " %";
  const puTxt = (p) => String(+p.toPrecision(6)).replace(".", ",");

  const somme = (lignes, codes, k) => lignes.reduce((a, r) => a + (codes.includes(r[5]) ? num(r[k]) : 0), 0);
  function puDominant(lignes, codes) {
    const c = counter();
    for (const r of lignes) if (codes.includes(r[5]) && num(r[9])) inc(c, num(r[9]), Math.max(num(r[11]), 1));
    return c.size ? mostCommon(c) : null;
  }
  const joursAbo = (l) => somme(l, ["E50E", "E50EPN"], 11) - l.reduce((a, r) => a + (r[5] === "E50E" && num(r[9]) === 0 ? num(r[11]) : 0), 0);
  function entrantsSortants(n, n1, codes) {
    const v26 = new Map(), v25 = new Map();
    for (const r of n) if (codes.includes(r[5])) inc(v26, r[2], num(r[11]));
    for (const r of n1) if (codes.includes(r[5])) inc(v25, r[2], num(r[11]));
    let ent = 0, sor = 0;
    for (const [k, v] of v26) if (!v25.has(k)) ent += v;
    for (const [k, v] of v25) if (!v26.has(k)) sor += v;
    return { ent, sor, v25, v26 };
  }

  function commentaires(n, n1, an, nbN, nbN1, es) {
    const out = {};
    for (const [lg, codes] of Object.entries(LIGNES)) {
      const ligne = +lg, cell = `H${COMMENTAIRE[ligne]}`;
      const ca = somme(n, codes, 13), ca1 = somme(n1, codes, 13);
      const vol = somme(n, codes, 11), vol1 = somme(n1, codes, 11);
      const p = puDominant(n, codes), p1 = puDominant(n1, codes);
      if (!ca && !ca1) { out[cell] = "N/A"; continue; }
      if (!ca1) { out[cell] = `Nouvelle ligne facturée en ${an}`; continue; }
      if (!ca) { out[cell] = "N'est plus facturé"; continue; }
      const unite = (ligne === 60 || ligne === 79) ? "€/jour" : "€/m³";
      let tarif = "";
      if (p && p1 && Math.abs(p - p1) > 1e-9) {
        const sens = p > p1 ? "Hausse" : "Baisse";
        const fort = (p / p1 > 2 || p1 / p > 2) ? "Forte " + sens.toLowerCase() : sens;
        tarif = `${fort} du tarif en ${an} (${puTxt(p1)} → ${puTxt(p)} ${unite})`;
      }
      if (ligne === 60) {
        const j = joursAbo(n), j1 = joursAbo(n1), df = nbN - nbN1;
        const debut = df ? `${fr(Math.abs(df))} factures en ${df < 0 ? "moins" : "plus"} (${fr(nbN)} contre ${fr(nbN1)})` : `Même nombre de factures (${fr(nbN)})`;
        out[cell] = `${debut} : ${frs(j - j1)} jours d'abonnement facturés (${j1 ? pct((j - j1) / j1) : "n.c."})` + (tarif ? `. ${tarif}` : ". Tarif inchangé");
      } else if (ligne === 61 || ligne === 80) {
        const e = es[ligne === 61 ? "EP" : "ASST"];
        const nom = ligne === 61 ? "Etude EP" : "Etude ASST";
        out[cell] = `Volume ${frs(vol - vol1)} m³ (${vol1 ? pct((vol - vol1) / vol1) : "n.c."}), dont entrants / sortants : ${frs(e.ent - e.sor)} m³ (voir ${nom})` + (tarif ? `. ${tarif}` : ". Tarif inchangé");
      } else {
        out[cell] = tarif || (p ? `Tarif inchangé (${puTxt(p)} ${unite}) : évolution identique au volume` : "Tarif inchangé");
      }
    }
    return out;
  }

  function ecrireJdv(ws, lignes) {
    lignes.forEach((row, i) => row.forEach((v, j) => {
      const cell = ws.getCell(i + 2, j + 1);
      cell.value = v === undefined ? null : v;
      if (j === 2) cell.numFmt = "@";
      else if (j === 7 || j === 8) cell.numFmt = DATE;
    }));
  }

  async function remplirControle(ExcelJS, template, d) {
    const n = d.jdvN, n1 = d.jdvN1, an = d.annee, an1 = an - 1;
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(template);
    const alertes = [];
    ecrireJdv(wb.getWorksheet("JDV N"), n);
    if (n1) ecrireJdv(wb.getWorksheet("JDV N-1"), n1);

    // Tarifs
    let ws = wb.getWorksheet("Tarifs");
    const okFont = { name: "Calibri", size: 11, color: { theme: 9 } };
    [["A3", "Étiquettes de lignes"], ["B3", "Somme de MNT HT"], ["C3", an1], ["D3", an]].forEach(([a, v]) => (ws.getCell(a).value = v));
    const piv = new Map();
    for (const r of n) {
      if (!piv.has(r[5])) piv.set(r[5], new Map());
      const m = piv.get(r[5]); const pu = num(r[9]);
      if (!m.has(pu)) m.set(pu, { mnt: 0, years: counter() });
      const e = m.get(pu); e.mnt += num(r[12]);
      const fin = toDate(r[8]); if (isDate(fin)) inc(e.years, fin.getUTCFullYear());
    }
    let row = 4, total = 0;
    for (const code of [...piv.keys()].sort()) {
      const pus = piv.get(code);
      const tc = [...pus.values()].reduce((a, e) => a + e.mnt, 0);
      set(ws, row, 1, code, { alignment: { horizontal: "left" } }); set(ws, row, 2, tc); row++;
      for (const pu of [...pus.keys()].sort((a, b) => a - b)) {
        const e = pus.get(pu);
        set(ws, row, 1, pu, { alignment: { horizontal: "left", indent: 1 } }); set(ws, row, 2, e.mnt);
        if (e.years.size) {
          const y = mostCommon(e.years), k = e.years.get(y), tot = [...e.years.values()].reduce((a, b) => a + b, 0);
          if (k < 0.9 * tot) { set(ws, row, 5, "à vérifier : périodes sur plusieurs années", { font: okFont }); alertes.push(`tarif ${code} ${puTxt(pu)} : périodes sur plusieurs années`); }
          else if (y === an) set(ws, row, 4, "ok", { font: okFont });
          else if (y === an1) set(ws, row, 3, "ok", { font: okFont });
          else set(ws, row, 5, `tarif ${y} (rattrapage)`, { font: okFont });
        }
        row++;
      }
      total += tc;
    }
    set(ws, row, 1, "Total général", { alignment: { horizontal: "left" } }); set(ws, row, 2, total);
    ws.getColumn("A").width = 20; ws.getColumn("B").width = 18; ws.getColumn("E").width = 34;

    // Concordance VOL
    ws = wb.getWorksheet("Concordance VOL");
    ["Étiquettes de lignes", "Somme de VOLFACT", "Somme de MNT HT", "Somme de MNTMTTC"].forEach((h, j) => set(ws, 3, j + 1, h));
    const agg = new Map();
    for (const r of n) { if (!agg.has(r[5])) agg.set(r[5], [0, 0, 0]); const a = agg.get(r[5]); a[0] += num(r[11]); a[1] += num(r[12]); a[2] += num(r[13]); }
    row = 4;
    for (const code of [...agg.keys()].sort()) {
      [code, ...agg.get(code)].forEach((v, j) => set(ws, row, j + 1, v, row === 4 ? { fill: solid({ argb: "FF00B050" }) } : undefined));
      row++;
    }
    const tots = [0, 1, 2].map((k) => [...agg.values()].reduce((a, x) => a + x[k], 0));
    ["Total général", ...tots].forEach((v, j) => set(ws, row, j + 1, v));
    [20, 18, 18, 20].forEach((w, j) => (ws.getColumn(j + 1).width = w));

    // TES + études
    const es = { EP: { ent: 0, sor: 0 }, ASST: { ent: 0, sor: 0 } };
    if (n1) {
      ws = wb.getWorksheet("TES");
      ws.getCell("L1").value = `VOLFACT ${an}`; ws.getCell("M1").value = `VOLFACT ${an1}`;
      let i = 2;
      for (const [src, courant] of [[n, true], [n1, false]]) {
        for (const r of src) {
          const vals = [...r.slice(0, 11), ...(courant ? [r[11], null] : [null, r[11]]), ...r.slice(12, 16)];
          vals.forEach((v, j) => {
            const cell = ws.getCell(i, j + 1); cell.value = v === undefined ? null : v;
            if (j === 2) cell.numFmt = "@"; else if (j === 7 || j === 8) cell.numFmt = DATE;
          });
          i++;
        }
      }
      for (const [sheet, codes, key] of [["Etude EP", CODES_EP, "EP"], ["Etude ASST", CODES_ASST, "ASST"]]) {
        const w = wb.getWorksheet(sheet);
        const r = entrantsSortants(n, n1, codes);
        es[key] = r;
        w.getCell("B1").value = `Somme de VOLFACT ${an1}`; w.getCell("C1").value = `Somme de VOLFACT ${an}`;
        w.getCell("J2").value = an1; w.getCell("J3").value = an;
        const style = JSON.parse(JSON.stringify(w.getCell("D2").style || {}));
        const refs = [...new Set([...r.v25.keys(), ...r.v26.keys()])].sort();
        refs.forEach((rf, k0) => {
          const k = k0 + 2;
          const a = w.getCell(k, 1); a.value = rf; a.numFmt = "@";
          w.getCell(k, 2).value = r.v25.has(rf) ? r.v25.get(rf) : null;
          w.getCell(k, 3).value = r.v26.has(rf) ? r.v26.get(rf) : null;
          const c = w.getCell(k, 4);
          c.value = { formula: `IF(C${k}="","sortant",IF(B${k}="","entrant",IF(AND(B${k}<>0,C${k}<>0),C${k}-B${k},"rien")))` };
          c.style = JSON.parse(JSON.stringify(style));
        });
        w.getCell("K5").value = { formula: "K4-G4" };
        w.getCell("L5").value = { formula: "(((K3-G4)-K2)/K2)" };
      }
    }

    // CONTROLE CA
    ws = wb.getWorksheet("CONTROLE CA");
    const S = (a, v) => (ws.getCell(a).value = v === undefined ? null : v);
    S("R2", d.lot); S("F4", d.communes.join(", "));
    S("K28", d.dateFacturation); S("M28", d.nbFactures); S("O28", d.train); S("Q28", d.population);
    if (d.nbMensus) { S("J29", null); S("J30", "X"); }
    S("M30", d.nbMensus);
    S("F34", d.abo[0]); S("H34", d.abo[1]); S("F35", d.conso[0]); S("H35", d.conso[1]);
    const p = d.prec;
    if (p) {
      S("K13", toDate(p.K28)); S("M13", p.M28); S("O13", p.M30 || 0); S("Q13", p.O28); S("P14", p.Q28);
      S("F21", toDate(p.F34)); S("H21", toDate(p.H34)); S("F22", toDate(p.F35)); S("H22", toDate(p.H35));
    }
    if (n1) {
      for (const [lg, codes] of Object.entries(LIGNES)) {
        S(`L${lg}`, Math.round(somme(n1, codes, 13) * 100) / 100);
        S(`O${lg}`, +lg === 60 ? joursAbo(n1) : somme(n1, codes, 11));
      }
      const nbN1 = p ? p.M28 : new Set(n1.map((r) => r[0])).size;
      for (const [cell, txt] of Object.entries(commentaires(n, n1, an, d.nbFactures, nbN1, es))) S(cell, txt);
    }
    wb.calcProperties.fullCalcOnLoad = true;
    return { wb, alertes };
  }

  // ------------------------------------------------------------------ orchestration
  /**
   * files : [{name, buffer(ArrayBuffer|Buffer)}], template : contrôle vierge (ArrayBuffer|Buffer)
   * Renvoie { outputs: [{name, buffer}], rapport, ok }
   */
  Factu.run = async function ({ ExcelJS, files, template, lot: lotForce, log: logFn, progress }) {
    const journal = [];
    const log = (m = "") => { journal.push(m); if (logFn) logFn(m); };
    const step = progress || (() => {});
    const t0 = Date.now();
    const trouves = {};
    const add = (k, t) => (trouves[k] = trouves[k] || []).push(t);

    log(`Lecture de ${files.length} fichier(s)…`);
    for (const f of files) {
      if (f.name.startsWith("~$")) continue;
      await step(`Lecture de ${f.name}`);
      try {
        const tables = await lireFichier(ExcelJS, f, log);
        if (!tables.length) log(`  - aucun export reconnu dans : ${f.name}`);
        for (const t of tables) add(t.kind, t);
      } catch (e) {
        log(`  ! impossible de lire ${f.name} : ${e.message || e}`);
      }
    }

    // population / sans facture
    const abos = trouves.abonnements || [];
    delete trouves.abonnements;
    const reste = [];
    for (const t of abos) {
      const nm = norm(t.name + " " + t.sheet);
      if (nm.includes("SANS")) add("sans_facture", t);
      else if (nm.includes("POPUL")) add("population", t);
      else reste.push(t);
    }
    reste.sort((a, b) => b.rows.length - a.rows.length).forEach((t) => add(trouves.population ? "sans_facture" : "population", t));

    if (!trouves.jdv) { log("\n!! Aucun JDV trouvé : impossible de continuer."); return { ok: false, rapport: journal.join("\n"), outputs: [] }; }
    const parAn = new Map();
    for (const t of trouves.jdv) { const y = anneeJdv(t); if (!parAn.has(y)) parAn.set(y, []); parAn.get(y).push(t); }
    const anJdv = Math.max(...parAn.keys());
    const jdvT = choisirJdv(parAn.get(anJdv));
    const jdvN1T = parAn.has(anJdv - 1) ? choisirJdv(parAn.get(anJdv - 1)) : null;

    const T = {};
    for (const [k, v] of Object.entries(trouves)) if (k !== "jdv" && k !== "controle_n1") T[k] = v[0];
    T.jdv = jdvT;
    log("\nFichiers reconnus :");
    for (const k of ["jdv", "listing", "ctrl", "factures", "anomalies", "population", "sans_facture", "m2a"])
      log(`  ${NOMS[k].padEnd(32)} ${T[k] ? T[k].toString() : "— absent"}`);
    if (jdvN1T) log(`  ${"JDV de l'an dernier".padEnd(32)} ${jdvN1T}`);
    for (const c of trouves.controle_n1 || []) log(`  ${"Contrôle d'une année précédente".padEnd(32)} ${c.name}`);

    const n = lignesJdv(jdvT);
    const refsFact = new Set(n.map((r) => r[2]));
    const listing = T.listing;

    // lot
    let lot = lotForce ? +lotForce : null, groupe = null;
    if (listing) {
      const gc = counter(); for (const r of listing.rows) inc(gc, listing.get(r, "GROUPE_FACTURATION"));
      groupe = mostCommon(gc);
      const m = /LOT\s*0*(\d+)/.exec(String(groupe).toUpperCase());
      if (m && !lot) lot = +m[1];
    }
    if (!lot) for (const f of files) { const m = /LOT[ _-]*0*(\d+)/.exec(norm(f.name)); if (m) { lot = +m[1]; break; } }
    if (!lot) { log("\n!! Numéro de lot introuvable : indique-le dans la case « Lot » puis relance."); return { ok: false, rapport: journal.join("\n"), outputs: [] }; }
    groupe = groupe || `EPN Masse Lot ${lot}`;

    // date, train
    const fac = T.factures;
    let dateFact = null, train = null, finTrait = null;
    if (fac) {
      const emis = fac.rows.map((r) => toDate(fac.get(r, "Date émission"))).filter(isDate);
      const calc = fac.rows.map((r) => toDate(fac.get(r, "Date calcul"))).filter(isDate);
      const base = emis.length ? emis : calc;
      if (base.length) dateFact = jour(new Date(Math.min(...base.map((x) => x.getTime()))));
      if (calc.length) finTrait = new Date(Math.max(...calc.map((x) => x.getTime())));
      const tr = counter(); for (const r of fac.rows) inc(tr, `${fac.get(r, "Exercice")}-${fac.get(r, "N° rôle")}`);
      train = mostCommon(tr);
    }
    if (!dateFact) {
      const fins = n.map((r) => toDate(r[8])).filter(isDate).sort((a, b) => a - b);
      dateFact = jour(fins[Math.floor(fins.length / 2)]);
      log(`\n! Pas d'export « Factures tous états » : date de facturation estimée au ${dmy(dateFact)}, à corriger.`);
    }
    const annee = dateFact.getUTCFullYear(), sem = dateFact.getUTCMonth() >= 6 ? 2 : 1;
    train = train || "x";

    const nbFact = new Set(n.map((r) => r[0])).size;
    let population;
    if (T.population) population = new Set(T.population.rows.map((r) => ref(T.population.get(r, "Référence")))).size;
    else { population = nbFact + (T.sans_facture ? T.sans_facture.rows.length : 0); log("! Pas d'export « Population totale » : population estimée (factures + abonnements sans facture)."); }

    const lst = new Map();
    if (listing) for (const r of listing.rows) lst.set(ref(listing.get(r, "REF_ABONNEMENT", "REF_PDS")), r);
    let nbMensus = 0; const communes = new Set();
    for (const x of refsFact) {
      const r = lst.get(x); if (!r) continue;
      if (listing.get(r, "MODE_PAIEMENT") === "Mensualisation") nbMensus++;
      const v = listing.get(r, "VILLE_PDI", "VILLE_PDL"); if (v) communes.add(v);
    }

    let abo, conso;
    const ctrl = T.ctrl;
    if (ctrl) {
      const md = (c) => mediane(ctrl.rows.map((r) => toDate(ctrl.get(r, c))));
      abo = [md("Période Abo EP. Debut"), md("Période Abo EP. Fin")];
      conso = [md("Période Conso EP. Debut"), md("Période Conso EP. Fin")];
    } else {
      const mj = (codes, k) => mediane(n.filter((r) => codes.includes(r[5])).map((r) => toDate(r[k])));
      abo = [mj(["E50EPN", "E50E"], 7), mj(["E50EPN", "E50E"], 8)];
      conso = [mj(["E60EPN", "E60C"], 7), mj(["E60EPN", "E60C"], 8)];
      log("! Pas d'export « Contrôle factu » : périodes calculées sur le JDV.");
    }

    // contrôle de l'an dernier
    let prec = null, jdvN1 = null;
    const cands = [];
    for (const c of trouves.controle_n1 || []) {
      const k28 = toDate(c.v.K28), r2 = c.v.R2;
      if (!isDate(k28) || k28 >= dateFact) continue;
      if (r2 != null && r2 !== "" && String(r2).trim().replace(/^0+/, "") !== String(lot)) continue;
      const meme = k28.getUTCFullYear() === annee - 1 && ((k28.getUTCMonth() >= 6) === (sem === 2));
      cands.push({ meme, k28, c });
    }
    if (cands.length) {
      cands.sort((a, b) => (a.meme - b.meme) || (a.k28 - b.k28));
      const best = cands[cands.length - 1];
      prec = best.c.v; jdvN1 = best.c.jdv;
      log(`\nFacturation précédente : ${best.c.name} (facturée le ${dmy(best.k28)})`);
      if (!best.meme) log("! Ce n'est pas la même période de l'an dernier : vérifie la comparaison.");
    } else if (jdvN1T) {
      jdvN1 = lignesJdv(jdvN1T);
      log("\n! Pas de contrôle de l'an dernier : comparaison faite avec le JDV N-1 fourni ; date, nombre de factures et population N-1 à saisir à la main.");
    } else {
      log(`\n! Aucun contrôle de l'an dernier pour le lot ${lot} parmi les fichiers déposés : la partie N-1 restera vide.`);
    }

    const d = { tables: T, jdvN: n, jdvN1, prec, lot, groupe, annee, dateFacturation: dateFact, train,
      nbFactures: nbFact, population, nbMensus, communes: [...communes].sort(), abo, conso,
      refsFacturees: refsFact, finTraitement: finTrait };

    const L2 = String(lot).padStart(2, "0");
    log(`\nLot ${lot} — facturation du ${dmy(dateFact)} (S${sem} ${annee}), train ${train}`);
    log(`  ${nbFact} factures, population ${population}, ${population - nbFact} sans facture, ${nbMensus} mensualisé(s)`);
    log(`  période abonnement ${dmy(abo[0])} → ${dmy(abo[1])}, consommation ${dmy(conso[0])} → ${dmy(conso[1])}`);

    await step("Création de l'essai");
    log("\nCréation de l'essai…");
    const wbE = construireEssai(ExcelJS, d);
    if (d.populationListing && d.populationListing !== population)
      log(`! Le listing compte ${d.populationListing} abonnés non mensualisés, la population totale ${population} : à vérifier.`);
    const essaiBuf = await wbE.xlsx.writeBuffer();

    await step("Création du contrôle des tendances (le plus long)");
    log("Création du contrôle des tendances…");
    const { wb: wbC, alertes } = await remplirControle(ExcelJS, template, d);
    for (const a of alertes) log(`! ${a}`);
    const ctrlBuf = await wbC.xlsx.writeBuffer();

    log(`\nTerminé en ${Math.round((Date.now() - t0) / 1000)} s.`);
    log("À faire à la main : relire les commentaires « Validation des évolutions », remplir les colonnes commentaire / justification de l'essai.");
    log(`Garde le contrôle produit : l'an prochain, dépose-le avec les exports pour la comparaison.`);
    const outputs = [
      { name: `ESSAI_EPN_LOT_${L2}_${annee}.xlsx`, buffer: essaiBuf },
      { name: `EVREUX_LOT_${L2}_ControlesTendances_S${sem}_${annee}.xlsx`, buffer: ctrlBuf },
    ];
    return { ok: true, outputs, rapport: journal.join("\n"), lot, annee, sem };
  };

  if (typeof module !== "undefined" && module.exports) module.exports = Factu;
  else root.Factu = Factu;
})(typeof self !== "undefined" ? self : this);
