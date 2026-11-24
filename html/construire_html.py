"""Assemble factu.html (page autonome) : interface + ExcelJS + cœur + contrôle vierge intégrés.

Usage : python html/construire_html.py
"""
import base64, os

ICI = os.path.dirname(os.path.abspath(__file__))
RACINE = os.path.dirname(ICI)


def lire(p):
    with open(p, encoding="utf-8") as f:
        return f.read()


page = lire(os.path.join(ICI, "page.html"))
exceljs = lire(os.path.join(ICI, "vendor", "exceljs.min.js")).replace("</script", "<\\/script")
core = lire(os.path.join(ICI, "factu-core.js")).replace("</script", "<\\/script")
with open(os.path.join(RACINE, "modeles", "CONTROLE_TENDANCES_VIERGE.xlsx"), "rb") as f:
    modele = base64.b64encode(f.read()).decode("ascii")

html = (page.replace("/*EXCELJS*/", exceljs, 1)
            .replace("/*CORE*/", core, 1)
            .replace("/*MODELE*/", modele, 1))
out = os.path.join(RACINE, "factu.html")
with open(out, "w", encoding="utf-8", newline="\n") as f:
    f.write(html)
print(out, f"{os.path.getsize(out) / 1e6:.1f} Mo")
