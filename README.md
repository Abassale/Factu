# Factu

Outil qui prépare automatiquement, pour un lot EPN :

1. **l'essai** mis en forme comme l'essai 2025 (onglets 04.04, POUR MEDIANE, NON FACT, ANO AVANT CALCUL, TCD LISTING…) ;
2. **le contrôle des tendances** rempli et comparé à la facturation de l'an dernier.

## Version la plus simple : `factu.html` (rien à installer)

1. Télécharger le fichier **`factu.html`** et l'ouvrir d'un double-clic (Edge ou Chrome).
2. Glisser dans la page les exports bruts du lot **et le contrôle des tendances de l'an dernier** du même lot
   (dossier `historique\LOT_xx\`).
3. Cliquer sur **Lancer**, puis télécharger l'essai, le contrôle et le rapport.

Tout se passe dans le navigateur : aucune donnée n'est envoyée sur Internet.
Pour reconstruire la page après une modification : `python html/construire_html.py`.

## Version Python (`lancer.bat`)

Même résultat, avec en plus le rangement automatique dans `historique`.

### Installation (une seule fois)

1. Installer Python depuis https://www.python.org (cocher « Add Python to PATH »).
2. Récupérer ce dépôt sur le PC :
   `git clone https://github.com/Abassale/Factu`
   (ou bouton vert **Code → Download ZIP** sur GitHub, puis dézipper).

### Utilisation, à chaque lot

1. Vider le dossier `entree`.
2. Y déposer les exports bruts du lot (Excel `.xlsx` ou `.csv`, noms libres) :
   - le JDV de la facturation ;
   - le listing clients ;
   - le contrôle factu (détail des contrôles factures) ;
   - les factures tous états ;
   - les PDI en anomalie ;
   - la population totale ;
   - les abonnements sans facture (mettre « sans » dans le nom du fichier ou de l'onglet) ;
   - le déchargement M2A.

   Plusieurs exports peuvent être dans un même classeur (un onglet chacun) : l'outil reconnaît chaque onglet à ses colonnes.
3. Double-cliquer sur **`lancer.bat`**.
   Si le numéro de lot n'est pas trouvé dans le listing : ouvrir une invite de commandes dans le dossier et taper `lancer.bat 12`.
4. Les deux fichiers arrivent dans `sortie\LOT_xx_AAAA_Sx\`, avec un `rapport.txt` qui liste les fichiers reconnus, les chiffres clés et les points à vérifier.

Il reste à faire à la main : relire les commentaires « Validation des évolutions » (pré-rédigés à partir des chiffres), et remplir les colonnes commentaire / justification / « A SORTIR DU TRAIN » de l'essai.

## Comparaison avec l'an dernier

L'outil cherche le contrôle de l'année précédente du même lot dans `historique\LOT_xx\` (ou dans `entree`).
Chaque contrôle produit y est rangé automatiquement : il servira de N-1 l'année suivante.
Pour un lot traité pour la première fois avec l'outil, déposer le contrôle de l'an dernier dans `entree` (ou dans `historique\LOT_xx\`).

## Contenu

- `factu.html` : la page autonome ; ses sources sont dans `html/` (`factu-core.js`, `page.html`, ExcelJS dans `html/vendor`).
- `outil/` : le programme Python (`factu.py` lance tout, `lecture.py` reconnaît les exports, `essai.py` et `controle.py` construisent les fichiers).
- `modeles/` : le contrôle des tendances vierge et le modèle rempli du lot 06.
- `historique/` : les contrôles des années précédentes, par lot.
- `lots/`, `essais/` : fichiers du lot 12 préparés le 24/11/2026.
