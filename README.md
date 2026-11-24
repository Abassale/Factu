# Factu

Contrôles de tendances de facturation.

- `modeles/` : fichiers de référence déjà remplis, à suivre comme exemple.
  - `EVREUX_LOT_06_ControlesTendances_S2_2026.xlsx` : 271 - Évreux Porte de Normandie, lot 06, 2e semestre 2026.
- `lots/` : fichiers de contrôle remplis.
  - `EVREUX_LOT_12_ControlesTendances_S2_2026.xlsx` : lot 12 (facturation de novembre 2026, train 2026-393). Comparé à la facturation de novembre 2025 (train 2025-14).
- `scripts/remplir_controle_tendances.py` : script qui remplit le fichier vierge à partir du JDV, du fichier d'essai, du modèle, du JDV N-1 et du contrôle N-1 (`commentaires_lot12.json` : commentaires de validation).
