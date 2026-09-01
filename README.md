# Business Plan IA — Lots 1, 2 & 3

## Ce qui existe a ce stade

- `model/Business_Plan_Template.xlsx` : gabarit de reference restructure (12 produits / 20 charges fixes / 15 investissements provisionnes). Produit par `scripts/restructure_template.py` a partir du classeur original `Business_Plan_Boissons_Naturelles_CI_avec_Guide_1.xlsx`. Ne jamais editer ce fichier a la main — rejouer le script si le gabarit doit evoluer.
- `engine/` : moteur Python isole, sans interface, testable en ligne de commande (Lot 1). `engine/preview.py` convertit une feuille recalculee en tableau HTML (Lot 3).
- `data/*.json` : 3 jeux de donnees de test (mono-produit, multi-produits + saisonnalite, investissement lourd).
- `tests/test_engine.py` : suite de non-regression (structure, zero erreur de formule, formules non modifiees, plausibilite financiere).
- `tests/test_preview.py` : verifie l'apercu HTML (11 feuilles, aucune formule brute affichee, valeurs recalculees presentes, format FCFA preserve).
- `backend/app.py` : API FastAPI. `POST /api/generate` (Lot 2, telechargement direct, conserve pour compatibilite) ; `POST /api/preview` (Lot 3, valide + genere + recalcule + retourne le HTML des 11 feuilles, garde le fichier recalcule en cache serveur par `job_id`) ; `GET /api/download/{job_id}` (sert ce meme fichier deja recalcule, sans second recalcul).
- `web/index.html` : formulaire web fidele a la maquette. La soumission appelle `/api/preview`, affiche les 11 feuilles sous forme d'onglets, puis revele le bouton de telechargement final (`/api/download/{job_id}`) une fois l'apercu genere sans erreur.

## Installation

```bash
pip install -r requirements.txt
```

## Utilisation — moteur seul (Lot 1)

```bash
python -m engine.generate data/exemple_mono_produit.json --out sortie.xlsx
```

Necessite LibreOffice installe (recalcul obligatoire avant toute livraison — cf. cahier des charges §3). `--no-recalc` desactive cette etape, a usage debug uniquement.

## Utilisation — application web (Lot 2)

```bash
uvicorn backend.app:app --reload --port 8000
```

Puis ouvrir `http://127.0.0.1:8000/` : le backend sert le formulaire (`web/index.html`) et expose `POST /api/generate` (JSON du projet -> fichier `.xlsx` recalcule en telechargement direct). Validation testee de bout en bout dans un navigateur reel (formulaire rempli -> generation -> telechargement).

Points de correspondance formulaire -> `ProjetInput` :
- Profil de croissance (chips) -> `taux_croissance_a2`/`a3` (valeurs pre-cablees, ou personnalise).
- Saisonnalite (niveaux Faible/Normal/Fort/Pic par mois) -> coefficients numeriques `[0.7, 1.0, 1.2, 1.35]`.
- "Detailler ce cout" -> `cout_matiere_premiere`/`cout_emballage`/`cout_production` au lieu de `cout_variable_unitaire` (le moteur les additionne, cf. decision ci-dessous).
- Evolution des charges par poste (maquette) -> un seul taux global en mode avance (`taux_evolution_charges_a2`/`a3`), coherent avec la restructuration du Lot 1.

## Mapping champ -> cellule (etat actuel)

| Champ JSON | Feuille | Cellules |
|---|---|---|
| `produits[i].nom/prix_unitaire` | `1_HYPOTHESES` | A/B, lignes 6+i |
| `produits[i].volume_initial_mensuel` | `1_HYPOTHESES` | A/B/C, lignes 18+i |
| `produits[i].cout_variable_unitaire` | `1_HYPOTHESES` | A/B, lignes 35+i |
| `produits[i].nom` (libelles) | `2_PREVISION_CA` | A, lignes 6+i / 22+i / 35+i |
| `charges_fixes[i]` | `1_HYPOTHESES` | A/B, lignes 50+i |
| `investissements[i]` | `4_INVESTISSEMENTS` | A/B/C, lignes 6+i |
| `taux_croissance_a2/a3` | `1_HYPOTHESES` | B30 / B31 |
| `taux_evolution_charges_a2/a3` | `1_HYPOTHESES` | B79 / B80 |
| `saisonnalite[0..11]` | `1_HYPOTHESES` | B84..B95 |
| `part_fonds_propres/emprunt`, `taux_interet_annuel`, `duree_remboursement_mois`, `taux_is`, `delai_clients/fournisseurs_jours`, `jours_ouvres_mensuel`, `capacite_max_mensuelle` | `1_HYPOTHESES` | B72..B78, B33, B34 |

`i` est l'index 0-based du produit/charge/investissement dans la liste JSON (max 12 / 20 / 15).

## Ajouter un nouveau champ au formulaire

1. Verifier si une cellule de saisie existe deja pour ce champ dans `1_HYPOTHESES` / `2_PREVISION_CA` / `4_INVESTISSEMENTS` (voir tableau ci-dessus et les commentaires en tete de `scripts/restructure_template.py`).
2. Si oui : ajouter le champ dans `engine/models.py` (dataclass `ProjetInput` ou sous-classe concernee), l'ecrire dans `engine/generator.py::generate_workbook`, et documenter la ligne dans ce README.
3. Si non : le champ necessite une modification du gabarit lui-meme (nouvelle cellule/ligne). Cela touche a la structure du classeur modele — a traiter comme une nouvelle iteration de `scripts/restructure_template.py`, jamais en ecrivant directement une formule ou une nouvelle ligne depuis le moteur.

## Mettre a jour le classeur modele sans casser le moteur

- Ne jamais editer `model/Business_Plan_Template.xlsx` a la main : toute evolution passe par `scripts/restructure_template.py`, qui repart du classeur original et reconstruit les feuilles de facon deterministe et testable.
- Apres toute modification du gabarit, relancer `pytest tests/` : la verification "aucune formule modifiee" et la verification "11 feuilles, meme ordre" detectent une regression de structure.
- Si une ligne/colonne de saisie change de position, mettre a jour en parallele les constantes de position dans `engine/generator.py` (les commentaires en tete du fichier recensent toutes les plages ecrites) et le tableau de mapping ci-dessus.

## Decisions actees

- **Champs de metadonnees** (`nom_projet`, `forme_juridique`, `pays_devise`, `date_demarrage`) : aucune cellule correspondante dans le classeur (verifie sur l'original et sur le gabarit restructure) et le gabarit n'est pas rouvert pour si peu — ces champs sont purement informatifs. Ils servent au nom de fichier de sortie (CLI : deduit de `nom_projet` si `--out` est omis) et serviront de contexte au texte IA au Lot 4. Jamais ecrits dans une feuille.
- **Cout variable par produit** : la cellule unique `cout_variable_unitaire` par produit (deja testee) est conservee. Le formulaire "Detailler" de la maquette est couvert cote moteur : si `cout_matiere_premiere`/`cout_emballage`/`cout_production` sont fournis a la place de `cout_variable_unitaire`, le moteur les additionne automatiquement (`Produit.__post_init__` dans `engine/models.py`). Aucune nouvelle cellule requise.

## Limites connues (a lever dans un lot ulterieur)

- **Duree de projection fixe a 3 ans / 36 mois** : cablee dans toute la structure du classeur (colonnes de `5_TRESORERIE`, etc.) ; la rendre variable impliquerait une nouvelle restructuration, hors perimetre du Lot 1.
- Textes IA (`description_activite`, `justification_croissance`) : champs presents dans `ProjetInput` mais non encore utilises ni ecrits dans le classeur — prevu au Lot 4.
