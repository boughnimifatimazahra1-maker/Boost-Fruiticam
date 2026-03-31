---
name: excel-dashboard
description: >
  Génère automatiquement un dashboard HTML interactif, téléchargeable et aux couleurs de l'ENCG Settat
  (vert #006633, or #D4A017) à partir d'un fichier Excel uploadé. Couvre tous les types de données :
  financières/comptables, étudiants/notes, commerciales/ventes, logistique, et données génériques.
  Utiliser ce skill dès que l'utilisateur mentionne : "dashboard", "tableau de bord", "visualiser mon Excel",
  "graphiques depuis Excel", "analyser mon fichier Excel", "créer un dashboard", "visualisation de données",
  "KPI depuis Excel", "rapport visuel", ou quand l'utilisateur uploade un fichier .xlsx/.xls et demande
  une visualisation, un résumé graphique, ou une analyse visuelle — même sans mentionner explicitement
  "dashboard". Déclencher aussi pour : "montre-moi mes données", "fais quelque chose de visuel avec ce
  fichier", "analyse et visualise". Le livrable est toujours un fichier .html autonome (Chart.js via CDN).
  Ne pas utiliser si le livrable est un fichier Excel modifié (utiliser le skill xlsx à la place).
---

# Excel Dashboard Skill — ENCG Settat

Ce skill transforme n'importe quel fichier Excel en un dashboard HTML interactif, professionnel,
téléchargeable, aux couleurs de l'ENCG Settat.

## Principe général

- **Sortie unique** : fichier `.html` autonome (tout-en-un, pas de dépendances locales)
- **Branding** : ENCG Settat par défaut (`#006633` vert, `#D4A017` or) — sauf si l'utilisateur demande explicitement autre chose
- **Données supportées** : financières, académiques, commerciales, logistique, génériques
- **Graphiques** : Chart.js 4.x via CDN, sélectionnés automatiquement selon les données

## Workflow

### Étape 1 — Lire le fichier Excel

Le fichier est disponible dans `/mnt/user-data/uploads/`. Utiliser pandas et openpyxl pour le lire :

```python
pip install pandas openpyxl --break-system-packages -q

import pandas as pd

# Lire toutes les feuilles
xl = pd.ExcelFile("/mnt/user-data/uploads/FICHIER.xlsx")
sheets = xl.sheet_names

# Lire la feuille principale (ou toutes si pertinent)
df = pd.read_excel("/mnt/user-data/uploads/FICHIER.xlsx", sheet_name=0)
print(df.dtypes)
print(df.describe())
print(df.head())
```

### Étape 2 — Analyser la structure des données

Identifier automatiquement :

| Type de colonne | Critère | Usage dashboard |
|---|---|---|
| **Numérique** | dtype int/float | Barres, lignes, aires, KPI cards |
| **Catégorielle** | dtype object/str, cardinalité ≤20 | Axes X, filtres, camemberts |
| **Date/Temps** | dtype datetime ou col nommée date/année/mois/trimestre | Séries temporelles |
| **Texte libre** | dtype object/str, cardinalité >20 | Labels uniquement, tableaux |

**Détection automatique du mode par mots-clés dans les noms de colonnes :**

| Mode | Mots-clés détectés | Graphiques prioritaires |
|---|---|---|
| **Académique** | note, mention, étudiant, module, filière, résultat, examen | Barres modules, donut mentions, radar compétences |
| **Financier** | CA, chiffre, revenu, montant, bilan, actif, passif, charges, produit, résultat, trésorerie | LineChart temporel, barres comparatif, KPI financiers |
| **Commercial** | vente, commande, client, produit, quantité, remise, région, commercial | Barres ventes, donut répartition, top N |
| **Logistique** | stock, fournisseur, délai, livraison, référence, entrepôt | Barres stocks, alertes seuil, tableaux |
| **Générique** | aucun mot-clé reconnu | KPI count/sum/mean, barres, donut, tableau |

### Étape 3 — Choisir les graphiques et KPI appropriés

**Algorithme de sélection des graphiques :**
```
1. S'il y a une colonne date → LineChart (évolution temporelle) en premier
2. Pour chaque colonne numérique vs catégorielle → BarChart horizontal ou vertical
3. Si ≤7 catégories avec 1 métrique → DoughnutChart
4. Si ≥3 colonnes numériques comparables → RadarChart
5. Si 2 colonnes numériques corrélables → ScatterChart
6. Toujours : 4-6 KPI Cards en haut
7. Toujours : Tableau de données paginé en bas (10 lignes/page)
```

**KPI Cards recommandées par mode :**

| Mode | KPI 1 | KPI 2 | KPI 3 | KPI 4 |
|---|---|---|---|---|
| Académique | Total étudiants | Moyenne générale | Taux de réussite | Meilleure note |
| Financier | CA Total / Revenus | Évolution vs N-1 | Charges totales | Résultat net |
| Commercial | Total commandes | CA total | Panier moyen | Top produit |
| Logistique | Total références | Stock total | Ruptures | Valeur stock |
| Générique | Count total | Somme col principale | Moyenne | Max |

### Étape 4 — Générer le dashboard HTML

Générer un fichier HTML autonome (tout-en-un) avec :

**Structure HTML obligatoire :**
```
├── Header (titre + logo optionnel + date)
├── KPI Cards row (4-6 métriques clés)
├── Grille de graphiques (2-3 graphiques principaux)
├── Tableau de données (paginé, 10 lignes/page)
└── Footer
```

**Stack technique :**
- **Chart.js 4.x** via CDN pour les graphiques
- **Vanilla JS** (pas de framework) pour la simplicité
- **CSS Grid/Flexbox** pour le layout responsive
- Données embarquées directement en JSON dans le `<script>`

**Template de base :**

```html
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Dashboard — {TITRE}</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <style>
    /* Variables CSS — adapter selon branding */
    :root {
      --primary: {COULEUR_PRIMAIRE};
      --secondary: {COULEUR_SECONDAIRE};
      --bg: #f8f9fa;
      --card-bg: #ffffff;
      --text: #2d3748;
      --border: #e2e8f0;
    }
    /* ... reste du CSS */
  </style>
</head>
<body>
  <!-- Header -->
  <!-- KPI Cards -->
  <!-- Charts Grid -->
  <!-- Data Table -->
  <script>
    const DATA = {/* JSON des données */};
    // Initialisation Chart.js
  </script>
</body>
</html>
```

### Étape 5 — Branding

Le branding ENCG est **appliqué par défaut** dans tous les cas, sauf exception explicite.

**Variables CSS ENCG (toujours utiliser) :**
```css
:root {
  --primary: #006633;        /* Vert ENCG */
  --secondary: #D4A017;      /* Or ENCG */
  --bg: #f0f4f1;             /* Fond légèrement teinté vert */
  --card-bg: #ffffff;
  --text: #1a2e22;
  --text-light: #5a7a65;
  --border: #d0e4d8;
  --shadow: 0 2px 12px rgba(0,102,51,0.08);
}
```

**Header ENCG obligatoire :**
- Gradient `#006633 → #004d26`
- Titre du dashboard + sous-titre "ENCG Settat — Université Hassan 1er"
- Badge doré `#D4A017` indiquant le type de données (ex: "Module M243", "Données Commerciales", etc.)
- Police : Segoe UI / system-ui pour la cohérence institutionnelle

**Exceptions (uniquement si demande explicite) :**
- L'utilisateur spécifie d'autres couleurs → les utiliser
- L'utilisateur dit "sans branding ENCG" → design neutre bleu `#2563EB / #10B981`

**Adaptation par type de données :**

| Mode | Sous-titre header | Badge |
|---|---|---|
| Académique | "ENCG Settat — Résultats Étudiants" | Nom du module ou filière |
| Financier | "ENCG Settat — Analyse Financière" | Exercice ou période |
| Commercial | "ENCG Settat — Performance Commerciale" | Période ou région |
| Générique | "ENCG Settat — Tableau de Bord" | Nom du fichier source |

### Étape 6 — Sauvegarder et présenter

```python
# Sauvegarder dans outputs
output_path = "/mnt/user-data/outputs/dashboard.html"
with open(output_path, "w", encoding="utf-8") as f:
    f.write(html_content)
```

Puis appeler `present_files` avec le chemin du fichier.

---

## Cas particuliers

### Fichier multi-feuilles
Si le fichier Excel contient plusieurs feuilles pertinentes :
- Créer un onglet par feuille dans le dashboard (navigation par tabs)
- Ou demander à l'utilisateur quelle feuille analyser

### Données manquantes
- Ignorer les colonnes avec >50% de NaN
- Afficher un avertissement dans le dashboard si des données ont été exclues
- Ne jamais planter : utiliser `df.dropna()` ou `df.fillna(0)` selon le contexte

### Grandes quantités de données (>1000 lignes)
- Agréger avant d'embarquer en JSON (groupby + sum/mean)
- Limiter le tableau paginé à 500 lignes max
- Afficher le nombre total de lignes dans un KPI "Total enregistrements"

### Colonnes dates mal formatées
```python
# Tentative de parsing robuste
for col in df.columns:
    try:
        df[col] = pd.to_datetime(df[col])
    except:
        pass
```

---

## Checklist qualité avant livraison

- [ ] Le HTML s'ouvre sans erreur dans un navigateur
- [ ] Tous les graphiques se chargent (pas de canvas vide)
- [ ] Les KPI Cards affichent des valeurs sensées
- [ ] Le tableau de données est lisible et paginé
- [ ] Le design est cohérent (couleurs, typographie)
- [ ] Le fichier est autonome (pas de dépendances locales, CDN uniquement)
- [ ] Le titre du dashboard reflète le contenu du fichier

---

## Exemples de prompts déclencheurs

**Dashboard général :**
- "Crée un dashboard à partir de ce fichier Excel"
- "J'ai uploadé un Excel, génère-moi des graphiques"
- "Analyse ce fichier et montre-moi un dashboard"
- "Tableau de bord KPI depuis mon fichier"
- "Fais quelque chose de visuel avec ce fichier"

**Données académiques :**
- "Fais un tableau de bord pour mes résultats d'étudiants"
- "Dashboard des notes du module M243"
- "Visualise les résultats par filière"

**Données financières :**
- "Dashboard financier à partir de mon bilan Excel"
- "Visualise mon compte de résultat"
- "Tableau de bord de mes indicateurs financiers"

**Données commerciales :**
- "Dashboard des ventes du trimestre"
- "Visualise mes données clients"
- "Analyse et graphiques de mes commandes"

**Données génériques :**
- "Montre-moi mes données sous forme visuelle"
- "Résumé graphique de ce fichier"
- "Crée des graphiques depuis mon Excel"
