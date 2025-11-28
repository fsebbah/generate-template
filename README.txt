# Générateur de Présentations BMD²

Date 2025-11-29

> Outil de génération automatique de présentations PowerPoint pour les diagnostics BMD² (Business Model Digital Dynamique)

---

## 📖 Genèse du projet

### Contexte pédagogique

Ce projet est né d'un besoin concret : transformer des **analyses BMD² structurées en JSON** en présentations PowerPoint professionnelles et lisibles, utilisables dans un contexte pédagogique (BTS, formations professionnelles).

Le cas d'étude initial porte sur la **Fromagerie Tyrode**, une entreprise artisanale confrontée à sa transformation digitale. L'analyse BMD² décompose cette transformation selon plusieurs dimensions :

- **Maturité digitale** : où en est l'entreprise ?
- **Posture entrepreneuriale** : logique causale vs effectuelle
- **Complexité** : opérationnelle et organisationnelle
- **Axes BMD²** : expérience client, digitalisation des process, captation de valeur
- **Recommandations** : actions concrètes
- **Jalons DSIFAT** : Découverte, Sensibilisation, Intégration, Formation, Accompagnement, Transformation

### Évolution du projet

1. **Exploration des formats** : plusieurs options ont été envisagées (DOCX rapport, fiches individuelles, tableaux, PPTX)
2. **Choix du PowerPoint** : format retenu pour sa lisibilité et sa capacité à présenter un cas par slide
3. **Prototype avec html2pptx** : première version utilisant une librairie propriétaire Claude
4. **Portage vers python-pptx** : version finale 100% portable, utilisable sur n'importe quelle machine

### Pourquoi Python ?

| Alternative | Verdict |
|-------------|---------|
| **Node.js + PptxGenJS** | API moins fine pour le positionnement |
| **C# + Open XML SDK** | Plus verbeux, moins adapté au scripting |
| **Python + python-pptx** | ✅ Meilleur compromis : portable, paramétrable, documentation riche |

---

## 🎯 Fonctionnalités

- ✅ Génération automatique de slides à partir de données JSON
- ✅ Design professionnel avec code couleur par dimension
- ✅ Adaptation automatique au nombre de dimensions renseignées
- ✅ Troncature intelligente des textes longs
- ✅ Titres automatiques basés sur des mots-clés
- ✅ Configuration externalisée (YAML)
- ✅ Structure de projet organisée (data/, renders/)

---

## 📁 Structure du projet

```
projet-bmd2/
├── generate_bmd2_pptx.py    # Script principal
├── config_bmd2.yaml         # Configuration (couleurs, polices, marges)
├── README.md                # Ce fichier
├── data/
│   └── data_bmd2.json       # Données JSON des cas BMD²
└── renders/
    └── *.pptx               # Présentations générées
```

---

## 🚀 Installation

### Prérequis

- Python 3.8 ou supérieur
- pip (gestionnaire de paquets Python)

### Installation des dépendances

```bash
pip install python-pptx pyyaml
```

---

## 💻 Utilisation

### Commande de base

```bash
# Se placer dans le répertoire du projet
cd projet-bmd2

# Générer la présentation (cherche le JSON dans data/)
python generate_bmd2_pptx.py data_bmd2.json
```

### Options avancées

```bash
# Avec un nom de sortie personnalisé
python generate_bmd2_pptx.py data_bmd2.json ma_presentation.pptx

# Avec une configuration personnalisée
python generate_bmd2_pptx.py data_bmd2.json output.pptx ma_config.yaml

# Générer un fichier de configuration exemple
python generate_bmd2_pptx.py --generate-config

# Afficher l'aide
python generate_bmd2_pptx.py --help
```

### Comportement par défaut

| Élément | Valeur par défaut |
|---------|-------------------|
| Fichier JSON | Cherché dans `data/` |
| Fichier de sortie | `renders/BMD2_presentation.pptx` |
| Configuration | `config_bmd2.yaml` à la racine (si présent) |

---

## 📊 Format des données JSON

### Structure attendue

Le script accepte un **tableau JSON** ou un fichier **JSONL** (une ligne par objet).

```json
{
  "input": "Contexte : Description du contexte...\nTâche : Question d'analyse...",
  "output": {
    "maturite_digitale": "Analyse de la maturité...",
    "posture_entrepreneuriale": "Analyse de la posture...",
    "complexite": "Analyse de la complexité...",
    "axes_bmd2": "1) Expérience client... 2) Process... 3) Valeur...",
    "recommandations": "Actions recommandées...",
    "jalon_dsifat": "Jalon concerné..."
  }
}
```

### Champs du JSON

| Champ | Description | Obligatoire |
|-------|-------------|-------------|
| `input` | Contient le contexte et la tâche, séparés par "Tâche :" | ✅ |
| `output.maturite_digitale` | Niveau de maturité digitale | ❌ |
| `output.posture_entrepreneuriale` | Logique causale/effectuelle | ❌ |
| `output.complexite` | Complexité opérationnelle/organisationnelle | ❌ |
| `output.axes_bmd2` | Les 3 axes du BMD² | ❌ |
| `output.recommandations` | Actions préconisées | ❌ |
| `output.jalon_dsifat` | Étape du dispositif DSIFAT | ❌ |

> **Note** : Les champs vides sont automatiquement ignorés. Le script adapte la mise en page selon le nombre de dimensions renseignées.

---

## 🎨 Personnalisation

### Fichier config_bmd2.yaml

La configuration permet de personnaliser sans modifier le code :

```yaml
# Métadonnées de la présentation
presentation:
  title: "Diagnostic BMD²"
  subtitle: "Fromagerie Tyrode"
  author: "Mon nom"

# Polices
fonts:
  title: "Georgia"
  heading: "Arial"
  body: "Arial"

# Tailles de police (en points)
font_sizes:
  header_title: 22
  block_title: 11
  block_body: 10

# Couleurs (format hexadécimal sans #)
colors:
  primary: "1F4E79"
  primary_light: "2E75B6"
  
  # Couleurs des blocs par dimension
  contexte:
    bg: "F5F5F5"
    border: "1F4E79"
    title: "1F4E79"
  
  maturite:
    bg: "E3F2FD"
    border: "1565C0"
    title: "1565C0"
  
  # ... autres dimensions

# Marges et espacements (en pouces)
layout:
  margin_left: 0.4
  margin_right: 0.4
  block_gap: 0.12
  left_column_ratio: 0.38

# Limites de caractères
text_limits:
  contexte: 350
  question: 220
  dimension_few: 320      # 1-3 dimensions
  dimension_medium: 250   # 4 dimensions
  dimension_many: 200     # 5-6 dimensions
```

### Code couleur des dimensions

| Dimension | Couleur de fond | Couleur du titre |
|-----------|-----------------|------------------|
| Contexte | Gris clair | Bleu foncé |
| Question | Jaune pâle | Orange |
| Maturité digitale | Bleu clair | Bleu |
| Posture entrepreneuriale | Orange pâle | Orange |
| Complexité | Rose pâle | Rose foncé |
| Axes BMD² | Violet pâle | Violet |
| Recommandation | Vert pâle | Vert |
| Jalon DSIFAT | Gris bleuté | Gris foncé |

---

## 🖼️ Exemple de rendu

Chaque slide suit cette structure :

```
┌─────────────────────────────────────────────────────────────┐
│  [En-tête bleu avec dégradé]                                │
│  Diagnostic BMD² — Fromagerie Tyrode                        │
│  Cas n°X : Titre du cas                                     │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  ┌──────────────┐    ┌────────────────────────────────────┐ │
│  │ CONTEXTE     │    │ MATURITÉ DIGITALE                  │ │
│  │ ...          │    │ ...                                │ │
│  └──────────────┘    ├────────────────────────────────────┤ │
│  ┌──────────────┐    │ POSTURE ENTREPRENEURIALE           │ │
│  │ QUESTION     │    │ ...                                │ │
│  │ D'ANALYSE    │    ├────────────────────────────────────┤ │
│  │ ...          │    │ RECOMMANDATION                     │ │
│  └──────────────┘    │ ...                                │ │
│                      └────────────────────────────────────┘ │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## 🔧 Fonctionnement interne

### Pipeline de génération

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│  JSON       │ ──▶ │  Parser     │ ──▶ │  Générateur │
│  (données)  │     │  (extract)  │     │  de slides  │
└─────────────┘     └─────────────┘     └─────────────┘
                                               │
┌─────────────┐     ┌─────────────┐            │
│  YAML       │ ──▶ │  Config     │ ───────────┘
│  (config)   │     │  loader     │
└─────────────┘     └─────────────┘
                                               │
                                               ▼
                                        ┌─────────────┐
                                        │  PPTX       │
                                        │  (sortie)   │
                                        └─────────────┘
```

### Étapes détaillées

1. **Chargement** : lecture du JSON et de la configuration YAML
2. **Parsing** : extraction du contexte et de la tâche depuis le champ `input`
3. **Génération du titre** : détection automatique basée sur des mots-clés
4. **Filtrage des dimensions** : seules les dimensions non vides sont affichées
5. **Calcul des hauteurs** : estimation de la taille des blocs selon le contenu
6. **Troncature** : adaptation du texte si trop long (avec "...")
7. **Rendu** : création des formes, zones de texte et mise en forme
8. **Export** : sauvegarde du fichier PPTX

### Génération automatique des titres

Le script détecte des mots-clés dans la tâche pour générer un titre pertinent :

| Mot-clé détecté | Titre généré |
|-----------------|--------------|
| "effectuation" | Posture entrepreneuriale (effectuation) |
| "causale" | Logique causale vs effectuelle |
| "MIT" | Grille de maturité MIT |
| "DSIFAT" | Dispositif DSIFAT |
| "synthèse" | Synthèse BMD² globale |
| ... | ... |

---

## 📚 Références

### Cadre théorique BMD²

- **BMD²** (Business Model Digital Dynamique) : cadre d'analyse de la transformation digitale des entreprises
- **DSIFAT** : méthodologie de déploiement en 6 jalons (Découverte, Sensibilisation, Intégration, Formation, Accompagnement, Transformation)
- **Effectuation** : théorie entrepreneuriale de Saras Sarasvathy (qui suis-je, que sais-je, qui connais-je)
- **Grille MIT** : classification de la maturité digitale (Beginners, Fashionistas, Conservatives, Digirati)

### Ressources techniques

- [python-pptx Documentation](https://python-pptx.readthedocs.io/)
- [Format Open XML (PPTX)](https://docs.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document)

---

## 📝 Licence

Ce projet est mis à disposition pour un usage pédagogique et professionnel.

---

## 🤝 Contributions

Pour adapter ce générateur à d'autres cadres d'analyse :

1. Modifier les dimensions dans `DEFAULT_CONFIG`
2. Adapter le code couleur dans `colors`
3. Mettre à jour les mots-clés de génération de titres dans `title_keywords`
4. Ajuster la structure de parsing si le format JSON diffère

---

*Généré avec l'assistance de Claude (Anthropic) — Novembre 2025*
