#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Générateur de présentation BMD² - Fromagerie Tyrode
====================================================

Ce script génère une présentation PowerPoint à partir de données JSON
contenant des analyses BMD² (Business Model Digital Dynamique).

Usage:
    python generate_bmd2_pptx.py <fichier_json> [fichier_sortie.pptx] [fichier_config.yaml]

Exemple:
    python generate_bmd2_pptx.py data_bmd2.json presentation.pptx config.yaml

Prérequis:
    pip install python-pptx pyyaml

Format JSON attendu (fichier JSONL ou tableau JSON):
    {"input": "Contexte: ... Tâche: ...", "output": {...}}

Auteur: Claude
Version: 1.0
"""

import json
import re
import sys
from pathlib import Path
from typing import Optional

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.oxml.ns import nsmap
except ImportError:
    print("Erreur: python-pptx n'est pas installé.")
    print("Installez-le avec: pip install python-pptx")
    sys.exit(1)

try:
    import yaml
    YAML_AVAILABLE = True
except ImportError:
    YAML_AVAILABLE = False


# ============================================================================
# CONFIGURATION PAR DÉFAUT
# ============================================================================

DEFAULT_CONFIG = {
    # Métadonnées de la présentation
    "presentation": {
        "title": "Diagnostic BMD²",
        "subtitle": "Fromagerie Tyrode",
        "author": "Générateur BMD²",
        "width_inches": 13.333,  # 16:9 standard
        "height_inches": 7.5
    },
    
    # Polices
    "fonts": {
        "title": "Georgia",
        "heading": "Arial",
        "body": "Arial"
    },
    
    # Tailles de police (en points)
    "font_sizes": {
        "header_subtitle": 11,
        "header_title": 22,
        "block_title": 11,
        "block_body": 10,
        "block_body_small": 9
    },
    
    # Couleurs (format hex sans #)
    "colors": {
        "primary": "1F4E79",
        "primary_light": "2E75B6",
        "header_text": "FFFFFF",
        "header_subtitle": "B4D7F0",
        "text_dark": "333333",
        
        # Couleurs des blocs
        "contexte": {"bg": "F5F5F5", "border": "1F4E79", "title": "1F4E79"},
        "question": {"bg": "FFF8E1", "border": "E65100", "title": "E65100"},
        "maturite": {"bg": "E3F2FD", "border": "1565C0", "title": "1565C0"},
        "posture": {"bg": "FFF3E0", "border": "E65100", "title": "E65100"},
        "complexite": {"bg": "FCE4EC", "border": "C2185B", "title": "C2185B"},
        "axes": {"bg": "F3E5F5", "border": "7B1FA2", "title": "7B1FA2"},
        "recommandation": {"bg": "E8F5E9", "border": "2E7D32", "title": "2E7D32"},
        "jalon": {"bg": "ECEFF1", "border": "455A64", "title": "455A64"}
    },
    
    # Marges et espacements (en inches)
    "layout": {
        "margin_left": 0.4,
        "margin_right": 0.4,
        "margin_top": 0.15,
        "header_height": 0.95,
        "content_padding_top": 0.2,
        "content_padding_bottom": 0.3,
        "column_gap": 0.25,
        "block_gap": 0.12,
        "block_padding": 0.12,
        "border_width": 4,  # en points
        "left_column_ratio": 0.38  # 38% pour la colonne gauche
    },
    
    # Limites de texte (caractères max)
    "text_limits": {
        "contexte": 350,
        "question": 220,
        "dimension_few": 320,      # Quand peu de dimensions (1-3)
        "dimension_medium": 250,   # Quand dimensions moyennes (4)
        "dimension_many": 200      # Quand beaucoup de dimensions (5-6)
    },
    
    # Mapping des titres automatiques (mots-clés -> titre court)
    "title_keywords": {
        "effectuation": "Posture entrepreneuriale (effectuation)",
        "causale": "Logique causale vs effectuelle",
        "maturité digitale MIT": "Grille de maturité MIT",
        "maturité digitale": "Maturité digitale",
        "risques": "Analyse des risques",
        "complexité opérationnelle": "Complexité opérationnelle",
        "complexité organisationnelle": "Complexité organisationnelle",
        "DSIFAT": "Dispositif DSIFAT",
        "Découverte": "Jalon 1 : Découverte",
        "Intégration": "Jalon 3 : Intégration",
        "Formation": "Jalon 4 : Formation",
        "lockers": "Lockers réfrigérés",
        "dégustations": "Dégustations événementielles",
        "site internet": "Amélioration du site web",
        "jeux concours": "Animation communautaire",
        "résistance au changement": "Accompagnement du changement",
        "synthèse": "Synthèse BMD² globale",
        "expérience client": "Expérience client BMD²",
        "data": "Analyse data BMD²",
        "objectifs": "Objectifs et axes BMD²",
        "tendance": "Maturité digitale dynamique",
        "BMD²": "Diagnostic BMD²"
    }
}


# ============================================================================
# FONCTIONS UTILITAIRES
# ============================================================================

def hex_to_rgb(hex_color: str) -> RGBColor:
    """Convertit une couleur hexadécimale en RGBColor."""
    hex_color = hex_color.lstrip('#')
    return RGBColor(
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16)
    )


def truncate_text(text: str, max_length: int) -> str:
    """Tronque le texte intelligemment sans couper les mots."""
    if not text or len(text) <= max_length:
        return text
    
    truncated = text[:max_length - 3]
    last_space = truncated.rfind(' ')
    
    if last_space > max_length * 0.7:
        truncated = truncated[:last_space]
    
    return truncated.rstrip('.,;:') + '...'


def parse_input(input_text: str) -> dict:
    """Parse le champ 'input' pour extraire le contexte et la tâche."""
    contexte_match = re.search(r'Contexte\s*:\s*(.*?)(?=Tâche\s*:|$)', input_text, re.IGNORECASE | re.DOTALL)
    tache_match = re.search(r'Tâche\s*:\s*(.*?)$', input_text, re.IGNORECASE | re.DOTALL)
    
    return {
        'contexte': contexte_match.group(1).strip() if contexte_match else input_text,
        'tache': tache_match.group(1).strip() if tache_match else ''
    }


def generate_title(tache: str, index: int, config: dict) -> str:
    """Génère un titre court à partir de la tâche."""
    keywords = config.get('title_keywords', DEFAULT_CONFIG['title_keywords'])
    
    tache_lower = tache.lower()
    for keyword, title in keywords.items():
        if keyword.lower() in tache_lower:
            return title
    
    return f"Analyse BMD² n°{index + 1}"


def load_config(config_path: Optional[str]) -> dict:
    """Charge la configuration depuis un fichier YAML ou utilise les valeurs par défaut."""
    config = DEFAULT_CONFIG.copy()
    
    if config_path and Path(config_path).exists():
        if not YAML_AVAILABLE:
            print("Warning: pyyaml n'est pas installé, utilisation de la config par défaut")
            return config
        
        with open(config_path, 'r', encoding='utf-8') as f:
            user_config = yaml.safe_load(f)
        
        # Fusion récursive des configs
        def merge_dict(base, override):
            for key, value in override.items():
                if key in base and isinstance(base[key], dict) and isinstance(value, dict):
                    merge_dict(base[key], value)
                else:
                    base[key] = value
        
        merge_dict(config, user_config)
    
    return config


def load_json_data(json_path: str) -> list:
    """Charge les données depuis un fichier JSON ou JSONL."""
    with open(json_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    try:
        # Essayer JSON standard (tableau)
        data = json.loads(content)
        if not isinstance(data, list):
            data = [data]
    except json.JSONDecodeError:
        # Sinon, essayer JSONL
        data = []
        for line in content.strip().split('\n'):
            if line.strip():
                data.append(json.loads(line))
    
    return data


# ============================================================================
# CLASSE PRINCIPALE DE GÉNÉRATION
# ============================================================================

class BMD2PresentationGenerator:
    """Générateur de présentation PowerPoint BMD²."""
    
    def __init__(self, config: dict):
        self.config = config
        self.prs = Presentation()
        
        # Définir les dimensions de la slide
        self.prs.slide_width = Inches(config['presentation']['width_inches'])
        self.prs.slide_height = Inches(config['presentation']['height_inches'])
        
        # Dimensions calculées
        self.slide_width = config['presentation']['width_inches']
        self.slide_height = config['presentation']['height_inches']
        
        # Layout
        layout = config['layout']
        self.margin_left = layout['margin_left']
        self.margin_right = layout['margin_right']
        self.content_width = self.slide_width - self.margin_left - self.margin_right
        
        # Dimensions des colonnes
        self.left_col_width = self.content_width * layout['left_column_ratio']
        self.right_col_width = self.content_width * (1 - layout['left_column_ratio']) - layout['column_gap']
        self.right_col_left = self.margin_left + self.left_col_width + layout['column_gap']
    
    def add_gradient_header(self, slide, cas_number: int, title: str):
        """Ajoute l'en-tête avec dégradé (simulé par rectangle)."""
        config = self.config
        layout = config['layout']
        colors = config['colors']
        fonts = config['fonts']
        sizes = config['font_sizes']
        pres = config['presentation']
        
        # Rectangle d'en-tête (dégradé simulé par couleur unie)
        header = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0),
            Inches(0),
            Inches(self.slide_width),
            Inches(layout['header_height'])
        )
        header.fill.solid()
        header.fill.fore_color.rgb = hex_to_rgb(colors['primary'])
        header.line.fill.background()
        
        # Sous-titre (Diagnostic BMD² — Fromagerie Tyrode)
        subtitle_top = layout['margin_top'] + 0.1
        subtitle_box = slide.shapes.add_textbox(
            Inches(layout['margin_left']),
            Inches(subtitle_top),
            Inches(self.content_width),
            Inches(0.3)
        )
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.word_wrap = False
        p = subtitle_frame.paragraphs[0]
        p.text = f"{pres['title']} — {pres['subtitle']}"
        p.font.name = fonts['body']
        p.font.size = Pt(sizes['header_subtitle'])
        p.font.color.rgb = hex_to_rgb(colors['header_subtitle'])
        
        # Titre principal
        title_top = subtitle_top + 0.28
        title_box = slide.shapes.add_textbox(
            Inches(layout['margin_left']),
            Inches(title_top),
            Inches(self.content_width),
            Inches(0.5)
        )
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        p = title_frame.paragraphs[0]
        p.text = f"Cas n°{cas_number} : {title}"
        p.font.name = fonts['title']
        p.font.size = Pt(sizes['header_title'])
        p.font.color.rgb = hex_to_rgb(colors['header_text'])
        p.font.bold = True
    
    def add_block(self, slide, left: float, top: float, width: float, height: float,
                  title: str, content: str, color_scheme: dict, 
                  with_border: bool = False, italic: bool = False) -> float:
        """Ajoute un bloc coloré avec titre et contenu. Retourne la hauteur réelle."""
        config = self.config
        layout = config['layout']
        fonts = config['fonts']
        sizes = config['font_sizes']
        colors = config['colors']
        
        # Rectangle de fond
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left),
            Inches(top),
            Inches(width),
            Inches(height)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = hex_to_rgb(color_scheme['bg'])
        shape.line.fill.background()
        
        # Ajustement du rayon des coins
        shape.adjustments[0] = 0.05
        
        # Bordure gauche (si demandée)
        if with_border:
            border = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(left),
                Inches(top),
                Inches(0.05),
                Inches(height)
            )
            border.fill.solid()
            border.fill.fore_color.rgb = hex_to_rgb(color_scheme['border'])
            border.line.fill.background()
        
        padding = layout['block_padding']
        text_left = left + padding + (0.08 if with_border else 0)
        text_width = width - (2 * padding) - (0.08 if with_border else 0)
        
        # Titre du bloc
        title_box = slide.shapes.add_textbox(
            Inches(text_left),
            Inches(top + padding * 0.6),
            Inches(text_width),
            Inches(0.25)
        )
        title_frame = title_box.text_frame
        title_frame.word_wrap = False
        p = title_frame.paragraphs[0]
        p.text = title
        p.font.name = fonts['heading']
        p.font.size = Pt(sizes['block_title'])
        p.font.color.rgb = hex_to_rgb(color_scheme['title'])
        p.font.bold = True
        
        # Contenu du bloc
        content_top = top + padding * 0.6 + 0.22
        content_box = slide.shapes.add_textbox(
            Inches(text_left),
            Inches(content_top),
            Inches(text_width),
            Inches(height - padding - 0.25)
        )
        content_frame = content_box.text_frame
        content_frame.word_wrap = True
        p = content_frame.paragraphs[0]
        p.text = content
        p.font.name = fonts['body']
        p.font.size = Pt(sizes['block_body'])
        p.font.color.rgb = hex_to_rgb(colors['text_dark'])
        if italic:
            p.font.italic = True
        p.line_spacing = 1.15
        
        return height
    
    def calculate_block_height(self, content: str, width: float, font_size: int) -> float:
        """Estime la hauteur nécessaire pour un bloc de texte."""
        # Estimation approximative : ~6 caractères par ligne pour la largeur donnée
        chars_per_line = int((width - 0.24) * 11 * (10 / font_size))
        num_lines = max(1, len(content) // chars_per_line + 1)
        line_height = font_size * 1.3 / 72  # Conversion points -> inches
        return 0.35 + (num_lines * line_height)
    
    def add_slide(self, cas: dict, index: int):
        """Ajoute un slide pour un cas BMD²."""
        config = self.config
        layout = config['layout']
        colors = config['colors']
        limits = config['text_limits']
        sizes = config['font_sizes']
        
        # Parser les données
        parsed = parse_input(cas['input'])
        output = cas['output']
        title = generate_title(parsed['tache'], index, config)
        
        # Créer le slide (layout vierge)
        slide_layout = self.prs.slide_layouts[6]  # Layout vierge
        slide = self.prs.slides.add_slide(slide_layout)
        
        # Ajouter l'en-tête
        self.add_gradient_header(slide, index + 1, title)
        
        # Position de départ du contenu
        content_top = layout['header_height'] + layout['content_padding_top']
        available_height = self.slide_height - content_top - layout['content_padding_bottom']
        
        # === COLONNE GAUCHE : Contexte et Question ===
        left_blocks = []
        
        # Bloc Contexte
        contexte_text = truncate_text(parsed['contexte'], limits['contexte'])
        contexte_height = self.calculate_block_height(contexte_text, self.left_col_width, sizes['block_body'])
        left_blocks.append(('CONTEXTE', contexte_text, colors['contexte'], True, False, contexte_height))
        
        # Bloc Question (si présent)
        if parsed['tache']:
            question_text = truncate_text(parsed['tache'], limits['question'])
            question_height = self.calculate_block_height(question_text, self.left_col_width, sizes['block_body'])
            left_blocks.append(("QUESTION D'ANALYSE", question_text, colors['question'], True, True, question_height))
        
        # === COLONNE DROITE : Dimensions BMD² ===
        dimensions = [
            ('maturite_digitale', 'MATURITÉ DIGITALE', colors['maturite']),
            ('posture_entrepreneuriale', 'POSTURE ENTREPRENEURIALE', colors['posture']),
            ('complexite', 'COMPLEXITÉ', colors['complexite']),
            ('axes_bmd2', 'AXES BMD²', colors['axes']),
            ('recommandations', 'RECOMMANDATION', colors['recommandation']),
            ('jalon_dsifat', 'JALON DSIFAT', colors['jalon'])
        ]
        
        # Filtrer les dimensions non vides
        filled_dimensions = [
            (key, label, color_scheme)
            for key, label, color_scheme in dimensions
            if output.get(key, '').strip()
        ]
        
        # Adapter la limite de texte selon le nombre de dimensions
        num_dims = len(filled_dimensions)
        if num_dims <= 3:
            text_limit = limits['dimension_few']
        elif num_dims == 4:
            text_limit = limits['dimension_medium']
        else:
            text_limit = limits['dimension_many']
        
        # Calculer les hauteurs des blocs droits
        right_blocks = []
        for key, label, color_scheme in filled_dimensions:
            content = truncate_text(output[key], text_limit)
            height = self.calculate_block_height(content, self.right_col_width, sizes['block_body'])
            right_blocks.append((label, content, color_scheme, False, False, height))
        
        # Ajuster les hauteurs pour remplir l'espace disponible
        block_gap = layout['block_gap']
        
        # Hauteur totale nécessaire pour la colonne gauche
        left_total = sum(b[5] for b in left_blocks) + block_gap * (len(left_blocks) - 1)
        
        # Hauteur totale nécessaire pour la colonne droite
        right_total = sum(b[5] for b in right_blocks) + block_gap * (len(right_blocks) - 1)
        
        # Dessiner les blocs de gauche
        current_top = content_top
        for title_text, content, color_scheme, with_border, italic, height in left_blocks:
            self.add_block(slide, self.margin_left, current_top, self.left_col_width, 
                          height, title_text, content, color_scheme, with_border, italic)
            current_top += height + block_gap
        
        # Dessiner les blocs de droite
        current_top = content_top
        for title_text, content, color_scheme, with_border, italic, height in right_blocks:
            self.add_block(slide, self.right_col_left, current_top, self.right_col_width,
                          height, title_text, content, color_scheme, with_border, italic)
            current_top += height + block_gap
    
    def generate(self, data: list, output_path: str):
        """Génère la présentation complète."""
        print(f"\n📊 Génération de {len(data)} slides...\n")
        
        for i, cas in enumerate(data):
            try:
                self.add_slide(cas, i)
                print(f"  ✓ Slide {i + 1}/{len(data)} généré")
            except Exception as e:
                print(f"  ✗ Erreur slide {i + 1}: {e}")
        
        # Sauvegarder
        self.prs.save(output_path)
        print(f"\n✓ Présentation sauvegardée : {output_path}")


# ============================================================================
# GÉNÉRATION DU FICHIER DE CONFIGURATION EXEMPLE
# ============================================================================

def generate_sample_config(output_path: str):
    """Génère un fichier de configuration YAML exemple."""
    if not YAML_AVAILABLE:
        print("Erreur: pyyaml n'est pas installé pour générer le fichier config")
        return
    
    with open(output_path, 'w', encoding='utf-8') as f:
        yaml.dump(DEFAULT_CONFIG, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
    
    print(f"✓ Fichier de configuration exemple généré : {output_path}")


# ============================================================================
# POINT D'ENTRÉE
# ============================================================================

def main():
    args = sys.argv[1:]
    
    # Déterminer le répertoire racine du script
    script_dir = Path(__file__).parent.resolve()
    
    # Chemins par défaut basés sur la structure de répertoires
    default_config = script_dir / 'config_bmd2.yaml'
    default_data_dir = script_dir / 'data'
    default_renders_dir = script_dir / 'renders'
    
    if len(args) < 1 or args[0] in ['-h', '--help']:
        print(f"""
╔════════════════════════════════════════════════════════════════╗
║         Générateur de Présentation BMD² (Python)               ║
╚════════════════════════════════════════════════════════════════╝

Structure de répertoires attendue:
    mon-projet/
    ├── generate_bmd2_pptx.py    # Ce script
    ├── config_bmd2.yaml         # Configuration (optionnel)
    ├── data/                    # Données JSON
    │   └── mon_fichier.json
    └── renders/                 # Sorties générées
        └── ma_presentation.pptx

Usage: 
    python generate_bmd2_pptx.py <fichier_json> [fichier_sortie.pptx] [config.yaml]
    python generate_bmd2_pptx.py --generate-config [config.yaml]

Arguments:
    fichier_json        Fichier JSON (chemin relatif depuis data/ ou absolu)
    fichier_sortie      (optionnel) Nom du fichier PPTX (sera créé dans renders/)
                        Par défaut: BMD2_presentation.pptx
    config.yaml         (optionnel) Fichier de configuration personnalisé
                        Par défaut: config_bmd2.yaml à la racine

Options:
    --generate-config   Génère un fichier de configuration exemple
    -h, --help          Affiche cette aide

Prérequis:
    pip install python-pptx pyyaml

Exemples:
    python generate_bmd2_pptx.py data_bmd2.json
    python generate_bmd2_pptx.py data_bmd2.json ma_presentation.pptx
    python generate_bmd2_pptx.py data/autre.json output.pptx config.yaml
""")
        sys.exit(0)
    
    # Génération du fichier config exemple
    if args[0] == '--generate-config':
        config_path = args[1] if len(args) > 1 else str(default_config)
        generate_sample_config(config_path)
        sys.exit(0)
    
    # Arguments normaux
    json_input = args[0]
    output_name = args[1] if len(args) > 1 else 'BMD2_presentation.pptx'
    config_input = args[2] if len(args) > 2 else None
    
    # Résolution du chemin JSON
    json_path = Path(json_input)
    if not json_path.is_absolute():
        # Chercher d'abord dans data/, sinon chemin relatif direct
        json_in_data = default_data_dir / json_input
        if json_in_data.exists():
            json_path = json_in_data
        elif not json_path.exists():
            # Essayer avec le chemin tel quel
            json_path = Path(json_input)
    
    # Résolution du chemin de sortie
    output_path = Path(output_name)
    if not output_path.is_absolute():
        # Créer le répertoire renders/ si nécessaire
        default_renders_dir.mkdir(parents=True, exist_ok=True)
        output_path = default_renders_dir / output_name
    
    # Résolution du chemin config
    if config_input:
        config_path = config_input
    elif default_config.exists():
        config_path = str(default_config)
    else:
        config_path = None
    
    # Vérifier que le fichier JSON existe
    if not json_path.exists():
        print(f"Erreur: Le fichier '{json_path}' n'existe pas.")
        print(f"  Cherché dans: {json_path.resolve()}")
        sys.exit(1)
    
    # Charger la configuration
    config = load_config(config_path)
    
    # Charger les données
    data = load_json_data(str(json_path))
    print(f"\n📂 {len(data)} cas trouvés dans {json_path}")
    
    # Générer la présentation
    generator = BMD2PresentationGenerator(config)
    generator.generate(data, str(output_path))


if __name__ == '__main__':
    main()
