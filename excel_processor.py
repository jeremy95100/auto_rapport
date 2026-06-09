"""
Excel Processor - Traitement ULTRA-RAPIDE
==========================================
- Parsing regex optimisÃ© (plus rapide que lxml pour Excel)
- Conversion directe vers Parquet via PyArrow streaming
- ParallÃ©lisation des feuilles avec ThreadPoolExecutor
- Polars lazy mode exclusif (scan_parquet)
"""

import os
import uuid
import shutil
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass
from concurrent.futures import ThreadPoolExecutor, as_completed
import zipfile
import io
import re
import time

import polars as pl
import pyarrow as pa
import pyarrow.parquet as pq
from html.parser import HTMLParser


# Fonction pour nettoyer le HTML et garder seulement le texte
class HTMLTextExtractor(HTMLParser):
    """Extracteur de texte depuis HTML"""
    def __init__(self):
        super().__init__()
        self.text_parts = []
        self.skip_data = False

    def handle_starttag(self, tag, attrs):
        # Ignorer le contenu des balises style et script
        if tag in ('style', 'script', 'head'):
            self.skip_data = True
        # Ajouter un saut de ligne pour les balises de bloc
        elif tag in ('br', 'p', 'div', 'tr', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
            self.text_parts.append('\n')

    def handle_endtag(self, tag):
        if tag in ('style', 'script', 'head'):
            self.skip_data = False
        elif tag in ('p', 'div', 'tr', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
            self.text_parts.append('\n')

    def handle_data(self, data):
        if not self.skip_data:
            self.text_parts.append(data)

    def get_text(self):
        text = ''.join(self.text_parts)
        # Nettoyer les espaces multiples et lignes vides
        lines = [line.strip() for line in text.split('\n')]
        lines = [line for line in lines if line]  # Enlever les lignes vides
        return '\n'.join(lines)


def strip_html_tags(html_content: str) -> str:
    """
    Enlève toutes les balises HTML et retourne le texte brut.
    Utilisé pour nettoyer la colonne Body des Emails.
    """
    if not html_content or not isinstance(html_content, str):
        return html_content or ""

    # Détecter si c'est du HTML (avec ou sans < au début)
    html_patterns = ['<div', '<p>', '<br', '<span', '<table', '<a ', '<img',
                     'div>', '/div>', '/p>', '/span>', 'style="', 'class="',
                     'dir="auto"', '&nbsp;', '&amp;', '&#']
    is_html = any(pattern in html_content.lower() for pattern in html_patterns)

    if not is_html:
        return html_content

    # Ajouter < au début si manquant et que ça ressemble à du HTML
    content = html_content
    if not content.strip().startswith('<') and any(content.strip().startswith(tag) for tag in ['div', 'p ', 'span', 'table', 'html', 'body']):
        content = '<' + content

    try:
        parser = HTMLTextExtractor()
        parser.feed(content)
        result = parser.get_text()
        # Si le résultat est vide ou identique, utiliser regex
        if not result.strip() or result == content:
            raise ValueError("Parser failed")
        return result
    except Exception:
        # Fallback: regex robuste pour nettoyer le HTML
        text = content
        # Supprimer les balises style et leur contenu
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
        # Supprimer les balises script et leur contenu
        text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
        # Remplacer <br>, </p>, </div> par des sauts de ligne
        text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</(?:p|div|tr|li|h[1-6])>', '\n', text, flags=re.IGNORECASE)
        # Supprimer toutes les autres balises HTML
        text = re.sub(r'<[^>]+>', '', text)
        # Décoder les entités HTML courantes
        text = text.replace('&nbsp;', ' ')
        text = text.replace('&amp;', '&')
        text = text.replace('&lt;', '<')
        text = text.replace('&gt;', '>')
        text = text.replace('&quot;', '"')
        text = text.replace('&#39;', "'")
        # Nettoyer les espaces multiples
        text = re.sub(r'[ \t]+', ' ', text)
        # Nettoyer les lignes vides multiples
        lines = [line.strip() for line in text.split('\n')]
        lines = [line for line in lines if line]
        return '\n'.join(lines)


# CONFIGURATION

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

MAX_UPLOAD_SIZE_MB = 5000


def cleanup_previous_import():
    """
    Nettoie les fichiers des imports précédents avant un nouvel import.
    - Supprime tous les sous-dossiers dans data/ (fichiers parquet)
    - Supprime les graphiques générés dans uploads/ (chart_*.png)
    """
    import shutil

    cleaned_data = 0
    cleaned_uploads = 0

    # Nettoyer le dossier data/ (sous-dossiers d'imports)
    if DATA_DIR.exists():
        for item in DATA_DIR.iterdir():
            try:
                if item.is_dir():
                    shutil.rmtree(item)
                    cleaned_data += 1
                    print(f"[CLEANUP] Supprimé dossier data: {item.name}")
                elif item.is_file():
                    item.unlink()
                    cleaned_data += 1
            except Exception as e:
                print(f"[CLEANUP] Erreur suppression {item}: {e}")

    # Nettoyer les graphiques dans uploads/ (chart_*.png)
    if UPLOAD_DIR.exists():
        for item in UPLOAD_DIR.iterdir():
            try:
                if item.is_file() and item.name.startswith("chart_") and item.suffix.lower() == ".png":
                    item.unlink()
                    cleaned_uploads += 1
                    print(f"[CLEANUP] Supprimé graphique: {item.name}")
            except Exception as e:
                print(f"[CLEANUP] Erreur suppression {item}: {e}")

    print(f"[CLEANUP] Nettoyage terminé: {cleaned_data} éléments data, {cleaned_uploads} graphiques")
    return {"data_cleaned": cleaned_data, "uploads_cleaned": cleaned_uploads}
MAX_WORKERS = 4  # Threads pour parallÃ©lisation

# Feuilles Ã  ignorer
SKIP_SHEETS = {
    'Summary', 'Image Hashes', 'Aggregated Application Usage',
    'Applications Usage Log', 'Applications', 'Exchange', 'File Uploads',
    'Configurations', 'Databases', 'Autofill', 'Log Entries',
    'Social Media', 'Text', 'Locations View', 'Device Notifications',
    'Cookies', 'Shortcut', 'Watch List Results', 'SharedStrings',
    'Calendar', 'Journey', 'Notes',
    'Locations', 'Timeline'  # SIM Data + Instant Messages preserves pour fichiers carte SIM
}

# Sources Ã  exclure (appliquÃ©es dÃ¨s la conversion Parquet)
EXCLUDED_SOURCES = {
    'Recents', 'InteractionC', 'KnowledgeC',
    'Threads', 'Biome', 'Unified Logs', 'SIM', 'Unspecified App Call',
    'Google Quick Search Box', 'SSRM Heating Log'
}

# Sources à exclure spécifiquement pour Chats (Facebook mais pas Facebook Messenger)
EXCLUDED_CHAT_SOURCES = EXCLUDED_SOURCES | {'Facebook'}

EXCLUDED_CONTACT_SOURCES = EXCLUDED_SOURCES | {'Native Messages'}

# Feuilles oÃ¹ filtrer les sources exclues + colonne Ã  utiliser
SHEETS_WITH_SOURCE_FILTER = {
    'Contacts': 'Source',
    'Call Log': 'Source',
    'Chats': 'Source',
}

# Traduction des valeurs anglais -> français pour certaines colonnes
VALUE_TRANSLATIONS = {
    "Interaction Statuses": {
        # Statuts de contact
        "Chat Participant": "Participant au chat",
        "Phone Book": "Répertoire téléphonique",
        "Phonebook": "Répertoire téléphonique",
        "Follower": "Abonné",
        "Following": "Abonnement",
        "Friend": "Ami",
        "Friends": "Amis",
        "Contact": "Contact",
        "Blocked": "Bloqué",
        "Favorite": "Favori",
        "Favourites": "Favoris",
        "Favorites": "Favoris",
        "Group Member": "Membre du groupe",
        "Group Admin": "Administrateur du groupe",
        "Pending": "En attente",
        "Pending Request": "Demande en attente",
        "Requested": "Demandé",
        "Accepted": "Accepté",
        "Declined": "Refusé",
        "Rejected": "Rejeté",
        "Muted": "Muet",
        "Archived": "Archivé",
        "Hidden": "Masqué",
        "Unknown": "Inconnu",
        "Saved": "Enregistré",
        "Searched": "Recherché",
        "Recent": "Récent",
        "Frequent": "Fréquent",
        "Call History": "Historique d'appels",
        "Message History": "Historique de messages",
        "Subscriber": "Abonné",
        "Subscribed": "Abonné",
        "Close Friend": "Ami proche",
        "Close Friends": "Amis proches",
        "Best Friend": "Meilleur ami",
        "Best Friends": "Meilleurs amis",
        "Mutual Friend": "Ami commun",
        "Mutual Friends": "Amis communs",
        "Family": "Famille",
        "Work": "Travail",
        "Colleague": "Collègue",
        "Business": "Professionnel",
        "Other": "Autre",
        "Shared": "Partagé",
    }
}

# Colonnes à traduire automatiquement
COLUMNS_TO_TRANSLATE = {"Interaction Statuses", "Interaction Status"}


# REGEX COMPILÃ‰S (beaucoup plus rapide que recompiler Ã  chaque fois)


RE_SHEET = re.compile(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', re.IGNORECASE)
RE_SHEET_ALT = re.compile(r'<sheet[^>]*r:id="(rId\d+)"[^>]*name="([^"]+)"', re.IGNORECASE)
RE_REL = re.compile(r'<Relationship[^>]*Id="(rId\d+)"[^>]*Target="([^"]+)"', re.IGNORECASE)
RE_SI = re.compile(r'<si[^>]*>(.*?)</si>', re.DOTALL)
RE_T = re.compile(r'<t[^>]*>([^<]*)</t>')
RE_ROW = re.compile(r'<row[^>]*r="(\d+)"[^>]*>(.*?)</row>', re.DOTALL)
RE_CELL = re.compile(r'<c\s+r="([A-Z]+)\d+"([^>]*)>(?:<v>([^<]*)</v>)?</c>')
RE_CELL_TYPE_S = re.compile(r't="s"')
RE_CLEAN = re.compile(r'_x[0-9A-Fa-f]{4}_')
RE_SPACES = re.compile(r'\s+')



# DATA CLASSES


@dataclass
class ImportResult:
    """RÃ©sultat d'un import Excel"""
    import_id: str
    import_path: Path
    sheets: List[str]
    parquet_files: Dict[str, Path]
    device_info: Dict[str, str]
    row_counts: Dict[str, int]


@dataclass
class SheetInfo:
    """Informations sur une feuille Excel"""
    name: str
    rel_id: str
    file_path: str


# FONCTIONS UTILITAIRES OPTIMISÃ‰ES


def clean_value(value: str) -> str:
    """Nettoie une valeur Excel - version optimisÃ©e"""
    if not value:
        return ""
    value = RE_CLEAN.sub('', value)
    value = value.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
    value = RE_SPACES.sub(' ', value)
    return value.strip()


def safe_filename(name: str) -> str:
    """Convertit un nom de feuille en nom de fichier sÃ»r"""
    return re.sub(r'[^\w\-]', '_', name)


def col_letter_to_index(letter: str) -> int:
    """Convertit une lettre de colonne en index (A=0, B=1, ..., AA=26)"""
    result = 0
    for char in letter:
        result = result * 26 + (ord(char.upper()) - ord('A') + 1)
    return result - 1



# PARSING OPTIMISE


def parse_shared_strings_fast(zf: zipfile.ZipFile) -> List[str]:
    """Parse sharedStrings.xml avec regex - ULTRA RAPIDE"""
    try:
        ss_xml = zf.read('xl/sharedStrings.xml').decode('utf-8')
    except KeyError:
        return []

    shared_strings = []
    for match in RE_SI.finditer(ss_xml):
        si_content = match.group(1)
        texts = RE_T.findall(si_content)
        shared_strings.append(''.join(texts))

    return shared_strings


def get_sheets_info_fast(zf: zipfile.ZipFile) -> Dict[str, SheetInfo]:
    """RÃ©cupÃ¨re les infos des feuilles - version regex optimisÃ©e"""
    sheets: Dict[str, SheetInfo] = {}

    # Parser workbook.xml
    workbook_xml = zf.read('xl/workbook.xml').decode('utf-8')

    for match in RE_SHEET.finditer(workbook_xml):
        name, rel_id = match.group(1), match.group(2)
        if name not in SKIP_SHEETS:
            sheets[name] = SheetInfo(name=name, rel_id=rel_id, file_path='')

    # Format alternatif
    for match in RE_SHEET_ALT.finditer(workbook_xml):
        rel_id, name = match.group(1), match.group(2)
        if name not in SKIP_SHEETS and name not in sheets:
            sheets[name] = SheetInfo(name=name, rel_id=rel_id, file_path='')

    # Parser relations
    try:
        rels_xml = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        rels = {m.group(1): m.group(2) for m in RE_REL.finditer(rels_xml)}

        for name, info in sheets.items():
            if info.rel_id in rels:
                target = rels[info.rel_id]
                if target.startswith('/'):
                    info.file_path = target[1:]
                elif target.startswith('xl/'):
                    info.file_path = target
                else:
                    info.file_path = 'xl/' + target
    except KeyError:
        pass

    return sheets


def parse_sheet_to_parquet_fast(
    zf: zipfile.ZipFile,
    sheet_info: SheetInfo,
    shared_strings: List[str],
    output_path: Path,
    is_sim_file: bool = False
) -> Tuple[int, int]:
    """
    Parse une feuille et Ã©crit directement en Parquet.
    APPLIQUE LE FILTRAGE DES SOURCES EXCLUES dÃ¨s la conversion.
    UNE SEULE PASSE sur le XML.
    Retourne (nombre de lignes conservÃ©es, nombre de lignes filtrÃ©es).

    Si is_sim_file=True, ne filtre pas les sources 'SIM' (pour fichiers carte SIM).
    """
    if not sheet_info.file_path:
        return 0, 0

    try:
        sheet_xml = zf.read(sheet_info.file_path).decode('utf-8')
    except KeyError:
        return 0, 0

    # 1. Extraire headers (ligne 2) en une seule passe
    headers: Dict[str, str] = {}  # {col_letter: col_name}
    header_match = re.search(r'<row[^>]*r="2"[^>]*>(.*?)</row>', sheet_xml, re.DOTALL)

    if not header_match:
        return 0, 0

    for cell_match in RE_CELL.finditer(header_match.group(1)):
        col_letter = cell_match.group(1)
        attrs = cell_match.group(2) or ""
        v_val = cell_match.group(3)

        if v_val is None:
            continue

        if RE_CELL_TYPE_S.search(attrs) and shared_strings:
            try:
                idx = int(v_val)
                if idx < len(shared_strings):
                    headers[col_letter] = clean_value(shared_strings[idx])
            except ValueError:
                pass
        else:
            headers[col_letter] = clean_value(v_val)

    if not headers:
        return 0, 0

    # 2. Préparer les colonnes triées par position
    sorted_cols = sorted(headers.items(), key=lambda x: col_letter_to_index(x[0]))
    col_names = [name for _, name in sorted_cols]
    col_letters = [letter for letter, _ in sorted_cols]
    letter_to_idx = {letter: i for i, letter in enumerate(col_letters)}

    # 3. DÃ©terminer si cette feuille nÃ©cessite un filtrage des sources
    source_filter_col = SHEETS_WITH_SOURCE_FILTER.get(sheet_info.name)
    source_col_idx = None
    if source_filter_col and source_filter_col in col_names:
        source_col_idx = col_names.index(source_filter_col)

    # 4. Collecter TOUTES les donnÃ©es en une seule passe (avec filtrage)
    all_rows: List[List[Optional[str]]] = []
    filtered_count = 0

    for row_match in RE_ROW.finditer(sheet_xml):
        row_num = int(row_match.group(1))
        if row_num <= 2:
            continue

        row_data = [None] * len(col_names)
        row_content = row_match.group(2)

        for cell_match in RE_CELL.finditer(row_content):
            col_letter = cell_match.group(1)

            if col_letter not in letter_to_idx:
                continue

            col_idx = letter_to_idx[col_letter]
            attrs = cell_match.group(2) or ""
            v_val = cell_match.group(3)

            if v_val is None:
                continue

            if RE_CELL_TYPE_S.search(attrs) and shared_strings:
                try:
                    idx = int(v_val)
                    if idx < len(shared_strings):
                        row_data[col_idx] = clean_value(shared_strings[idx])
                except ValueError:
                    pass
            else:
                row_data[col_idx] = clean_value(v_val)

        # FILTRAGE DES SOURCES EXCLUES + Remplacement des sources vides par "Natif"
        if source_col_idx is not None:
            source_value = row_data[source_col_idx]
            # Filtrer les sources exclues selon le type de feuille
            if sheet_info.name == 'Chats':
                excluded_set = EXCLUDED_CHAT_SOURCES
            elif sheet_info.name == 'Contacts':
                excluded_set = EXCLUDED_CONTACT_SOURCES
            else:
                excluded_set = EXCLUDED_SOURCES
            # Pour les fichiers SIM, ne pas filtrer la source 'SIM'
            if is_sim_file and source_value == 'SIM':
                pass  # Garder les contacts SIM pour fichiers carte SIM
            elif source_value and source_value in excluded_set:
                filtered_count += 1
                continue  # Skip cette ligne
            # Remplacer les sources vides par "Natif"
            if not source_value or source_value.strip() == "":
                row_data[source_col_idx] = "Natif"

        all_rows.append(row_data)

    if not all_rows:
        return 0, filtered_count

    # 5. CrÃ©er le DataFrame et Ã©crire en Parquet (une seule opÃ©ration)
    # Forcer toutes les colonnes en String pour éviter les conversions datetime automatiques
    data_dict = {col_names[i]: [row[i] for row in all_rows] for i in range(len(col_names))}
    schema = {col: pl.Utf8 for col in col_names}
    df = pl.DataFrame(data_dict, schema=schema)
    df.write_parquet(output_path, compression="zstd", compression_level=1)

    return len(all_rows), filtered_count


def extract_device_info_fast(zf: zipfile.ZipFile, shared_strings: List[str]) -> Dict[str, str]:
    """Extrait Device Info avec regex - version rapide"""
    device_data = {}

    # Trouver la feuille Device Info
    workbook_xml = zf.read('xl/workbook.xml').decode('utf-8')

    match = re.search(r'<sheet[^>]*name="Device Info"[^>]*r:id="(rId\d+)"', workbook_xml)
    if not match:
        match = re.search(r'<sheet[^>]*r:id="(rId\d+)"[^>]*name="Device Info"', workbook_xml)
    if not match:
        return device_data

    rel_id = match.group(1)

    # Trouver le fichier
    try:
        rels_xml = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    except:
        return device_data

    match = re.search(rf'<Relationship[^>]*Id="{rel_id}"[^>]*Target="([^"]+)"', rels_xml)
    if not match:
        return device_data

    target = match.group(1)
    sheet_path = 'xl/' + target if not target.startswith('xl/') else target

    try:
        sheet_xml = zf.read(sheet_path).decode('utf-8')
    except:
        return device_data

    # Parser colonnes C et D, lignes 3+
    for row_match in RE_ROW.finditer(sheet_xml):
        row_num = int(row_match.group(1))
        if row_num < 3:
            continue
        if row_num > 100:
            break

        row_content = row_match.group(2)
        cells = {}

        for cell_match in re.finditer(r'<c\s+r="([CD])\d+"([^>]*)>(?:<v>([^<]*)</v>)?</c>', row_content):
            col = cell_match.group(1)
            attrs = cell_match.group(2) or ""
            v_val = cell_match.group(3)

            if v_val is None:
                continue

            if 't="s"' in attrs and shared_strings:
                try:
                    idx = int(v_val)
                    if idx < len(shared_strings):
                        cells[col] = shared_strings[idx].strip()
                except:
                    pass
            else:
                cells[col] = v_val.strip()

        if 'C' in cells and 'D' in cells and cells['C'] and cells['D']:
            device_data[cells['C']] = cells['D']

    return device_data



# TRAITEMENT PARALLÃˆLE


def process_sheet_worker(args: Tuple) -> Tuple[str, Path, int, int]:
    """
    Worker pour traiter une feuille en parallÃ¨le.
    Retourne (nom, path, lignes_conservÃ©es, lignes_filtrÃ©es)
    """
    zf_bytes, sheet_info, shared_strings, output_path, is_sim_file = args

    with zipfile.ZipFile(io.BytesIO(zf_bytes), 'r') as zf:
        row_count, filtered_count = parse_sheet_to_parquet_fast(zf, sheet_info, shared_strings, output_path, is_sim_file)

    return sheet_info.name, output_path, row_count, filtered_count


def process_excel_streaming(file_content: bytes) -> ImportResult:
    """
    Traite un fichier Excel avec optimisations maximales :
    - Regex compilÃ©s
    - Une seule passe par feuille
    - ParallÃ©lisation des feuilles
    """
    import_id = str(uuid.uuid4())
    import_path = DATA_DIR / import_id
    import_path.mkdir(parents=True, exist_ok=True)

    parquet_files: Dict[str, Path] = {}
    row_counts: Dict[str, int] = {}

    start_total = time.perf_counter()

    with zipfile.ZipFile(io.BytesIO(file_content), 'r') as zf:
        # 1. Shared strings (doit Ãªtre fait en premier)
        t0 = time.perf_counter()
        shared_strings = parse_shared_strings_fast(zf)
        print(f"[FAST] SharedStrings: {len(shared_strings)} en {(time.perf_counter()-t0)*1000:.0f}ms")

        # 2. Info des feuilles
        t0 = time.perf_counter()
        sheets_info = get_sheets_info_fast(zf)
        print(f"[FAST] Sheets info: {len(sheets_info)} feuilles en {(time.perf_counter()-t0)*1000:.0f}ms")
        print(f"[FAST] Liste des feuilles Excel: {list(sheets_info.keys())}")

        # 3. Device Info (petit, fait sÃ©quentiellement)
        t0 = time.perf_counter()
        device_info = extract_device_info_fast(zf, shared_strings)
        print(f"[FAST] Device Info: {len(device_info)} valeurs en {(time.perf_counter()-t0)*1000:.0f}ms")

        # 4. Traitement des feuilles EN PARALLÃˆLE
        t0 = time.perf_counter()

        # Détecter si c'est un fichier SIM (présence de "SIM Data" dans les feuilles)
        is_sim_file = "SIM Data" in sheets_info
        if is_sim_file:
            print("[FAST] Fichier carte SIM détecté - source 'SIM' non filtrée")

        # PrÃ©parer les tÃ¢ches
        tasks = []
        for sheet_name, sheet_info in sheets_info.items():
            if sheet_info.file_path:
                output_path = import_path / f"{safe_filename(sheet_name)}.parquet"
                tasks.append((file_content, sheet_info, shared_strings, output_path, is_sim_file))

        # ExÃ©cution parallÃ¨le
        total_filtered = 0
        skipped_sheets = []  # Feuilles ignorees (0 lignes)
        error_sheets = []    # Feuilles avec erreur

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(process_sheet_worker, task): task[1].name for task in tasks}

            for future in as_completed(futures):
                sheet_name = futures[future]
                try:
                    name, path, count, filtered = future.result()
                    if count > 0:
                        parquet_files[name] = path
                        row_counts[name] = count
                        if filtered > 0:
                            print(f"[FAST] {name}: {count} lignes ({filtered} sources exclues filtrÃ©es)")
                            total_filtered += filtered
                        else:
                            print(f"[FAST] {name}: {count} lignes")
                    else:
                        skipped_sheets.append(name)
                except Exception as e:
                    print(f"[FAST ERROR] {sheet_name}: {e}")
                    error_sheets.append(sheet_name)

        # Afficher les feuilles ignorees
        if skipped_sheets:
            print(f"[FAST] Feuilles ignorees (0 lignes): {', '.join(skipped_sheets)}")
        if error_sheets:
            print(f"[FAST] Feuilles en erreur: {', '.join(error_sheets)}")

        print(f"[FAST] Feuilles traitÃ©es en {(time.perf_counter()-t0)*1000:.0f}ms")
        if total_filtered > 0:
            print(f"[FAST] Total lignes filtrÃ©es (sources exclues): {total_filtered}")

    total_time = time.perf_counter() - start_total
    print(f"[FAST] TOTAL: {total_time:.2f}s pour {len(parquet_files)} feuilles")

    return ImportResult(
        import_id=import_id,
        import_path=import_path,
        sheets=list(parquet_files.keys()),
        parquet_files=parquet_files,
        device_info=device_info,
        row_counts=row_counts
    )


# ANALYSES POLARS LAZY MODE


class LazyAnalyzer:
    """Analyseur utilisant exclusivement Polars lazy mode (scan_parquet)"""

    def __init__(self, import_path: Path):
        self.import_path = import_path
        self._schema_cache: Dict[str, List[str]] = {}

    def _get_parquet_path(self, sheet_name: str) -> Optional[Path]:
        """Trouve le fichier parquet d'une feuille"""
        path = self.import_path / f"{safe_filename(sheet_name)}.parquet"
        return path if path.exists() else None

    def _scan(self, sheet_name: str) -> Optional[pl.LazyFrame]:
        """Retourne un LazyFrame pour une feuille"""
        path = self._get_parquet_path(sheet_name)
        if path:
            return pl.scan_parquet(path)
        return None

    def get_columns(self, sheet_name: str) -> List[str]:
        """Liste les colonnes d'une feuille (avec cache)"""
        if sheet_name in self._schema_cache:
            return self._schema_cache[sheet_name]

        lf = self._scan(sheet_name)
        if lf is None:
            return []

        cols = list(lf.collect_schema().keys())
        self._schema_cache[sheet_name] = cols
        return cols

    def get_non_empty_columns(self, sheet_name: str, source_filter: Optional[str] = None) -> List[str]:
        """
        Liste les colonnes d'une feuille qui contiennent au moins une valeur non vide.
        Exclut les colonnes où toutes les valeurs sont null, vides, ou "N/C".
        """
        lf = self._scan(sheet_name)
        if lf is None:
            return []

        all_cols = self.get_columns(sheet_name)
        if not all_cols:
            return []

        # Appliquer le filtre source si demandé
        if source_filter and "Source" in all_cols:
            lf = lf.filter(pl.col("Source") == source_filter)

        # Collecter les données pour vérifier les colonnes non vides
        df = lf.collect()
        if len(df) == 0:
            return []

        non_empty_cols = []
        for col in all_cols:
            # Vérifier si la colonne a au moins une valeur non vide et non "N/C"
            col_data = df[col]
            has_valid_value = False
            for val in col_data:
                if val is not None:
                    str_val = str(val).strip()
                    if str_val and str_val.lower() != "n/c" and str_val != "":
                        has_valid_value = True
                        break
            if has_valid_value:
                non_empty_cols.append(col)

        return non_empty_cols

    def count_rows(self, sheet_name: str) -> int:
        """Compte les lignes (lazy - optimisÃ©)"""
        lf = self._scan(sheet_name)
        if lf is None:
            return 0
        return lf.select(pl.len()).collect().item()

    def get_sim_msisdn(self) -> Optional[str]:
        """
        Extrait le MSISDN depuis la feuille Sim Data.
        Cherche dans la colonne Name une entrée contenant "MSISDN" et retourne la Value correspondante.
        """
        lf = self._scan("SIM Data")
        if lf is None:
            # Essayer avec "Sim Data" (casse différente)
            lf = self._scan("Sim Data")
        if lf is None:
            return None

        cols = self.get_columns("SIM Data") or self.get_columns("Sim Data") or []
        if "Name" not in cols or "Value" not in cols:
            return None

        try:
            # Filtrer les lignes où Name contient "MSISDN"
            df = lf.filter(
                pl.col("Name").is_not_null() &
                pl.col("Name").str.to_lowercase().str.contains("msisdn")
            ).select(["Name", "Value"]).collect()

            if df.height > 0:
                msisdn = df.row(0)[1]  # Première valeur trouvée
                if msisdn:
                    return str(msisdn).strip()
        except Exception as e:
            print(f"[SIM Data] Erreur extraction MSISDN: {e}")

        return None

    def count_rows_filtered(self, sheet_name: str, source_filter: Optional[str] = None) -> int:
        """Compte les lignes avec filtre source optionnel (lazy)"""
        lf = self._scan(sheet_name)
        if lf is None:
            return 0

        if source_filter:
            cols = self.get_columns(sheet_name)
            if "Source" in cols:
                lf = lf.filter(pl.col("Source") == source_filter)

        return lf.select(pl.len()).collect().item()

    def count_media_dcim(self, sheet_name: str) -> int:
        """
        Compte les médias (Images, Videos, Audio) dont le chemin contient 'DCIM'.
        Vérifie d'abord la colonne 'Path', sinon cherche dans 'Meta Data' après 'File path:'.
        """
        lf = self._scan(sheet_name)
        if lf is None:
            return 0

        cols = self.get_columns(sheet_name)

        # Construire le filtre DCIM
        # Pattern pour Meta Data: "File path:...DCIM" (format standard)
        meta_dcim_pattern = r"File path:[^\n]*DCIM"

        if "Path" in cols and "Meta Data" in cols:
            # Cas 1: Path contient DCIM
            # Cas 2: Path vide/null mais Meta Data contient "File path:...DCIM"
            dcim_filter = (
                (pl.col("Path").is_not_null() & (pl.col("Path") != "") & pl.col("Path").str.contains("DCIM")) |
                (
                    (pl.col("Path").is_null() | (pl.col("Path") == "")) &
                    pl.col("Meta Data").is_not_null() &
                    pl.col("Meta Data").str.contains(meta_dcim_pattern)
                )
            )
        elif "Path" in cols:
            # Seulement colonne Path
            dcim_filter = pl.col("Path").is_not_null() & (pl.col("Path") != "") & pl.col("Path").str.contains("DCIM")
        elif "Meta Data" in cols:
            # Seulement colonne Meta Data
            dcim_filter = pl.col("Meta Data").is_not_null() & pl.col("Meta Data").str.contains(meta_dcim_pattern)
        else:
            # Aucune colonne pertinente, retourner 0
            return 0

        return lf.filter(dcim_filter).select(pl.len()).collect().item()

    def get_installed_app_names(self) -> List[str]:
        """
        Récupère la liste des noms d'applications installées depuis la feuille 'Installed Applications'.
        Exclut les applications système (com.apple, com.android, com.google).
        """
        lf = self._scan("Installed Applications")
        if lf is None:
            return []

        cols = self.get_columns("Installed Applications")
        if "Name" not in cols:
            return []

        # Récupérer les noms non vides
        names_df = (
            lf
            .select(pl.col("Name"))
            .filter(pl.col("Name").is_not_null() & (pl.col("Name") != ""))
            .collect()
        )

        if names_df.is_empty():
            return []

        # Convertir en liste et filtrer les apps système
        all_names = names_df["Name"].to_list()
        system_prefixes = ("com.apple", "com.android", "com.google")
        filtered_names = [
            name for name in all_names
            if name and not any(name.lower().startswith(prefix) for prefix in system_prefixes)
        ]

        return filtered_names

    def count_audio_with_installed_apps(self) -> int:
        """
        Compte les audios dont le Path contient une application installée.
        """
        lf = self._scan("Audio")
        if lf is None:
            return 0

        cols = self.get_columns("Audio")
        if "Path" not in cols:
            return 0

        # Récupérer les noms des applications installées
        app_names = self.get_installed_app_names()
        if not app_names:
            return 0

        # Construire un pattern regex pour matcher n'importe quelle app
        # Échapper les caractères spéciaux regex dans les noms d'apps
        import re
        escaped_names = [re.escape(name) for name in app_names]
        pattern = "|".join(escaped_names)

        # Filtrer les audios dont le Path contient une app installée
        audio_filter = (
            pl.col("Path").is_not_null() &
            (pl.col("Path") != "") &
            pl.col("Path").str.contains(pattern)
        )

        return lf.filter(audio_filter).select(pl.len()).collect().item()

    def get_contacts_stats(self) -> Tuple[int, Dict[str, int]]:
        """
        Stats contacts par source.
        NOTE: Les sources exclues sont dÃ©jÃ  filtrÃ©es Ã  la conversion Parquet.
        """
        lf = self._scan("Contacts")
        if lf is None:
            return 0, {}

        cols = self.get_columns("Contacts")
        if "Source" not in cols:
            total = self.count_rows("Contacts")
            return total, {"Natif": total}

        # Pas de filtre - les sources exclues sont dÃ©jÃ  absentes du Parquet
        result = (
            lf
            .with_columns(pl.col("Source").fill_null("Natif"))
            .group_by("Source")
            .agg(pl.len().alias("count"))
            .sort("count", descending=True)
            .collect()
        )

        counts_dict = {str(row["Source"]): row["count"] for row in result.to_dicts()}
        return sum(counts_dict.values()), counts_dict

    def _build_contacts_user_id_lookup(self, source: str = None) -> Dict[str, Dict[str, str]]:
        """
        Construit un dictionnaire de lookup pour trouver les infos depuis la feuille Contacts.

        Pour Snapchat/Instagram, depuis colonne Entries:
        - Colonne Name → pseudonyme (nom enregistré dans le carnet d'adresses)
        - User ID-Username: xxx → nom_utilisateur (username sur l'app)
        - User ID-User ID: xxx (Snap) ou User ID-Instagram Id: xxx (Insta) → identifiant_utilisateur

        Pour Signal, depuis colonne Entries:
        - Colonne Name → pseudonyme (nom enregistré dans le carnet d'adresses)
        - Phone-mobile: xxx → identifiant_utilisateur
        - Identifier: xxx → identifiant_utilisateur (si Phone-mobile absent)

        Retourne: {clé_recherche: {"pseudonyme": ..., "nom_utilisateur": ..., "identifiant_utilisateur": ...}}
        """
        lookup = {}

        lf = self._scan("Contacts")
        if lf is None:
            return lookup

        cols = self.get_columns("Contacts")
        source_lower = source.lower() if source else ""
        is_signal = source_lower == "signal"

        # Filtrer par source si spécifié
        if source and "Source" in cols:
            lf = lf.filter(pl.col("Source") == source)

        df = lf.collect()

        for row in df.to_dicts():
            name = str(row.get("Name", "") or "").strip()  # → pseudonyme (carnet d'adresses)
            entries = str(row.get("Entries", "") or "").strip()
            identifier_col = str(row.get("Identifier", "") or "").strip()  # Colonne Identifier si présente

            username_app = "-"  # User ID-Username = nom d'utilisateur sur l'app
            identifiant_utilisateur = ""

            if is_signal:
                # Pour Signal: Name → pseudonyme (carnet d'adresses), Phone-mobile: ou Identifier → identifiant_utilisateur
                # Chercher Phone-mobile dans Entries
                # Format possible: "Phone-Mobile: +34627578609 User ID-Profile Name: Zverev"
                # On doit s'arrêter avant "User ID-Profile Name:" si présent
                if entries:
                    match_phone = re.search(r'Phone-mobile:\s*([^\n\r]+?)(?=\s*(?:Phone-|User ID-)|\s*$)', entries, re.IGNORECASE)
                    if match_phone:
                        identifiant_utilisateur = match_phone.group(1).strip()

                    # Si pas de Phone-mobile, chercher Identifier dans Entries
                    if not identifiant_utilisateur:
                        match_id = re.search(r'Identifier:\s*([^\n\r]+?)(?=\s*[A-Z][a-z]*[-:]|\s*$)', entries, re.IGNORECASE)
                        if match_id:
                            identifiant_utilisateur = match_id.group(1).strip()

                # Utiliser la colonne Identifier si pas trouvé dans Entries
                if not identifiant_utilisateur and identifier_col:
                    identifiant_utilisateur = identifier_col

                # Pour Signal, pas de nom d'utilisateur (username_app reste "-")
                username_app = "-"

            else:
                # Pour Snapchat/Instagram et autres sources sociales
                if entries:
                    # Pattern pour détecter le prochain champ en gras
                    # Format: "Mot-:" ou "Mot Mot-:" ou "Mot-Mot:" ou "Mot Mot-Mot Mot:"
                    # Ex: Phone-:, User ID-:, User ID-Username:, User ID-User ID:
                    next_field_pattern = r'\s+[A-Za-z][A-Za-z0-9 ]*-[A-Za-z0-9 ]*:'

                    # Extraire UNIQUEMENT User ID-Username → nom_utilisateur (username sur l'app)
                    match_username = re.search(r'User ID-Username:\s*([^\n\r]+?)(?=' + next_field_pattern + r'|$)', entries, re.IGNORECASE)
                    if match_username:
                        username_app = match_username.group(1).strip()

                    # Extraire User ID-User ID → identifiant_utilisateur (pour Snapchat)
                    match_userid = re.search(r'User ID-User ID:\s*([^\n\r]+?)(?=' + next_field_pattern + r'|$)', entries, re.IGNORECASE)
                    if match_userid:
                        identifiant_utilisateur = match_userid.group(1).strip()

                    # Extraire User ID-Instagram Id → identifiant_utilisateur (pour Instagram)
                    if not identifiant_utilisateur:
                        match_insta = re.search(r'User ID-Instagram Id:\s*([^\n\r]+?)(?=' + next_field_pattern + r'|$)', entries, re.IGNORECASE)
                        if match_insta:
                            identifiant_utilisateur = match_insta.group(1).strip()

                    # Extraire "User ID-:" (champ sans sous-libelle) → identifiant_utilisateur
                    # Ex: "User ID-: 3939be6c-c8bc-4b6a-aa4b-bda762c92524" (TikTok et autres sources)
                    if not identifiant_utilisateur:
                        match_userid_dash = re.search(r'User ID-\s*:\s*([^\n\r]+?)(?=' + next_field_pattern + r'|$)', entries, re.IGNORECASE)
                        if match_userid_dash:
                            identifiant_utilisateur = match_userid_dash.group(1).strip()

                    # Pour TikTok et autres: chercher des patterns similaires
                    if not identifiant_utilisateur and username_app == "-":
                        # Chercher Username ou User-Id générique
                        match_generic_user = re.search(r'(?:Username|User[-_]?name):\s*([^\n\r]+?)(?=' + next_field_pattern + r'|$)', entries, re.IGNORECASE)
                        if match_generic_user:
                            username_app = match_generic_user.group(1).strip()

                        match_generic_id = re.search(r'(?:User[-_]?Id|Id):\s*([^\n\r]+?)(?=' + next_field_pattern + r'|$)', entries, re.IGNORECASE)
                        if match_generic_id:
                            identifiant_utilisateur = match_generic_id.group(1).strip()

            # Créer l'entrée de lookup avec toutes les infos
            # pseudonyme = name (carnet d'adresses), nom_utilisateur = username_app (username sur l'app)
            contact_info = {
                "pseudonyme": name if name else "-",
                "nom_utilisateur": username_app,
                "identifiant_utilisateur": identifiant_utilisateur
            }

            # Ajouter au lookup avec différentes clés de recherche
            # On peut chercher par name, username_app, ou identifiant
            if name:
                lookup[name.lower()] = contact_info
            if username_app and username_app != "-":
                lookup[username_app.lower()] = contact_info
            if identifiant_utilisateur:
                lookup[identifiant_utilisateur.lower()] = contact_info

        return lookup

    def get_calls_stats(self) -> Tuple[int, Dict[str, int]]:
        """Stats appels par source"""
        lf = self._scan("Call Log")
        if lf is None:
            return 0, {}

        cols = self.get_columns("Call Log")
        source_col = next((c for c in ["Source", "Application", "App"] if c in cols), None)

        if source_col:
            result = (
                lf
                .with_columns(pl.col(source_col).fill_null("Natif"))
                .group_by(source_col)
                .agg(pl.len().alias("count"))
                .sort("count", descending=True)
                .collect()
            )
            counts_dict = {str(row[source_col]): row["count"] for row in result.to_dicts()}
            return sum(counts_dict.values()), counts_dict
        else:
            total = self.count_rows("Call Log")
            return total, {"Natif": total}

    def get_chats_stats(self) -> Tuple[int, int, Dict[str, int]]:
        """Stats chats: (nb_messages, nb_conversations, {source: count})"""
        lf = self._scan("Chats")
        if lf is None:
            return 0, 0, {}

        cols = self.get_columns("Chats")

        # Filtrer les lignes où Instant Message # est vide
        im_col = "Instant Message #" if "Instant Message #" in cols else None
        if im_col:
            lf = lf.filter(pl.col(im_col).is_not_null() & (pl.col(im_col).cast(pl.Utf8).str.strip_chars() != ""))

        nb_messages = lf.select(pl.len()).collect().item()

        # Nombre de conversations
        nb_conversations = 0
        conv_col = next((c for c in ["Chat #", "Chat", "Conversation"] if c in cols), None)
        if conv_col:
            try:
                result = lf.select(pl.col(conv_col).max()).collect()
                val = result.item()
                nb_conversations = int(val) if val is not None else 0
            except:
                nb_conversations = lf.select(pl.col(conv_col).n_unique()).collect().item()

        # Stats par source
        source_col = next((c for c in ["Source", "Application", "App", "Platform"] if c in cols), None)

        if source_col:
            result = (
                lf
                .with_columns(pl.col(source_col).fill_null("Natif"))
                .group_by(source_col)
                .agg(pl.len().alias("count"))
                .sort("count", descending=True)
                .collect()
            )
            counts_dict = {str(row[source_col]): row["count"] for row in result.to_dicts()}
        else:
            counts_dict = {"Messages": nb_messages}

        return nb_messages, nb_conversations, counts_dict

    def get_accounts_stats(self) -> Tuple[int, Dict[str, int]]:
        """Stats comptes utilisateur par source"""
        lf = self._scan("User Accounts")
        if lf is None:
            return 0, {}

        cols = self.get_columns("User Accounts")
        source_col = next((c for c in ["Source", "Application", "App"] if c in cols), None)

        if source_col:
            result = (
                lf
                .with_columns(pl.col(source_col).fill_null("Natif"))
                .group_by(source_col)
                .agg(pl.len().alias("count"))
                .sort("count", descending=True)
                .collect()
            )
            counts_dict = {str(row[source_col]): row["count"] for row in result.to_dicts()}
            return sum(counts_dict.values()), counts_dict
        else:
            total = self.count_rows("User Accounts")
            return total, {"Comptes": total}

    def get_sheet_data(
        self,
        sheet_name: str,
        columns: List[str],
        max_rows: int = 100,
        source_filter: Optional[str] = None
    ) -> List[Dict[str, str]]:
        """RÃ©cupÃ¨re des donnÃ©es d'une feuille (lazy + collect limitÃ©)"""
        lf = self._scan(sheet_name)
        if lf is None:
            return []

        all_sheet_cols = self.get_columns(sheet_name)
        available_cols = [c for c in columns if c in all_sheet_cols]
        if not available_cols:
            return []

        # Appliquer le filtre source AVANT de sÃ©lectionner les colonnes
        # Le filtre fonctionne mÃªme si Source n'est pas dans les colonnes affichÃ©es
        if source_filter and "Source" in all_sheet_cols:
            lf = lf.filter(pl.col("Source") == source_filter)

        # Pour "Installed Applications", filtrer les lignes où Name est vide
        if "installed" in sheet_name.lower() and "application" in sheet_name.lower():
            if "Name" in all_sheet_cols:
                lf = lf.filter(
                    pl.col("Name").is_not_null() &
                    (pl.col("Name").cast(pl.Utf8).str.strip_chars() != "")
                )

        # SÃ©lectionner uniquement les colonnes demandÃ©es pour l'affichage
        query = lf.select(available_cols)
        result = query.head(max_rows).collect()

        def format_value(v):
            """Formate une valeur pour l'affichage, avec dates en format lisible"""
            if v is None:
                return ""
            # Gérer les types datetime de Polars/Python
            from datetime import datetime, date, time as dt_time, timedelta
            if hasattr(v, 'strftime'):  # datetime, date, time
                if isinstance(v, datetime):
                    return v.strftime("%d/%m/%Y %H:%M:%S")
                elif isinstance(v, date):
                    return v.strftime("%d/%m/%Y")
                elif isinstance(v, dt_time):
                    return v.strftime("%H:%M:%S")

            # Convertir les numéros de série Excel en dates lisibles
            # Les numéros Excel sont entre ~1 (1900) et ~50000 (2037)
            str_v = str(v)
            try:
                # Vérifier si c'est un numéro de série Excel (ex: 45848.840925925928)
                if '.' in str_v and str_v.replace('.', '').replace('-', '').isdigit():
                    num = float(str_v)
                    # Plage valide pour les dates Excel (1900-2100 environ)
                    if 1 < num < 100000:
                        # Convertir le numéro de série Excel en date
                        # Excel commence le 1er janvier 1900, mais il y a un bug (1900 n'est pas bissextile)
                        excel_epoch = datetime(1899, 12, 30)
                        days = int(num)
                        fraction = num - days
                        dt = excel_epoch + timedelta(days=days)
                        # Ajouter la partie fractionnaire (temps)
                        if fraction > 0:
                            seconds = int(fraction * 86400)  # 86400 = secondes par jour
                            dt = dt + timedelta(seconds=seconds)
                            return dt.strftime("%d/%m/%Y %H:%M:%S")
                        else:
                            return dt.strftime("%d/%m/%Y")
            except (ValueError, OverflowError):
                pass

            return str_v

        def translate_value(column_name: str, value: str) -> str:
            """Traduit une valeur si la colonne est dans COLUMNS_TO_TRANSLATE"""
            if column_name not in COLUMNS_TO_TRANSLATE:
                return value
            # Récupérer le dictionnaire de traduction pour cette colonne
            translations = VALUE_TRANSLATIONS.get(column_name, {})
            if not translations:
                return value
            # Traduire chaque partie si la valeur contient plusieurs statuts séparés par virgule ou point-virgule
            parts = [p.strip() for p in re.split(r'[,;]', value) if p.strip()]
            translated_parts = []
            for part in parts:
                # Chercher traduction exacte (insensible à la casse)
                translated = None
                for eng, fr in translations.items():
                    if eng.lower() == part.lower():
                        translated = fr
                        break
                translated_parts.append(translated if translated else part)
            return ", ".join(translated_parts) if len(translated_parts) > 1 else (translated_parts[0] if translated_parts else value)

        def clean_value(column_name: str, value: str) -> str:
            """Nettoie le HTML de la colonne Body pour les Emails"""
            # Nettoyer le HTML seulement pour la colonne Body de la feuille Emails
            if column_name == "Body" and "email" in sheet_name.lower():
                return strip_html_tags(value)
            return value

        return [
            {k: clean_value(k, translate_value(k, format_value(v))) for k, v in row.items()}
            for row in result.to_dicts()
        ]

    def _extract_contacts(self, parties: str, source_value: str) -> List[Dict[str, Any]]:
        """
        Extrait les contacts depuis la colonne Parties selon la logique CallLogProcessor.ts.
        Retourne une liste de dicts avec: Username/UUID/Phone, Name, Prefix
        """
        contacts = []
        if not parties:
            return contacts

        source_lower = source_value.lower() if source_value else ""
        lines = parties.split('\n')
        prefix_pattern = re.compile(r'^(From:|To:|General:)\s*(.+)$')

        for line in lines:
            trimmed_line = line.strip()
            if not trimmed_line:
                continue

            # Extraire le prÃ©fixe (From:, To:, General:)
            match = prefix_pattern.match(trimmed_line)
            prefix = None
            contact_info = ""

            if match:
                prefix = match.group(1)
                contact_info = match.group(2).strip()
            else:
                contact_info = trimmed_line

            # Certaines lignes contiennent plusieurs participants en chaîne: "... To: ... To: ..."
            # Spliter par "To:" et assigner le bon prefix à chaque segment
            contact_segments_raw = re.split(r'\s+To:\s+', contact_info)
            contact_segments = []
            for idx, seg in enumerate(contact_segments_raw):
                seg = seg.strip()
                if not seg:
                    continue
                # Le premier segment garde le prefix original, les suivants ont "To:"
                seg_prefix = prefix if idx == 0 else "To:"

                # Nettoyer To: ou From: au début si présent (quand le préfixe précédent est vide)
                # Ex: "From: To: princevbs..." -> contact_info = "To: princevbs..."
                if seg.startswith('To:'):
                    seg = seg[3:].strip()
                    seg_prefix = "To:"  # Forcer le prefix
                elif seg.startswith('From:'):
                    seg = seg[5:].strip()
                    seg_prefix = "From:"

                if not seg:
                    continue

                contact_segments.append((seg, seg_prefix))

            for segment, current_prefix in contact_segments:
                contact = None

                # Extraction selon le type de source (comme CallLogProcessor.ts)
                if source_lower == 'snapchat':
                    # Pattern: username nom
                    snap_match = re.match(r'([^\s]+)\s+(.*)', segment)
                    if snap_match:
                        # Nettoyer le nom: supprimer "To:" ou "From:" en fin de chaîne
                        name = snap_match.group(2).strip() or "Inconnu"
                        name = re.sub(r'\s*(To:|From:)\s*$', '', name).strip() or "Inconnu"
                        contact = {
                            "Username": snap_match.group(1).strip() or "Inconnu",
                            "Name": name,
                            "Prefix": current_prefix
                        }

                elif source_lower == 'whatsapp' or source_lower == 'whatsapp business':
                    # Pattern WhatsApp: phone (doit se terminer par un chiffre) + suffixe optionnel @s.whatsapp.net + nom
                    whatsapp_match = re.match(r'(\+?[\d\s\-\.]+\d)(?:@s\.whatsapp\.net)?(?:\s+(.+))?', segment)
                    if whatsapp_match:
                        # Normaliser le numéro pour éviter les variantes avec espaces/tirets
                        phone = re.sub(r'[\s\-\.]', '', whatsapp_match.group(1))
                        # Nettoyer le nom: supprimer "To:" ou "From:" en fin de chaîne
                        name = whatsapp_match.group(2) or "Inconnu"
                        name = re.sub(r'\s*(To:|From:)\s*$', '', name).strip() or "Inconnu"
                        contact = {
                            "Phone": phone,
                            "Name": name,
                            "Prefix": current_prefix
                        }

                elif source_lower == 'signal':
                    # Pattern Signal: UUID ou phone suivi de "User ID-Profile Name: xxx" ou juste un nom
                    # Ex: "+33603444467 User ID-Profile Name: morgane17102011"
                    # Ex: "FCF060ED-231F-4DD5-8390-01B519F22801"
                    # Ex: "0FD9CF43-C376-4D8E-8E58-108F33796F4 coub" (UUID de 35 caractères)
                    # UUID peut avoir 20-50 caractères (avec ou sans tirets, parfois tronqué)
                    signal_match = re.match(r'([A-Fa-f0-9\-]{20,50}|\+?\d[\d\s\-\.]{8,}|\d{5,20})\s*(.*)', segment)
                    if signal_match:
                        identifier = signal_match.group(1).strip()
                        raw_name = signal_match.group(2).strip()

                        # Parser "User ID-Profile Name: xxx" pour extraire le vrai nom
                        profile_name_match = re.search(r'User ID-Profile Name:\s*(.+?)(?:\s*User ID-|$)', raw_name, re.IGNORECASE)
                        if profile_name_match:
                            name = profile_name_match.group(1).strip()
                        else:
                            # Pas de format "User ID-Profile Name:", utiliser raw_name comme nom
                            name = raw_name if raw_name else "Inconnu"

                        # Nettoyer le nom si vide ou juste un point
                        if not name or name == ".":
                            name = "Inconnu"

                        contact = {"Name": name, "Prefix": current_prefix}
                        # UUID (20-50 caractères hex avec tirets) ou Phone
                        if re.match(r'^[A-Fa-f0-9\-]{20,50}$', identifier):
                            contact["UUID"] = identifier
                        else:
                            # Nettoyer le numéro de téléphone
                            contact["Phone"] = re.sub(r'[\s\-\.]', '', identifier)
                    else:
                        # Fallback utile pour les identifiants Signal non strictement UUID/phone
                        generic_match = re.match(r'([^\s]+)\s+(.*)', segment)
                        if generic_match:
                            identifier = generic_match.group(1).strip()
                            raw_name = generic_match.group(2).strip()
                            # Parser aussi "User ID-Profile Name:" dans le fallback
                            profile_name_match = re.search(r'User ID-Profile Name:\s*(.+?)(?:\s*User ID-|$)', raw_name, re.IGNORECASE)
                            if profile_name_match:
                                name = profile_name_match.group(1).strip()
                            else:
                                name = raw_name if raw_name else "Inconnu"

                            if not name or name == ".":
                                name = "Inconnu"

                            contact = {
                                "Name": name,
                                "Prefix": current_prefix
                            }
                            # Déterminer si c'est un UUID ou autre
                            if re.match(r'^[A-Fa-f0-9\-]{20,}$', identifier):
                                contact["UUID"] = identifier
                            else:
                                contact["Username"] = identifier or "Inconnu"

                else:
                    # Extraction générique: essayer phone d'abord, puis username
                    phone_match = re.match(r'(\+?\d[\d\s\-\.]{8,}|\d{5,20})\s*(.*)', segment)
                    if phone_match:
                        # Nettoyage du phone ICI comme dans le TS
                        phone = re.sub(r'[\s\-\.]', '', phone_match.group(1))
                        contact = {
                            "Phone": phone,
                            "Name": phone_match.group(2).strip() or "Inconnu",
                            "Prefix": current_prefix
                        }
                    else:
                        # Essayer username nom
                        snap_match = re.match(r'([^\s]+)\s+(.*)', segment)
                        if snap_match:
                            contact = {
                                "Username": snap_match.group(1).strip() or "Inconnu",
                                "Name": snap_match.group(2).strip() or "Inconnu",
                                "Prefix": current_prefix
                            }

                if contact:
                    contacts.append(contact)

        return contacts

    def _extract_owner_from_account(self, account_value: str, source: str) -> Optional[str]:
        """
        Extrait et normalise l'identifiant du propriétaire depuis la colonne Account.
        Ex: "33660048980@s.whatsapp.net" -> "33660048980"
        Ex: "yapasrienla7575" -> "yapasrienla7575"
        """
        if not account_value:
            return None

        account = str(account_value).strip()
        if not account:
            return None

        # Pour WhatsApp, extraire le numéro avant @s.whatsapp.net
        if "@s.whatsapp.net" in account:
            account = account.split("@")[0]

        # Normaliser (supprimer espaces, tirets, points, +)
        clean_id = re.sub(r'[\s\-\.+]', '', account)
        return clean_id if clean_id else None

    def _get_owner_usernames_from_user_accounts(self) -> set:
        """
        Récupère les usernames du propriétaire depuis la feuille User Accounts.
        Retourne un set de clean_id (nettoyés et en minuscules).
        """
        owner_usernames: set = set()

        lf = self._scan("User Accounts")
        if lf is None:
            return owner_usernames

        cols = self.get_columns("User Accounts")
        if "Username" not in cols:
            return owner_usernames

        try:
            df = lf.select("Username").collect()
            for row in df.to_dicts():
                username = str(row.get("Username", "") or "").strip()
                if username:
                    # Nettoyer le username
                    clean_username = re.sub(r'@s\.whatsapp\.net$', '', username, flags=re.IGNORECASE)
                    clean_username = re.sub(r'[\s\-\.\+]', '', clean_username).lower()
                    if clean_username:
                        owner_usernames.add(clean_username)
        except Exception:
            pass

        return owner_usernames

    def _identify_owner(self, rows: List[Dict[str, Any]], parties_col: str, source: str) -> set:
        """
        Identifie le(s) propriétaire(s) du téléphone:
        1. D'abord chercher dans la feuille User Accounts (colonne Username)
        2. Sinon, chercher dans la colonne Account
        3. Sinon, ceux dont le nom/identifiant contient "(owner)"
        Retourne un set de clean_id des propriétaires.
        """
        owners: set = set()

        # 1. Chercher dans User Accounts (même logique que les bulles vertes)
        user_account_owners = self._get_owner_usernames_from_user_accounts()
        if user_account_owners:
            # Parcourir les contacts pour trouver ceux qui matchent les owner usernames
            for row in rows:
                parties = str(row.get(parties_col, "") or "")
                extracted_contacts = self._extract_contacts(parties, source)

                for contact in extracted_contacts:
                    identifier = contact.get("Username") or contact.get("UUID") or contact.get("Phone")
                    if not identifier:
                        continue

                    clean_id = re.sub(r'[\s\-\.\+]', '', identifier).lower()

                    # Vérifier si ce contact est un owner (depuis User Accounts)
                    for owner_id in user_account_owners:
                        if owner_id in clean_id or clean_id in owner_id:
                            owners.add(re.sub(r'[\s\-\.+]', '', identifier))
                            break

            if owners:
                return owners

        # 2. Chercher dans la colonne Account
        for row in rows:
            account_value = row.get("Account", "") or ""
            owner_id = self._extract_owner_from_account(account_value, source)
            if owner_id:
                owners.add(owner_id)

        if owners:
            return owners

        # 3. Chercher les marqueurs "(owner)" dans les contacts
        owners_by_marker: set = set()

        for row in rows:
            parties = str(row.get(parties_col, "") or "")
            extracted_contacts = self._extract_contacts(parties, source)

            for contact in extracted_contacts:
                identifier = contact.get("Username") or contact.get("UUID") or contact.get("Phone")
                if not identifier:
                    continue

                clean_id = re.sub(r'[\s\-\.+]', '', identifier)
                name = contact.get("Name", "") or ""

                # Vérifier si "(owner)" est présent
                haystack = f"{identifier} {name}".lower()
                if "(owner)" in haystack:
                    owners_by_marker.add(clean_id)

        if owners_by_marker:
            return owners_by_marker

        return set()

    def _analyze_call_log(self, source: str) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, set]]:
        """
        Analyse le journal d'appels pour une source et retourne les stats agrégées par contact.
        Utilise la nouvelle logique pour déterminer émis/reçu:

        | Contact dans | Proprio dans | Direction | Résultat |
        |--------------|--------------|-----------|----------|
        | From         | To           | incoming  | Émis     |
        | From         | To           | outgoing  | Reçu     |
        | To           | From         | outgoing  | Reçu     |
        | To           | From         | incoming  | Émis     |
        | From         | vide         | incoming  | Émis     |
        | To           | vide         | outgoing  | Reçu     |

        Retourne: (contact_stats, id_to_names)
        """
        lf = self._scan("Call Log")
        if lf is None:
            return {}, {}

        cols = self.get_columns("Call Log")

        df = lf.collect()
        if len(df) == 0:
            return {}, {}

        # Identifier la colonne des parties/contacts
        parties_col = next((c for c in ["Parties", "Party", "Contact", "Number", "Phone"] if c in cols), None)
        if not parties_col:
            return {}, {}

        rows = self._preprocess_call_log_rows(df.to_dicts(), cols, parties_col, source)

        # Identifier les propriétaires (depuis Account, "(owner)", ou max appels)
        owner_ids = self._identify_owner(rows, parties_col, source)

        # Liste intermédiaire pour l'agrégation
        contacts_list = []
        id_to_names: Dict[str, set] = {}

        for row in rows:
            parties = str(row.get(parties_col, "") or "")
            duration = self._duration_to_seconds(str(row.get("Duration", "00:00:00") or "00:00:00"))
            direction = str(row.get("Direction", "") or "").strip().lower()

            # Nouvelles colonnes pour le tableau résumé
            video_call = str(row.get("Video call", "") or "").strip().lower() == "yes"
            status_missed = str(row.get("Status", "") or "").strip().lower() == "missed"
            deleted = str(row.get("Deleted", "") or "").strip().lower() in ("yes", "trash")

            # Extraire les contacts de la colonne Parties
            extracted_contacts = self._extract_contacts(parties, source)

            # Identifier où se trouve le propriétaire dans cet appel (From, To, ou absent)
            owner_prefix = None  # "From:", "To:", ou None (absent)
            for contact in extracted_contacts:
                identifier = contact.get("Username") or contact.get("UUID") or contact.get("Phone")
                if not identifier:
                    continue
                clean_id = re.sub(r'[\s\-\.+]', '', identifier)
                if clean_id in owner_ids:
                    owner_prefix = contact.get("Prefix")
                    break

            # Dédupliquer les contacts dans cet appel
            unique_contacts: Dict[str, Dict[str, Any]] = {}
            for contact in extracted_contacts:
                identifier = contact.get("Username") or contact.get("UUID") or contact.get("Phone")
                if not identifier:
                    continue

                clean_id = re.sub(r'[\s\-\.+]', '', identifier)

                # Exclure les propriétaires des statistiques
                if clean_id in owner_ids:
                    continue

                if clean_id in unique_contacts:
                    continue

                unique_contacts[clean_id] = {
                    "contact": contact,
                    "display_identifier": str(identifier).strip(),
                }

            # Pour chaque contact unique (hors propriétaire)
            for clean_id, contact_data in unique_contacts.items():
                contact = contact_data["contact"]
                display_identifier = contact_data["display_identifier"]
                contact_prefix = contact.get("Prefix") or ""
                appel_emis = 0
                appel_recu = 0

                # Appliquer la nouvelle logique émis/reçu
                # contact_prefix = où se trouve le contact (From: ou To:)
                # owner_prefix = où se trouve le propriétaire (From:, To:, ou None)
                # direction = incoming ou outgoing

                if contact_prefix == "From:" and owner_prefix == "To:":
                    if direction == "incoming":
                        appel_emis = 1
                    elif direction == "outgoing":
                        appel_recu = 1
                elif contact_prefix == "To:" and owner_prefix == "From:":
                    if direction == "outgoing":
                        appel_recu = 1
                    elif direction == "incoming":
                        appel_emis = 1
                elif contact_prefix == "From:" and owner_prefix is None:
                    if direction == "incoming":
                        appel_emis = 1
                elif contact_prefix == "To:" and owner_prefix is None:
                    if direction == "outgoing":
                        appel_recu = 1
                # Cas General: (appels de groupe) - même logique que le TS original
                elif contact_prefix == "General:":
                    if direction == "incoming":
                        appel_recu = 1
                    elif direction == "outgoing":
                        appel_emis = 1

                # Utiliser clean_id comme Identifier
                source_lower = source.lower() if source else ""
                if source_lower in {"whatsapp", "whatsapp business"} and clean_id.startswith("+"):
                    clean_id = clean_id[1:]

                # Normalisation des numéros de téléphone pour source Natif
                final_id = clean_id
                final_display = display_identifier
                if source_lower == "natif":
                    normalized_root, is_phone, is_french = self._normalize_phone_natif(clean_id)
                    if is_phone:
                        final_id = normalized_root
                        if is_french:
                            final_display = f"+33{normalized_root}"
                        else:
                            # Numéro étranger
                            final_display = f"+{normalized_root}"

                contacts_list.append({
                    "Identifier": final_id,
                    "DisplayIdentifier": final_display,
                    "Duration": duration,
                    "Call_Count": 1,
                    "Appel_emis": appel_emis,
                    "Appel_recu": appel_recu,
                    "Appel_video": 1 if video_call else 0,
                    "Appel_manque": 1 if status_missed else 0,
                    "Appel_supprime": 1 if deleted else 0
                })

                # Collecter les noms pour cet identifiant (utiliser final_id pour regrouper)
                clean_name = re.sub(r'_x000d_', '', contact.get("Name", ""), flags=re.IGNORECASE).strip()
                # Supprimer les préfixes General:, From:, To: et tout ce qui suit
                clean_name = re.split(r'\s*(?:General:|From:|To:)\s*', clean_name)[0].strip()
                if clean_name and clean_name != "Inconnu":
                    if final_id not in id_to_names:
                        id_to_names[final_id] = set()
                    id_to_names[final_id].add(clean_name)

        # AgrÃ©gation par identifiant (clean_id)
        aggregated: Dict[str, Dict[str, Any]] = {}

        for item in contacts_list:
            identifier = item["Identifier"]
            if identifier not in aggregated:
                aggregated[identifier] = {
                    "Identifier": identifier,
                    "DisplayIdentifier": item.get("DisplayIdentifier", identifier),
                    "Name": "",
                    "Nombre_appels": 0,
                    "Appel_emis": 0,
                    "Appel_recu": 0,
                    "Duree_totale_sec": 0,
                    "Appel_video": 0,
                    "Appel_manque": 0,
                    "Appel_supprime": 0
                }

            aggregated[identifier]["Nombre_appels"] += item["Call_Count"]
            aggregated[identifier]["Appel_emis"] += item["Appel_emis"]
            aggregated[identifier]["Appel_recu"] += item["Appel_recu"]
            aggregated[identifier]["Duree_totale_sec"] += item["Duration"]
            aggregated[identifier]["Appel_video"] += item["Appel_video"]
            aggregated[identifier]["Appel_manque"] += item["Appel_manque"]
            aggregated[identifier]["Appel_supprime"] += item["Appel_supprime"]

        # Ajouter les noms
        for identifier in aggregated:
            if identifier in id_to_names:
                aggregated[identifier]["Name"] = ", ".join(sorted(id_to_names[identifier]))

        return aggregated, id_to_names

    def _filter_owner_and_remove_max(self, contacts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Filtre les contacts contenant '(owner)' dans leur identifiant/nom.
        Note: Le propriétaire principal est déjà exclu dans _analyze_call_log,
        mais cette méthode assure un filtrage supplémentaire pour les marqueurs explicites.
        """
        filtered: List[Dict[str, Any]] = []
        for c in contacts:
            identifier = str(c.get("Identifier", "") or "")
            display_identifier = str(c.get("DisplayIdentifier", "") or "")
            name = str(c.get("Name", "") or "")
            haystack = f"{identifier} {display_identifier} {name}".lower()
            if "(owner)" in haystack:
                continue
            filtered.append(c)

        return filtered

    def _filter_short_phone_numbers_for_natif(self, contacts: List[Dict[str, Any]], source: str) -> List[Dict[str, Any]]:
        """
        Pour la source Natif, filtre les numéros de téléphone avec moins de 6 chiffres
        car ce sont des appels système.
        """
        if source.lower() != "natif":
            return contacts

        filtered: List[Dict[str, Any]] = []
        for c in contacts:
            identifier = str(c.get("Identifier", "") or "")
            # Extraire uniquement les chiffres de l'identifiant
            digits_only = re.sub(r'[^\d]', '', identifier)
            # Si moins de 6 chiffres, c'est un appel système -> exclure
            if len(digits_only) < 6:
                continue
            filtered.append(c)

        return filtered

    def _display_identifier(self, identifier: str, source: str) -> str:
        source_lower = source.lower() if source else ""
        ident = str(identifier or "")
        if source_lower == "signal" and ident.startswith("SignalGroup:"):
            ident = ident[len("SignalGroup:"):]
        if source_lower in {"whatsapp", "whatsapp business"} and re.match(r"^\d+$", ident):
            return f"+{ident}"
        return ident

    def _normalize_phone_natif(self, identifier: str) -> Tuple[str, bool, bool]:
        """
        Normalise un numéro de téléphone pour la source Natif.
        Gère les numéros français et étrangers.

        Retourne: (racine_normalisée, est_numero_telephone, est_francais)
        Pour français: racine = 9 derniers chiffres (ex: 627172210)
        Pour étranger: racine = numéro complet avec indicatif
        """
        if not identifier:
            return identifier, False, False

        # Extraire uniquement les chiffres
        digits_only = re.sub(r'\D', '', identifier)

        if not digits_only:
            return identifier, False, False

        # Vérifier si c'est un numéro de téléphone (au moins 7 chiffres pour les pays avec numéros courts)
        if len(digits_only) < 7:
            return identifier, False, False

        # Exclure les numéros avec trop de zéros consécutifs (ex: +3300000000000)
        if re.search(r'0{5,}', digits_only):
            return identifier, False, False

        # Indicatifs de pays connus (pour validation)
        # Format: (préfixe, longueur_min_totale, longueur_max_totale)
        country_codes = {
            '33': (11, 12),   # France
            '32': (10, 11),   # Belgique
            '41': (11, 12),   # Suisse
            '352': (10, 12),  # Luxembourg
            '1': (11, 11),    # USA/Canada
            '44': (11, 13),   # UK
            '49': (11, 14),   # Allemagne
            '34': (11, 11),   # Espagne
            '39': (11, 13),   # Italie
            '351': (11, 12),  # Portugal
            '31': (11, 11),   # Pays-Bas
            '212': (11, 12),  # Maroc
            '213': (11, 12),  # Algérie
            '216': (11, 12),  # Tunisie
            '221': (11, 12),  # Sénégal
            '225': (11, 12),  # Côte d'Ivoire
            '237': (11, 12),  # Cameroun
            '243': (11, 12),  # RDC
            '224': (11, 12),  # Guinée
            '223': (11, 11),  # Mali
            '971': (11, 12),  # Émirats
            '966': (11, 12),  # Arabie Saoudite
            '90': (11, 12),   # Turquie
            '7': (11, 12),    # Russie
            '86': (12, 13),   # Chine
            '91': (12, 12),   # Inde
            '55': (12, 13),   # Brésil
            '81': (11, 12),   # Japon
            '82': (11, 12),   # Corée du Sud
            '61': (11, 11),   # Australie
            '27': (11, 11),   # Afrique du Sud
            '20': (11, 12),   # Égypte
            '48': (11, 11),   # Pologne
            '40': (11, 11),   # Roumanie
            '380': (12, 12),  # Ukraine
            '386': (10, 11),  # Slovénie
            '385': (10, 12),  # Croatie
            '381': (11, 12),  # Serbie
        }

        # Détecter si c'est un numéro français
        is_french = False

        # Numéro français: commence par 33 ou 0
        if digits_only.startswith("33") and len(digits_only) >= 11:
            # +33627172210 -> 627172210
            root = digits_only[2:]
            is_french = True
        elif digits_only.startswith("0") and len(digits_only) == 10:
            # 0627172210 -> 627172210
            root = digits_only[1:]
            is_french = True
        elif len(digits_only) == 9 and not digits_only.startswith("1"):
            # 627172210 -> 627172210 (probablement français)
            root = digits_only
            is_french = True
        else:
            # Vérifier si c'est un indicatif connu
            is_valid_intl = False
            for code, (min_len, max_len) in country_codes.items():
                if digits_only.startswith(code) and min_len <= len(digits_only) <= max_len:
                    is_valid_intl = True
                    break

            # Si pas d'indicatif reconnu mais assez de chiffres, accepter quand même
            if not is_valid_intl and len(digits_only) < 9:
                return identifier, False, False

            # Numéro étranger: garder tel quel
            root = digits_only

        return root, True, is_french

    def _preprocess_call_log_rows(
        self,
        rows: List[Dict[str, Any]],
        cols: List[str],
        parties_col: str,
        source: str,
    ) -> List[Dict[str, Any]]:
        """
        PrÃ©traitement strict CallLogProcessor.ts:
        - Exclusion source: Recents|InteractionC|KnowledgeC|Native Messages|Threads
        - Exclusion hash: colonne '#' contient '(\\d+)'
        - Normalisation: Source vide -> Natif, Parties vide -> ""
        - Filtre source demandÃ©e
        - DÃ©doublonnage Parties-Date-Time-Duration-Source
        """
        source_re = re.compile(r"Recents|InteractionC|KnowledgeC|Native Messages|Threads", re.IGNORECASE)
        source_match = str(source or "").strip().lower()
        hash_col = "#" if "#" in cols else None

        prepared: List[Dict[str, Any]] = []
        for row in rows:
            src_val = str(row.get("Source", "") or "").strip()
            hash_val = str(row.get(hash_col, "") or "") if hash_col else ""

            if source_re.search(src_val):
                continue
            if hash_col and re.search(r"\(\d+\)", hash_val):
                continue

            normalized = dict(row)
            normalized_source = src_val or "Natif"
            normalized["Source"] = normalized_source
            normalized[parties_col] = str(normalized.get(parties_col, "") or "")

            if source_match and normalized_source.lower() != source_match:
                continue

            prepared.append(normalized)

        # Déduplication améliorée:
        # - Comparer Parties, Date, Duration, Direction, Status, Video call, Source, Account
        # - Supprimer si Time diffère de 0-1 seconde
        # - Garder si Duration = 0 (impossible de déterminer si doublon)

        def parse_time_to_seconds(time_str: str) -> int:
            """Parse time string en secondes - supporte plusieurs formats"""
            try:
                if not time_str:
                    return 0
                # Chercher le pattern HH:MM:SS
                match = re.search(r'(\d{1,2}):(\d{2}):(\d{2})', str(time_str))
                if match:
                    h, m, s = int(match.group(1)), int(match.group(2)), int(match.group(3))
                    return h * 3600 + m * 60 + s
                # Chercher le pattern HH:MM
                match = re.search(r'(\d{1,2}):(\d{2})', str(time_str))
                if match:
                    h, m = int(match.group(1)), int(match.group(2))
                    return h * 3600 + m * 60
                return 0
            except:
                return 0

        def get_duration_seconds(duration_str: str) -> int:
            """Parse duration pour vérifier si c'est 0"""
            try:
                if not duration_str or duration_str == '' or duration_str == 'None':
                    return 0
                if ':' in str(duration_str):
                    parts = str(duration_str).split(':')
                    total = 0
                    for i, p in enumerate(parts):
                        total += int(p) * (60 ** (len(parts) - 1 - i))
                    return total
                return int(re.sub(r'[^0-9]', '', str(duration_str)) or 0)
            except:
                return 0

        seen: Dict[tuple, Dict[str, Any]] = {}
        unique_rows: List[Dict[str, Any]] = []

        print(f"[DEDUP] Début déduplication: {len(prepared)} lignes, parties_col={parties_col}")

        for idx, row in enumerate(prepared):
            duration_sec = get_duration_seconds(str(row.get('Duration', '')))
            time_val = str(row.get('Time', ''))
            date_val = str(row.get('Date', ''))

            print(f"[DEDUP] Row {idx}: Duration={row.get('Duration')} ({duration_sec}s), Time={time_val}, Date={date_val}")

            # Si Duration = 0, toujours garder (impossible de déterminer si doublon)
            if duration_sec == 0:
                print(f"[DEDUP] Row {idx}: Duration=0, gardé")
                unique_rows.append(row)
                continue

            # Normaliser la Date (Excel serial number -> juste la partie entière = jour)
            date_for_key = date_val
            try:
                # Si c'est un numéro de série Excel (ex: 45978.020951516206), prendre juste la partie entière
                float_date = float(date_val)
                date_for_key = str(int(float_date))
            except:
                pass

            # Créer clé de comparaison (colonnes: Parties, Date, Duration, Direction, Status, Video call, Source, Account)
            key = (
                str(row.get(parties_col, '')),
                date_for_key,
                str(row.get('Duration', '')),
                str(row.get('Direction', '')),
                str(row.get('Status', '')),
                str(row.get('Video call', '')),
                str(row.get('Source', '')),
                str(row.get('Account', ''))
            )

            time_seconds = parse_time_to_seconds(time_val)
            print(f"[DEDUP] Row {idx}: Date normalisée={date_for_key}, Key hash, Time={time_val} -> {time_seconds}s")

            if key in seen:
                # Vérifier si Time diffère de 0 ou 1 seconde
                existing_time = seen[key]['time_seconds']
                time_diff = abs(time_seconds - existing_time)
                print(f"[DEDUP] Row {idx}: Clé existante! Time diff = {time_seconds} - {existing_time} = {time_diff}s")
                if time_diff <= 1:
                    # Doublon (même clé + Time diffère de 0-1s) - ignorer cette entrée
                    print(f"[DEDUP] Row {idx}: DOUBLON SUPPRIMÉ (diff <= 1s)")
                    continue
                else:
                    # Time différent de plus d'1 seconde - garder les deux
                    print(f"[DEDUP] Row {idx}: Gardé (diff > 1s)")
                    unique_rows.append(row)
            else:
                print(f"[DEDUP] Row {idx}: Nouvelle clé, gardé")
                seen[key] = {'row': row, 'time_seconds': time_seconds}
                unique_rows.append(row)

        print(f"[DEDUP] Fin déduplication: {len(prepared)} -> {len(unique_rows)} lignes")
        return unique_rows

    def _call_datetime_sort_key(self, date_val: str, time_val: str) -> Tuple[int, int, int, int, int, int, str, str]:
        """
        Clé de tri Date+Time sans conversion datetime.
        Formats gérés: YYYY-MM-DD/ YYYY/MM/DD et DD/MM/YYYY.
        Gère le cas où la date est dans time_val au lieu de date_val.
        """
        # Nettoyer les valeurs pour enlever les numéros de série Excel
        date_str = self._clean_excel_datetime(str(date_val or "").strip())
        time_str = self._clean_excel_datetime(str(time_val or "").strip())

        year = month = day = 0
        hour = minute = second = 0

        # Essayer d'extraire la date depuis date_str d'abord
        datetime_source = date_str
        m_iso = re.match(r"^\s*(\d{4})[-/](\d{1,2})[-/](\d{1,2})", date_str)
        m_fr = re.match(r"^\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})", date_str)

        # Si pas de date dans date_str, chercher dans time_str (parfois la date complète y est)
        if not m_iso and not m_fr:
            m_iso = re.match(r"^\s*(\d{4})[-/](\d{1,2})[-/](\d{1,2})", time_str)
            m_fr = re.match(r"^\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})", time_str)
            if m_iso or m_fr:
                datetime_source = time_str

        if m_iso:
            year, month, day = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        elif m_fr:
            day, month, year = int(m_fr.group(1)), int(m_fr.group(2)), int(m_fr.group(3))

        # Extraire l'heure - chercher dans time_str puis dans datetime_source
        time_to_parse = time_str if time_str else datetime_source
        # Chercher l'heure après la date si présente (format "DD/MM/YYYY HH:MM:SS")
        m_time = re.search(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", time_to_parse)
        if m_time:
            hour = int(m_time.group(1))
            minute = int(m_time.group(2))
            second = int(m_time.group(3) or 0)

        return (year, month, day, hour, minute, second, date_str, time_str)

    def _clean_excel_datetime(self, value: str) -> str:
        """
        Supprime le numéro de série Excel (ex: 45846.863119085647) du début d'une valeur datetime.
        Ex: "45846.863119085647 08/07/2025 20:42:53(UTC+2)" -> "08/07/2025 20:42:53(UTC+2)"
        Gère aussi le cas où la valeur est UNIQUEMENT un numéro de série (retourne "").
        """
        val = str(value or "").strip()
        if not val:
            return ""

        # Cas 1: Numéro de série suivi d'un espace et d'une date (DD/MM/YYYY)
        match = re.match(r'^\d+(?:\.\d+)?\s+(\d{1,2}/\d{1,2}/\d{4}.*)$', val)
        if match:
            return match.group(1)

        # Cas 2: Numéro de série suivi d'un espace et d'une date (YYYY-MM-DD ou YYYY/MM/DD)
        match = re.match(r'^\d+(?:\.\d+)?\s+(\d{4}[-/]\d{1,2}[-/]\d{1,2}.*)$', val)
        if match:
            return match.group(1)

        # Cas 3: Valeur qui est UNIQUEMENT un numéro de série Excel (ex: "45846.863119085647")
        # Ces numéros sont typiquement entre 40000 et 50000 pour des dates récentes (2010-2040)
        if re.match(r'^\d{4,5}(?:\.\d+)?$', val):
            # C'est juste un numéro de série, on retourne vide
            return ""

        return val

    def _format_date_time(self, date_val: str, time_val: str) -> str:
        """
        Formate Date + Time en une seule chaîne, en nettoyant les numéros de série Excel.
        Gère le cas où Date est un numéro de série et Time contient la vraie datetime.
        """
        # Nettoyer les valeurs pour enlever les numéros de série Excel
        date_str = self._clean_excel_datetime(str(date_val or "").strip())
        time_str = self._clean_excel_datetime(str(time_val or "").strip())

        # Si time_str contient déjà une date (ex: "08/07/2025 20:42:53"),
        # ne pas ajouter date_str même s'il existe
        if time_str and re.match(r'\d{1,2}/\d{1,2}/\d{4}', time_str):
            return time_str
        if time_str and re.match(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', time_str):
            return time_str

        if date_str and time_str:
            return f"{date_str} {time_str}"
        return date_str or time_str or "N/C"

    def get_call_log_details_by_source(self, source: str) -> Dict[str, Any]:
        """
        RÃ©cupÃ¨re les dÃ©tails du journal d'appels pour une source spÃ©cifique:
        - nbappelentrant: nombre d'appels entrants (Direction = 'Incoming')
        - nbappelsortant: nombre d'appels sortants (Direction = 'Outgoing')
        - timedebut: date+heure la plus ancienne (texte)
        - timefin: date+heure la plus rÃ©cente (texte)
        """
        lf = self._scan("Call Log")
        if lf is None:
            return {
                "nbappelentrant": 0,
                "nbappelsortant": 0,
                "timedebut": "N/C",
                "timefin": "N/C"
            }

        cols = self.get_columns("Call Log")

        df = lf.collect()

        if len(df) == 0:
            return {
                "nbappelentrant": 0,
                "nbappelsortant": 0,
                "timedebut": "N/C",
                "timefin": "N/C"
            }

        parties_col = next((c for c in ["Parties", "Party", "Contact", "Number", "Phone"] if c in cols), None)
        if not parties_col:
            return {
                "nbappelentrant": 0,
                "nbappelsortant": 0,
                "timedebut": "N/C",
                "timefin": "N/C"
            }

        rows = self._preprocess_call_log_rows(df.to_dicts(), cols, parties_col, source)
        if not rows:
            return {
                "nbappelentrant": 0,
                "nbappelsortant": 0,
                "timedebut": "N/C",
                "timefin": "N/C"
            }

        nbappelentrant = 0
        nbappelsortant = 0

        for row in rows:
            direction = str(row.get("Direction", "") or "").strip().lower()
            if direction == "incoming":
                nbappelentrant += 1
            elif direction == "outgoing":
                nbappelsortant += 1

        timedebut = "N/C"
        timefin = "N/C"
        timestamps = []
        for row in rows:
            date_str = str(row.get("Date", "") or "").strip()
            time_str = str(row.get("Time", "") or "").strip()
            if not date_str and not time_str:
                continue
            sort_key = self._call_datetime_sort_key(date_str, time_str)
            display_val = self._format_date_time(date_str, time_str)
            timestamps.append((sort_key, display_val))

        if timestamps:
            timestamps.sort(key=lambda x: x[0])
            timedebut = timestamps[0][1]
            timefin = timestamps[-1][1]

        return {
            "nbappelentrant": nbappelentrant,
            "nbappelsortant": nbappelsortant,
            "timedebut": timedebut,
            "timefin": timefin
        }

    def get_call_log_top15_by_count(self, source: str) -> List[Dict[str, Any]]:
        """
        RÃ©cupÃ¨re le top 15 des contacts par nombre d'appels pour une source.
        Utilise la logique d'extraction et d'agrÃ©gation de CallLogProcessor.ts.
        """
        aggregated, _ = self._analyze_call_log(source)

        if not aggregated:
            return []

        # Trier par nombre d'appels et prendre le top 15
        sorted_contacts = sorted(aggregated.values(), key=lambda x: x["Nombre_appels"], reverse=True)
        sorted_contacts = self._filter_owner_and_remove_max(sorted_contacts)
        # Filtrer les numéros courts pour Natif (appels système)
        sorted_contacts = self._filter_short_phone_numbers_for_natif(sorted_contacts, source)[:15]

        # Construire le lookup Contacts pour les sources non-telephoniques
        # (Signal, Snapchat, Instagram, TikTok...) afin de recuperer l'identifiant utilisateur.
        source_lower = source.lower() if source else ""
        is_phone_source = source_lower in ("native messages", "natif", "whatsapp", "whatsapp business")
        needs_user_id = not is_phone_source
        contacts_lookup = self._build_contacts_user_id_lookup(source) if needs_user_id else {}

        # Formater la durÃ©e
        result = []
        for contact in sorted_contacts:
            identifier = self._display_identifier(contact.get("DisplayIdentifier", contact["Identifier"]), source)
            identifiant_utilisateur = ""

            # Recuperer l'identifiant utilisateur via le lookup Contacts
            # (Signal, Snapchat, Instagram, TikTok...). Cle = identifiant puis nom.
            if needs_user_id:
                lookup_info = None
                if identifier and identifier.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[identifier.lower()]
                elif contact["Name"] and contact["Name"].lower() in contacts_lookup:
                    lookup_info = contacts_lookup[contact["Name"].lower()]
                if lookup_info:
                    identifiant_utilisateur = lookup_info.get("identifiant_utilisateur") or ""

            result.append({
                "Identifier": identifier,
                "Identifiant_utilisateur": identifiant_utilisateur,
                "Name": contact["Name"],
                "Nombre_appels": contact["Nombre_appels"],
                "Appel_emis": contact["Appel_emis"],
                "Appel_recu": contact["Appel_recu"],
                "Duree_totale": self._seconds_to_duration(contact["Duree_totale_sec"])
            })

        return result

    def get_call_log_top15_by_duration(self, source: str) -> List[Dict[str, Any]]:
        """
        RÃ©cupÃ¨re le top 15 des contacts par durÃ©e d'appels pour une source.
        Utilise la logique d'extraction et d'agrÃ©gation de CallLogProcessor.ts.
        """
        aggregated, _ = self._analyze_call_log(source)

        if not aggregated:
            return []

        # Trier par durÃ©e et prendre le top 15
        sorted_contacts = sorted(aggregated.values(), key=lambda x: x["Duree_totale_sec"], reverse=True)
        sorted_contacts = self._filter_owner_and_remove_max(sorted_contacts)
        # Filtrer les numéros courts pour Natif (appels système)
        sorted_contacts = self._filter_short_phone_numbers_for_natif(sorted_contacts, source)[:15]

        # Construire le lookup Contacts pour les sources non-telephoniques
        # (Signal, Snapchat, Instagram, TikTok...) afin de recuperer l'identifiant utilisateur.
        source_lower = source.lower() if source else ""
        is_phone_source = source_lower in ("native messages", "natif", "whatsapp", "whatsapp business")
        needs_user_id = not is_phone_source
        contacts_lookup = self._build_contacts_user_id_lookup(source) if needs_user_id else {}

        # Formater la durÃ©e
        result = []
        for contact in sorted_contacts:
            identifier = self._display_identifier(contact.get("DisplayIdentifier", contact["Identifier"]), source)
            identifiant_utilisateur = ""

            # Recuperer l'identifiant utilisateur via le lookup Contacts
            # (Signal, Snapchat, Instagram, TikTok...). Cle = identifiant puis nom.
            if needs_user_id:
                lookup_info = None
                if identifier and identifier.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[identifier.lower()]
                elif contact["Name"] and contact["Name"].lower() in contacts_lookup:
                    lookup_info = contacts_lookup[contact["Name"].lower()]
                if lookup_info:
                    identifiant_utilisateur = lookup_info.get("identifiant_utilisateur") or ""

            result.append({
                "Identifier": identifier,
                "Identifiant_utilisateur": identifiant_utilisateur,
                "Name": contact["Name"],
                "Nombre_appels": contact["Nombre_appels"],
                "Duree_totale": self._seconds_to_duration(contact["Duree_totale_sec"])
            })

        return result

    def get_call_log_summary(self, source: str) -> List[Dict[str, Any]]:
        """
        Récupère le tableau résumé des appels pour une source.
        Pour sources sociales (Snapchat/Instagram/Signal/etc): Pseudonyme, Nom d'utilisateur, Identifiant utilisateur, Émis, Reçus, Appel vidéos, Appel manqué, Appel supprimé
        Pour sources téléphoniques (Natif/WhatsApp): Pseudonyme, Numéro de téléphone, Émis, Reçus, ...
        Trié par nombre d'appels (top 15).
        """
        aggregated, _ = self._analyze_call_log(source)

        if not aggregated:
            return []

        source_lower = source.lower() if source else ""
        # Sources téléphoniques: format 3 colonnes (Pseudonyme, Numéro de téléphone, stats)
        is_phone_source = source_lower in ("native messages", "natif", "whatsapp", "whatsapp business")
        # Signal: format spécial 2 colonnes (Pseudonyme, Identifiant utilisateur) - UUID, pas de numéro ni username
        is_signal = source_lower == "signal"
        # Sources sociales: format 4 colonnes (Pseudonyme, Nom utilisateur, Identifiant, stats)
        is_social_source = not is_phone_source and not is_signal

        # Construire le lookup depuis Contacts pour toutes les sources
        contacts_lookup = self._build_contacts_user_id_lookup(source)

        # Trier par nombre d'appels et prendre le top 15
        sorted_contacts = sorted(aggregated.values(), key=lambda x: x["Nombre_appels"], reverse=True)
        sorted_contacts = self._filter_owner_and_remove_max(sorted_contacts)
        # Filtrer les numéros courts pour Natif (appels système)
        sorted_contacts = self._filter_short_phone_numbers_for_natif(sorted_contacts, source)

        # Pour Signal: format spécial (Pseudonyme + Identifiant utilisateur uniquement)
        if is_signal:
            aggregated_rows: Dict[str, Dict[str, Any]] = {}

            for contact in sorted_contacts:
                current_name = contact["Name"] or ""
                current_identifier = contact["Identifier"] or ""  # UUID + pseudonyme pour Signal

                # Parser l'identifiant Signal: "UUID pseudonyme"
                # UUID avec tirets (ex: "0FD9CF43-C376-4D8E-8E58-108F337966F4 coub")
                uuid_match = re.match(r'^([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})\s*(.*)$', current_identifier)
                if uuid_match:
                    identifiant_utilisateur = uuid_match.group(1)
                    pseudonyme_from_identifier = uuid_match.group(2).strip() if uuid_match.group(2) else ""
                else:
                    # Essayer UUID sans tirets (32 caractères hex) et reformater avec tirets
                    uuid_no_dash = re.match(r'^([0-9A-Fa-f]{32})\s*(.*)$', current_identifier)
                    if uuid_no_dash:
                        raw_uuid = uuid_no_dash.group(1)
                        # Reformater: XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
                        identifiant_utilisateur = f"{raw_uuid[:8]}-{raw_uuid[8:12]}-{raw_uuid[12:16]}-{raw_uuid[16:20]}-{raw_uuid[20:]}"
                        pseudonyme_from_identifier = uuid_no_dash.group(2).strip() if uuid_no_dash.group(2) else ""
                    else:
                        # Pas un UUID, utiliser tel quel
                        identifiant_utilisateur = current_identifier
                        pseudonyme_from_identifier = ""

                # Chercher dans le lookup Contacts par UUID ou name
                lookup_info = None
                if identifiant_utilisateur and identifiant_utilisateur.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[identifiant_utilisateur.lower()]
                elif current_name and current_name.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[current_name.lower()]

                # Déterminer le pseudonyme final
                if lookup_info:
                    pseudonyme = lookup_info.get("pseudonyme") or "-"
                    # Utiliser l'identifiant du lookup si disponible
                    if lookup_info.get("identifiant_utilisateur"):
                        identifiant_utilisateur = lookup_info["identifiant_utilisateur"]
                else:
                    # Pas trouvé dans Contacts: utiliser pseudonyme extrait de l'identifier, sinon Name
                    if pseudonyme_from_identifier:
                        pseudonyme = pseudonyme_from_identifier
                    elif current_name:
                        pseudonyme = current_name
                    else:
                        pseudonyme = "-"

                # Clé d'agrégation
                key = f"{pseudonyme}|{identifiant_utilisateur}".lower()

                if key in aggregated_rows:
                    aggregated_rows[key]["Émis"] += contact["Appel_emis"]
                    aggregated_rows[key]["Reçus"] += contact["Appel_recu"]
                    aggregated_rows[key]["Appel vidéos"] += contact["Appel_video"]
                    aggregated_rows[key]["Appel manqué"] += contact["Appel_manque"]
                    aggregated_rows[key]["Appel supprimé"] += contact["Appel_supprime"]
                else:
                    aggregated_rows[key] = {
                        "Pseudonyme": pseudonyme,
                        "Identifiant utilisateur": identifiant_utilisateur,
                        "Émis": contact["Appel_emis"],
                        "Reçus": contact["Appel_recu"],
                        "Appel vidéos": contact["Appel_video"],
                        "Appel manqué": contact["Appel_manque"],
                        "Appel supprimé": contact["Appel_supprime"]
                    }

            result = sorted(aggregated_rows.values(),
                          key=lambda x: x["Émis"] + x["Reçus"] + x["Appel vidéos"] + x["Appel manqué"] + x["Appel supprimé"],
                          reverse=True)[:15]

        # Pour sources sociales: agréger les doublons après résolution des noms
        elif is_social_source:
            # D'abord résoudre les noms via Contacts, puis agréger
            aggregated_rows: Dict[str, Dict[str, Any]] = {}

            for contact in sorted_contacts:
                current_name = contact["Name"] or ""
                current_identifier = self._display_identifier(contact.get("DisplayIdentifier", contact["Identifier"]), source)

                # Chercher dans le lookup par identifier ou name
                lookup_info = None
                if current_identifier and current_identifier.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[current_identifier.lower()]
                elif current_name and current_name.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[current_name.lower()]

                # Déterminer les valeurs finales
                pseudonyme = "-"
                nom_utilisateur = current_name if current_name else ""
                identifiant_utilisateur = ""

                if lookup_info:
                    # Valeurs depuis Contacts
                    pseudonyme = lookup_info.get("pseudonyme") or "-"
                    if lookup_info.get("nom_utilisateur"):
                        nom_utilisateur = lookup_info["nom_utilisateur"]
                    identifiant_utilisateur = lookup_info.get("identifiant_utilisateur") or ""

                    # Si current_identifier == identifiant_utilisateur du lookup, c'était un ID pas un pseudo
                    # Si current_identifier == pseudonyme du lookup, c'était bien un pseudo
                    # Sinon, vérifier si current_identifier ressemble à l'identifiant
                    if current_identifier and current_identifier.lower() == identifiant_utilisateur.lower():
                        # L'identifier de base était en fait l'identifiant utilisateur
                        pass  # pseudonyme reste celui du lookup
                    elif current_identifier and pseudonyme and current_identifier.lower() != pseudonyme.lower():
                        # current_identifier n'est ni le pseudo ni l'ID du lookup
                        # C'est probablement l'identifiant utilisateur
                        if not identifiant_utilisateur:
                            identifiant_utilisateur = current_identifier
                else:
                    # Pas trouvé dans Contacts: règle de base pour sources sociales
                    # Name → Pseudonyme (carnet d'adresses), Identifier → Nom d'utilisateur (username app)
                    pseudonyme = current_name if current_name else "-"
                    nom_utilisateur = current_identifier
                    # Vérifier si nom_utilisateur ressemble à un ID (UUID ou numérique long)
                    # Si oui, c'est probablement un identifiant_utilisateur mal placé
                    if nom_utilisateur and (self._looks_like_user_id(nom_utilisateur)):
                        identifiant_utilisateur = nom_utilisateur
                        nom_utilisateur = "-"

                # Clé d'agrégation
                key = f"{pseudonyme}|{nom_utilisateur}|{identifiant_utilisateur}".lower()

                if key in aggregated_rows:
                    # Agréger les compteurs
                    aggregated_rows[key]["Émis"] += contact["Appel_emis"]
                    aggregated_rows[key]["Reçus"] += contact["Appel_recu"]
                    aggregated_rows[key]["Appel vidéos"] += contact["Appel_video"]
                    aggregated_rows[key]["Appel manqué"] += contact["Appel_manque"]
                    aggregated_rows[key]["Appel supprimé"] += contact["Appel_supprime"]
                else:
                    # Sources sociales: 3 colonnes d'identité + stats
                    aggregated_rows[key] = {
                        "Pseudonyme": pseudonyme,
                        "Nom d'utilisateur": nom_utilisateur,
                        "Identifiant utilisateur": identifiant_utilisateur,
                        "Émis": contact["Appel_emis"],
                        "Reçus": contact["Appel_recu"],
                        "Appel vidéos": contact["Appel_video"],
                        "Appel manqué": contact["Appel_manque"],
                        "Appel supprimé": contact["Appel_supprime"]
                    }

            # Trier par total d'appels et prendre top 15
            result = sorted(aggregated_rows.values(),
                          key=lambda x: x["Émis"] + x["Reçus"] + x["Appel vidéos"] + x["Appel manqué"] + x["Appel supprimé"],
                          reverse=True)[:15]
        else:
            # Pour les sources téléphoniques: 3 colonnes (Pseudonyme, Numéro, stats)
            result = []
            for contact in sorted_contacts[:15]:
                current_name = contact["Name"] or "Inconnu"
                current_identifier = self._display_identifier(contact.get("DisplayIdentifier", contact["Identifier"]), source)

                row = {
                    "Pseudonyme": current_name,
                    "Numéro de téléphone": current_identifier,
                    "Émis": contact["Appel_emis"],
                    "Reçus": contact["Appel_recu"],
                    "Appel vidéos": contact["Appel_video"],
                    "Appel manqué": contact["Appel_manque"],
                    "Appel supprimé": contact["Appel_supprime"]
                }
                result.append(row)

        return result

    def _looks_like_user_id(self, value: str) -> bool:
        """Vérifie si une valeur ressemble à un identifiant utilisateur (UUID ou numérique long)"""
        if not value:
            return False
        # UUID pattern (avec ou sans tirets)
        if re.match(r'^[0-9a-f]{8}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{12}$', value, re.IGNORECASE):
            return True
        # Numérique long (>= 10 chiffres)
        if re.match(r'^\d{10,}$', value):
            return True
        return False

    def _get_call_log_top15_placeholder(self, source: str) -> List[Dict[str, Any]]:
        """
        Version simplifiÃ©e - placeholder pour compatibilitÃ©.
        """
        lf = self._scan("Call Log")
        if lf is None:
            return []

        cols = self.get_columns("Call Log")

        if "Source" in cols:
            lf = lf.filter(pl.col("Source") == source)
        elif source != "Natif":
            return []

        df = lf.collect()
        if len(df) == 0:
            return []

        parties_col = next((c for c in ["Parties", "Party", "Contact", "Number", "Phone"] if c in cols), None)
        if not parties_col:
            return []

        contact_stats: Dict[str, Dict[str, Any]] = {}

        for row in df.to_dicts():
            parties = str(row.get(parties_col, "") or "")
            duration_str = str(row.get("Duration", "00:00:00") or "00:00:00")

            identifier = parties.split('\n')[0].strip() if parties else "Inconnu"
            for prefix in ["From:", "To:", "General:"]:
                if identifier.startswith(prefix):
                    identifier = identifier[len(prefix):].strip()
                    break

            if not identifier:
                identifier = "Inconnu"

            if identifier not in contact_stats:
                contact_stats[identifier] = {
                    "Identifier": identifier,
                    "Name": "",
                    "Nombre_appels": 0,
                    "Duree_totale_sec": 0
                }

            contact_stats[identifier]["Nombre_appels"] += 1
            contact_stats[identifier]["Duree_totale_sec"] += self._duration_to_seconds(duration_str)

        # Trier par durÃ©e et prendre le top 15
        sorted_contacts = sorted(contact_stats.values(), key=lambda x: x["Duree_totale_sec"], reverse=True)[:15]

        # Formater la durÃ©e
        for contact in sorted_contacts:
            contact["Duree_totale"] = self._seconds_to_duration(contact.pop("Duree_totale_sec"))

        return sorted_contacts

    def _duration_to_seconds(self, duration_str: str) -> int:
        """Convertit une durÃ©e HH:MM:SS en secondes"""
        try:
            parts = duration_str.split(':')
            if len(parts) == 3:
                return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
            elif len(parts) == 2:
                return int(parts[0]) * 60 + int(parts[1])
            return 0
        except:
            return 0

    def _seconds_to_duration(self, seconds: int) -> str:
        """Convertit des secondes en durÃ©e HH:MM:SS"""
        hours = seconds // 3600
        minutes = (seconds % 3600) // 60
        secs = seconds % 60
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"



    # ═══════════════════════════════════════════════════════════════════════════════
    # ANALYSE DES CHATS / MESSAGERIES (inspiré de ConvLogProcessor.ts)
    # ═══════════════════════════════════════════════════════════════════════════════

    def _normalize_phone_chat(self, phone: str, whatsapp_mode: bool = False) -> str:
        """Normalise un numéro de téléphone (comme ConvLogProcessor.ts)"""
        if not phone:
            return phone

        # Enlever le suffixe WhatsApp
        cleaned = re.sub(r'@s\.whatsapp\.net$', '', str(phone), flags=re.IGNORECASE).strip()

        # Enlever espaces, tirets, points, parenthèses
        digits = re.sub(r'[\s\-\.\(\)]', '', cleaned)

        # Si déjà au format international avec +
        if re.match(r'^\+\d{8,15}$', digits):
            return digits

        # Ne pas forcer le format international hors WhatsApp
        if not whatsapp_mode:
            return cleaned

        # Numéros longs sans 0 initial - probablement international
        if re.match(r'^\d{10,15}$', digits) and not digits.startswith('0'):
            return '+' + digits

        return cleaned

    def _has_enough_digits(self, value: str, min_digits: int = 6) -> bool:
        """Vérifie si une chaîne contient assez de chiffres"""
        digits = re.findall(r'\d', str(value))
        return len(digits) >= min_digits

    def _extract_chat_contacts(self, participants: str, source_value: str) -> list:
        """
        Extrait les contacts des participants - logique identique à ConvLogProcessor.ts extractContacts.
        Supporte multi-ligne et single-line.
        """
        contacts = []
        if not participants or not isinstance(participants, str):
            return contacts

        source_lower = (source_value or "").lower()

        # Nettoyer _x000D_ (retour chariot Windows encodé)
        cleaned = re.sub(r'_x000D_', '', participants, flags=re.IGNORECASE).strip()
        if not cleaned:
            return contacts

        # Split par newline pour traiter ligne par ligne
        lines = cleaned.split('\n')

        if source_lower == "snapchat":
            # Snapchat: format "username nom" - peut être multi-ligne ou single-line
            snap_pattern = re.compile(r'^([^\s]+)\s+(.*)$')

            def is_snapchat_username(token):
                """Un username Snapchat contient chiffres, . ou _ et est alphanumérique"""
                if len(token) < 3 or token.startswith('('):
                    return False
                has_special = bool(re.search(r'[\d\._]', token))
                is_alphanum = bool(re.match(r'^[a-zA-Z0-9\._]+$', token))
                return has_special and is_alphanum

            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # Vérifier si single-line avec plusieurs contacts
                tokens = line.split()
                usernames_in_line = [t for t in tokens if is_snapchat_username(t)]

                if len(usernames_in_line) > 1:
                    # Format single-line: "yapasrienla7575 juan lopez (owner) nautilus.75 D"
                    current_username = None
                    current_name_parts = []

                    for token in tokens:
                        if is_snapchat_username(token):
                            if current_username:
                                contacts.append({
                                    "Username": current_username,
                                    "Name": ' '.join(current_name_parts) or "Inconnu"
                                })
                            current_username = token
                            current_name_parts = []
                        else:
                            current_name_parts.append(token)

                    if current_username:
                        contacts.append({
                            "Username": current_username,
                            "Name": ' '.join(current_name_parts) or "Inconnu"
                        })
                else:
                    # Format classique: une ligne = un contact "username nom"
                    match = snap_pattern.match(line)
                    if match:
                        contacts.append({
                            "Username": match.group(1).strip() or "Inconnu",
                            "Name": match.group(2).strip() or "Inconnu"
                        })

        elif source_lower == "whatsapp":
            # WhatsApp: format @s.whatsapp.net ou numéro + pseudo
            for line in lines:
                contact_info = line.strip()

                # Nettoyer _x000D_
                contact_info = re.sub(r'_x000D_', '', contact_info, flags=re.IGNORECASE).strip()
                if not contact_info:
                    continue

                # Vérifier si plusieurs @s.whatsapp.net sur la même ligne (format single-line)
                if contact_info.count('@s.whatsapp.net') > 1:
                    # Format single-line: "33660048980@s.whatsapp.net RR (owner) 971504681584@s.whatsapp.net zey"
                    wa_pattern = re.compile(r'(\d+)@s\.whatsapp\.net\s*([^@]*?)(?=\d+@s\.whatsapp\.net|$)', re.IGNORECASE)
                    for match in wa_pattern.finditer(contact_info):
                        phone = match.group(1).strip()
                        name = match.group(2).strip() or 'Inconnu'
                        if phone and self._has_enough_digits(phone):
                            normalized = self._normalize_phone_chat(phone, whatsapp_mode=True)
                            contacts.append({"Phone": normalized, "Name": name})
                elif '@s.whatsapp.net' in contact_info:
                    # Format single: "33660048980@s.whatsapp.net RR" ou "+212 6 89 07 48 95 212689074895@s.whatsapp.net rita"
                    at_index = contact_info.find('@s.whatsapp.net')
                    before_at = contact_info[:at_index].strip()
                    after_at = contact_info[at_index + len('@s.whatsapp.net'):].strip()
                    name = re.sub(r'_x000D_', '', after_at, flags=re.IGNORECASE).strip() or 'Inconnu'

                    # before_at peut être: "33660048980" ou "+212 6 89 07 48 95 212689074895"
                    if ' ' in before_at:
                        parts = before_at.split()
                        if parts[0].startswith('+'):
                            phoneparts = []
                            for part in parts:
                                if len(phoneparts) == 0 or not re.match(r'^\d{6,}$', part):
                                    phoneparts.append(part)
                                else:
                                    break
                            phone = ' '.join(phoneparts)
                        else:
                            phone = parts[-1]
                    else:
                        phone = before_at

                    if phone:
                        normalized = self._normalize_phone_chat(phone, whatsapp_mode=True)
                        if self._has_enough_digits(normalized):
                            contacts.append({"Phone": normalized, "Name": name})
                else:
                    # Cas robuste: numero + pseudo
                    spaced_phone_match = re.match(r'^(\+?\d[\d\s\-\.\(\)]{5,}\d)\s+(.+)$', contact_info)
                    if spaced_phone_match:
                        phone = spaced_phone_match.group(1).strip()
                        name = spaced_phone_match.group(2).strip() or 'Inconnu'
                    else:
                        phone = contact_info
                        name = 'Inconnu'

                    if phone:
                        normalized = self._normalize_phone_chat(phone, whatsapp_mode=True)
                        if self._has_enough_digits(normalized):
                            contacts.append({"Phone": normalized, "Name": name})

        elif source_lower == "signal":
            # Signal: UUID, numéro de téléphone, ou SignalGroup:xxx
            # Pattern UUID: 8-4-4-4-12 hex
            uuid_pattern = re.compile(r'([A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12})', re.IGNORECASE)
            # Pattern phone: +digits
            phone_pattern = re.compile(r'(\+\d{10,15})')
            # Pattern Signal group ID
            signal_group_pattern = re.compile(r'SignalGroup:([A-Za-z0-9_\-]{20,})')

            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # Trouver tous les identifiants sur cette ligne
                all_matches = []
                for m in uuid_pattern.finditer(line):
                    all_matches.append((m.start(), m.end(), 'uuid', m.group(1)))
                for m in phone_pattern.finditer(line):
                    all_matches.append((m.start(), m.end(), 'phone', m.group(1)))
                for m in signal_group_pattern.finditer(line):
                    all_matches.append((m.start(), m.end(), 'uuid', m.group(1)))

                # Trier par position
                all_matches.sort(key=lambda x: x[0])

                if all_matches:
                    # Extraire nom après chaque identifiant jusqu'au prochain
                    for i, (start, end, id_type, identifier) in enumerate(all_matches):
                        next_start = all_matches[i + 1][0] if i + 1 < len(all_matches) else len(line)
                        name = line[end:next_start].strip() or 'Inconnu'

                        contact = {"Name": name}
                        if id_type == 'uuid':
                            contact["UUID"] = identifier
                        else:
                            contact["Phone"] = self._normalize_phone_chat(identifier, whatsapp_mode=False)
                        contacts.append(contact)
                else:
                    # Fallback: pattern simple identifiant + nom
                    signal_simple = re.compile(r'(?:SignalGroup:)?([A-Za-z0-9_\-]{20,}|[A-F0-9\-]{36}|\+?\d[\d\s\-\.]{8,}|\d{5,20})\s*(.*)')
                    match = signal_simple.match(line)
                    if match:
                        identifier = match.group(1).strip()
                        name = match.group(2).strip() or 'Inconnu'
                        contact = {"Name": name}

                        if re.match(r'^[A-F0-9\-]{36}$', identifier, re.IGNORECASE):
                            contact["UUID"] = identifier
                        elif re.match(r'^[A-Za-z0-9_\-]{20,}$', identifier) and not re.match(r'^\d+$', identifier):
                            contact["UUID"] = identifier
                        else:
                            contact["Phone"] = self._normalize_phone_chat(identifier)
                        contacts.append(contact)

        elif source_lower == "native messages":
            # Native Messages: formats possibles:
            # 1. Multi-ligne: "+33618472265\n+3300000000000 (owner)" → contact est +33618472265, ignorer (owner)
            # 2. Single-ligne: "+590690142735 nom (owner) Unknown" → contact est +590690142735 nom
            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # Si la ligne contient (owner), c'est le propriétaire → IGNORER cette ligne
                if '(owner)' in line.lower():
                    # Vérifier si c'est format single-line avec contact AVANT owner
                    # Ex: "+590690142735 nom (owner) Unknown" → extraire +590690142735 nom
                    parts = re.split(r'\s*\(owner\)\s*', line, flags=re.IGNORECASE)
                    contact_part = parts[0].strip()

                    # Si contact_part contient plusieurs numéros/tokens, le premier est le contact
                    # Ex: "+33618472265 +3300000000000" → +33618472265 est le contact
                    tokens = contact_part.split()
                    if len(tokens) >= 2:
                        # Vérifier si le premier token est un numéro (le contact)
                        first_token = tokens[0]
                        if re.match(r'^\+?\d[\d\s\-\.]*$', first_token):
                            phone = self._normalize_phone_chat(first_token)
                            # Le reste (sauf le dernier numéro qui pourrait être owner) est le nom
                            remaining = tokens[1:]
                            # Si le dernier token est un numéro, c'est probablement le owner inline
                            if remaining and re.match(r'^\+?\d+$', remaining[-1]):
                                remaining = remaining[:-1]
                            name = ' '.join(remaining) if remaining else "Inconnu"
                            if self._has_enough_digits(phone, 3):
                                contacts.append({"Phone": phone, "Name": name})
                    # Sinon, c'est juste le propriétaire seul sur la ligne → ignorer
                    continue

                # Pas de (owner), traiter comme contact (numéro ou username)
                phone_match = re.match(r'^(\+?\d[\d\s\-\.]*\d|\d+)\s*(.*)$', line)
                if phone_match:
                    phone = self._normalize_phone_chat(phone_match.group(1))
                    if self._has_enough_digits(phone, 3):
                        contacts.append({
                            "Phone": phone,
                            "Name": phone_match.group(2).strip() or "Inconnu"
                        })

        elif source_lower in ["instagram", "tiktok"]:
            # Instagram, TikTok: format single-line avec usernames
            # Format: "username1 nom1 username2 nom2 (owner)"

            def is_username(token, source):
                """Détecte si un token est un username"""
                if len(token) < 2 or token.startswith('('):
                    return False
                is_alphanum = bool(re.match(r'^[a-zA-Z0-9\._]+$', token))
                if not is_alphanum:
                    return False
                has_special = bool(re.search(r'[\d\._]', token))
                is_lowercase = token == token.lower()
                return has_special or is_lowercase

            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # Format username: chercher tous les usernames sur la ligne
                tokens = line.split()
                usernames_in_line = [t for t in tokens if is_username(t, source_lower)]

                if len(usernames_in_line) > 1:
                    # Plusieurs usernames - format single-line
                    current_username = None
                    current_name_parts = []

                    for token in tokens:
                        if is_username(token, source_lower):
                            if current_username:
                                contacts.append({
                                    "Username": current_username,
                                    "Name": ' '.join(current_name_parts) or "Inconnu"
                                })
                            current_username = token
                            current_name_parts = []
                        else:
                            current_name_parts.append(token)

                    if current_username:
                        contacts.append({
                            "Username": current_username,
                            "Name": ' '.join(current_name_parts) or "Inconnu"
                        })
                elif len(usernames_in_line) == 1:
                    # Un seul username
                    match = re.match(r'^([^\s]+)\s+(.*)$', line)
                    if match:
                        contacts.append({
                            "Username": match.group(1).strip() or 'Inconnu',
                            "Name": match.group(2).strip() or 'Inconnu'
                        })

        else:
            # Extraction générique (autres sources)
            phone_single_pattern = re.compile(r'^(\+?\d[\d\s\-\.]{6,}\d|\d{5,20})(?:@s\.whatsapp\.net)?(?:\s+(.+))?$')
            phone_multi_pattern = re.compile(r'(\+?\d[\d\s\-\.]{6,}\d|\d{5,20})(?:@s\.whatsapp\.net)?')
            snap_pattern = re.compile(r'^([^\s]+)\s+(.*)$')

            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # Chercher tous les numéros sur cette ligne
                phone_matches = list(phone_multi_pattern.finditer(line))

                if len(phone_matches) > 1:
                    # Plusieurs numéros sur la même ligne - format single-line
                    for i, m in enumerate(phone_matches):
                        phone = self._normalize_phone_chat(m.group(1))
                        if not self._has_enough_digits(phone):
                            continue
                        # Nom = texte entre fin de ce numéro et début du prochain
                        next_start = phone_matches[i + 1].start() if i + 1 < len(phone_matches) else len(line)
                        name = line[m.end():next_start].strip() or 'Inconnu'
                        contacts.append({"Phone": phone, "Name": name})
                elif len(phone_matches) == 1:
                    # Un seul numéro
                    phone_match = phone_single_pattern.match(line)
                    if phone_match:
                        phone = self._normalize_phone_chat(phone_match.group(1))
                        if self._has_enough_digits(phone):
                            contacts.append({
                                "Phone": phone,
                                "Name": (phone_match.group(2) or '').strip() or 'Inconnu'
                            })
                else:
                    # Pas de numéro - essayer format username
                    snap_match = snap_pattern.match(line)
                    if snap_match:
                        contacts.append({
                            "Username": snap_match.group(1).strip() or 'Inconnu',
                            "Name": snap_match.group(2).strip() or 'Inconnu'
                        })

        return contacts

    def _analyze_chats(self, source_value: str = None) -> list:
        """
        Analyse les chats par contact et source.
        NOUVELLE MÉTHODE: Compte les messages envoyés par chaque contact via la colonne "From".
        - Participants: identifie les contacts qui parlent avec l'owner
        - From: compte uniquement quand le contact est l'expéditeur du message
        """
        lf = self._scan("Chats")
        if lf is None:
            return []

        cols = self.get_columns("Chats")
        participants_col = next((c for c in ["Participants", "Parties", "Contact"] if c in cols), None)
        source_col = next((c for c in ["Source", "Application", "App", "Platform"] if c in cols), None)
        from_col = "From" if "From" in cols else None

        if not participants_col:
            return []

        # Ajouter Account si présent
        account_col = "Account" if "Account" in cols else None

        # Ajouter Instant Message # pour filtrer les lignes vides
        im_col = "Instant Message #" if "Instant Message #" in cols else None

        # Colonne Label pour filtrage
        label_col = "Label" if "Label" in cols else None

        # === ÉTAPE 1: Filtrer les messages avec Label contenant "Recalled" ou "System" ===
        if label_col:
            lf = lf.filter(
                pl.col(label_col).is_null() |
                (
                    ~pl.col(label_col).cast(pl.Utf8).str.to_lowercase().str.contains("recalled").fill_null(False) &
                    ~pl.col(label_col).cast(pl.Utf8).str.to_lowercase().str.contains("system").fill_null(False)
                )
            )

        # === ÉTAPE 2: Filtrer Instant Message # vide ===
        if im_col:
            lf = lf.filter(pl.col(im_col).is_not_null() & (pl.col(im_col).cast(pl.Utf8).str.strip_chars() != ""))

        # === ÉTAPE 3: Filtrer les sources exclues ===
        excluded_sources = ['recents', 'interactionc', 'knowledgec', 'threads', 'facebook']
        if source_col:
            lf = lf.filter(~pl.col(source_col).str.to_lowercase().is_in(excluded_sources))
            if source_value:
                lf = lf.filter(pl.col(source_col).str.to_lowercase() == source_value.lower())

        # === ÉTAPE 4: Dédoublonnage ===
        dedup_cols = []
        for col_name in ["Start Time: Time", "Last Activity: Time", "Timestamp: Time", "Participants", "Body"]:
            if col_name in cols:
                dedup_cols.append(col_name)
        # Ajouter Label pour dédoublonnage
        if label_col:
            dedup_cols.append(label_col)
        # Ajouter les colonnes contenant Attachment
        for col_name in cols:
            if "Attachment" in col_name:
                dedup_cols.append(col_name)

        if dedup_cols:
            lf = lf.unique(subset=dedup_cols, keep="first")

        # Collecter toutes les lignes pour traitement
        df = lf.collect()
        rows = df.to_dicts()

        # === ÉTAPE 5: Collecter les identifiants propriétaires ===
        # PRIORITÉ 1: Depuis User Accounts (colonne Username)
        owner_identifiers = self._get_owner_usernames_from_user_accounts()

        # PRIORITÉ 2: Depuis la colonne Account (TOUJOURS, plus seulement si UA vide).
        # On strip "@s.whatsapp.net" AVANT de retirer les points, sinon le suffix
        # devient "@swhatsappnet" et l'ancien regex echoue.
        if account_col:
            seen_accounts = set()
            for row in rows:
                account_val = str(row.get(account_col, "") or "").strip()
                if not account_val or account_val in seen_accounts:
                    continue
                seen_accounts.add(account_val)
                clean_owner = re.sub(r'@s\.?whatsapp\.?net$', '', account_val, flags=re.IGNORECASE)
                clean_owner = re.sub(r'[\s\-\.\+]', '', clean_owner).lower()
                if clean_owner:
                    owner_identifiers.add(clean_owner)

        # PRIORITÉ 3: Detecter "(owner)" dans la colonne FROM uniquement.
        # NB: on n'utilise plus _extract_chat_contacts sur Participants (parseur qui
        # mal separait les contacts -> faux positifs owner). Le From a la forme
        # "<username> <display name> (owner)" : on capte <username> via le 1er token
        # "lowercase+symboles" qui precede (owner).
        if from_col:
            seen_owner_froms = set()
            for row in rows:
                from_val = str(row.get(from_col, "") or "")
                if "(owner)" not in from_val.lower():
                    continue
                if from_val in seen_owner_froms:
                    continue
                seen_owner_froms.add(from_val)
                lower = from_val.lower()
                pos = lower.find("(owner)")
                if pos < 0:
                    continue
                prefix = from_val[:pos].rstrip()
                tokens = prefix.split()
                owner_username = None
                for t in tokens:
                    if re.match(r'^[a-z0-9][a-z0-9._\-@]*$', t):
                        owner_username = t
                        break
                if owner_username:
                    owner_identifiers.add(re.sub(r'[\s\-\.\+]', '', owner_username).lower())

        # PRIORITE 4 : "(owner)" dans Participants. Pour chaque occurrence de (owner)
        # on capte le dernier token "username-like" qui le precede.
        if participants_col:
            seen_participants = set()
            for row in rows:
                participants = str(row.get(participants_col, "") or "")
                if "(owner)" not in participants.lower():
                    continue
                if participants in seen_participants:
                    continue
                seen_participants.add(participants)
                lower = participants.lower()
                idx = 0
                while True:
                    pos = lower.find("(owner)", idx)
                    if pos < 0:
                        break
                    prefix = participants[:pos].rstrip()
                    tokens = prefix.split()
                    owner_username = None
                    for t in reversed(tokens):
                        if re.match(r'^[a-z0-9][a-z0-9._\-@]*$', t):
                            owner_username = t
                            break
                    if owner_username:
                        owner_identifiers.add(re.sub(r'[\s\-\.\+]', '', owner_username).lower())
                    idx = pos + len("(owner)")

        # === ÉTAPE 6: NOUVELLE MÉTHODE - Compter par colonne FROM ===
        contact_messages = {}  # {clean_id: {"username": str, "names": set(), "count": int}}

        for row in rows:
            source = str(row.get(source_col, "") or "Natif").strip()
            source_lower = source.lower() if source else ""
            is_natif_source = source_lower in {"natif", "native messages"}

            # Récupérer la valeur de From
            from_value = str(row.get(from_col, "") or "").strip() if from_col else ""

            if not from_value:
                continue

            # Ignorer les System Message
            if "system" in from_value.lower():
                continue

            # Extraire username et nom depuis From
            # Format: "username Nom" ou juste "username"
            tokens = from_value.split()
            if not tokens:
                continue

            username = tokens[0]
            name = " ".join(tokens[1:]) if len(tokens) > 1 else ""

            # Nettoyer l'identifiant
            clean_id = re.sub(r'[\s\-\.\+]', '', username).lower()

            # Normalisation pour sources Natif/Native Messages
            final_id = clean_id
            display_id = username
            if is_natif_source:
                normalized_root, is_phone, is_french = self._normalize_phone_natif(username)
                if is_phone:
                    final_id = normalized_root
                    if is_french:
                        display_id = f"+33{normalized_root}"
                    else:
                        # Numéro étranger: afficher avec + si pas déjà présent
                        display_id = f"+{normalized_root}" if not username.startswith('+') else username
                else:
                    # Pour Native Messages, ignorer les FROM qui ne sont pas des numéros de téléphone
                    # (ex: Doctolib, Google, Apple, WhatsApp, Lyca Mobile, etc.)
                    continue
            # Normalisation pour WhatsApp - retirer @s.whatsapp.net
            elif source_lower == "whatsapp":
                # Retirer le suffixe @s.whatsapp.net (ou @swhatsappnet après nettoyage)
                final_id = re.sub(r'@s\.?whatsapp\.?net$', '', clean_id, flags=re.IGNORECASE)
                # Garder le numéro pour display avec + devant
                phone_only = re.sub(r'@s\.whatsapp\.net$', '', username, flags=re.IGNORECASE)
                display_id = f"+{phone_only}" if not phone_only.startswith('+') else phone_only

            # Exclure l'owner
            if final_id in owner_identifiers:
                continue

            # Compter ce message pour ce contact
            if final_id not in contact_messages:
                contact_messages[final_id] = {
                    "username": display_id,
                    "names": set(),
                    "count": 0
                }

            contact_messages[final_id]["count"] += 1
            clean_name = re.sub(r'_x000d_', '', name, flags=re.IGNORECASE).strip()
            if clean_name and "(owner)" not in clean_name.lower():
                contact_messages[final_id]["names"].add(clean_name)

        # === ÉTAPE 7: Construire le résultat final ===
        # Pour Snapchat: construire le lookup pour récupérer l'identifiant utilisateur
        is_snapchat = source_lower == "snapchat"
        contacts_lookup = self._build_contacts_user_id_lookup(source_value) if is_snapchat else {}

        result = []
        for clean_id, data in contact_messages.items():
            identifier = data["username"]
            names = ", ".join(sorted(data["names"])) if data["names"] else ""

            # Exclure si "(owner)" dans Identifier ou Name
            if "(owner)" in identifier.lower() or "(owner)" in names.lower():
                continue

            # Pour Snapchat: récupérer l'identifiant utilisateur depuis le lookup
            identifiant_utilisateur = ""
            if is_snapchat:
                lookup_info = None
                if identifier and identifier.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[identifier.lower()]
                elif names and names.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[names.lower()]
                if lookup_info:
                    identifiant_utilisateur = lookup_info.get("identifiant_utilisateur") or ""

            result.append({
                "Identifier": identifier,
                "Identifiant_utilisateur": identifiant_utilisateur,
                "Name": names,
                "Nombre_messages": data["count"]
            })

        # Trier par nombre de messages décroissant
        result = sorted(result, key=lambda x: x["Nombre_messages"], reverse=True)

        # Retourner top 15
        return result[:15]

    def get_chats_top15_by_count(self, source_value: str) -> list:
        """Retourne les top 15 contacts par nombre de messages pour une source - VERSION OPTIMISÉE"""
        return self._analyze_chats_fast(source_value)

    def _analyze_chats_fast(self, source_value: str = None) -> list:
        """
        Version OPTIMISÉE de _analyze_chats utilisant les agrégations Polars natives.
        Beaucoup plus rapide pour les gros volumes (31000+ messages).
        """
        import time
        start_time = time.perf_counter()

        lf = self._scan("Chats")
        if lf is None:
            return []

        cols = self.get_columns("Chats")
        source_col = next((c for c in ["Source", "Application", "App", "Platform"] if c in cols), None)
        from_col = "From" if "From" in cols else None
        im_col = "Instant Message #" if "Instant Message #" in cols else None
        label_col = "Label" if "Label" in cols else None

        if not from_col:
            print("[CHATS TOP15 FAST] Colonne 'From' non trouvée")
            return []

        # === Filtres Polars (lazy - très rapide) ===

        # Filtrer Label "Recalled" ou "System"
        if label_col:
            lf = lf.filter(
                pl.col(label_col).is_null() |
                (
                    ~pl.col(label_col).cast(pl.Utf8).str.to_lowercase().str.contains("recalled").fill_null(False) &
                    ~pl.col(label_col).cast(pl.Utf8).str.to_lowercase().str.contains("system").fill_null(False)
                )
            )

        # Filtrer Instant Message # vide
        if im_col:
            lf = lf.filter(pl.col(im_col).is_not_null() & (pl.col(im_col).cast(pl.Utf8).str.strip_chars() != ""))

        # Filtrer sources exclues
        excluded_sources = ['recents', 'interactionc', 'knowledgec', 'threads', 'facebook']
        if source_col:
            lf = lf.filter(~pl.col(source_col).str.to_lowercase().is_in(excluded_sources))
            if source_value:
                lf = lf.filter(pl.col(source_col).str.to_lowercase() == source_value.lower())

        # Filtrer From non vide et non "System"
        lf = lf.filter(
            pl.col(from_col).is_not_null() &
            (pl.col(from_col).cast(pl.Utf8).str.strip_chars() != "") &
            ~pl.col(from_col).cast(pl.Utf8).str.to_lowercase().str.contains("system")
        )

        # === Dédoublonnage (aligné sur _analyze_chats pour cohérence Top 15 / résumé) ===
        dedup_cols = []
        for col_name in ["Start Time: Time", "Last Activity: Time", "Timestamp: Time", "Participants", "Body"]:
            if col_name in cols:
                dedup_cols.append(col_name)
        if label_col:
            dedup_cols.append(label_col)
        for col_name in cols:
            if "Attachment" in col_name:
                dedup_cols.append(col_name)
        if dedup_cols:
            lf = lf.unique(subset=dedup_cols, keep="first")

        # === Agrégation Polars native (ultra rapide) ===
        # Extraire le premier token de From comme identifiant
        agg_df = lf.select([
            pl.col(from_col).cast(pl.Utf8).str.split(" ").list.get(0).alias("username"),
            pl.col(from_col).cast(pl.Utf8).alias("full_from")
        ]).group_by("username").agg([
            pl.count().alias("count"),
            pl.col("full_from").first().alias("sample_from")
        ]).sort("count", descending=True).head(50).collect()  # Prendre 50 pour avoir de la marge après filtrage owner

        # Detection owner via TROIS sources cumulees :
        #  Priorite 1 : feuille User Accounts (colonne Username + autres champs).
        #  Priorite 2 : colonne Account de Chats (si User Accounts est vide).
        #  Priorite 3 : colonne From de Chats (token avant "(owner)").
        # Le set construit est utilise plus bas pour exclure ces identifiants du Top 15.
        owner_identifiers = self._get_owner_usernames_from_user_accounts()

        account_col = "Account" if "Account" in cols else None
        # Priorite 2 : colonne Account (TOUJOURS executee, plus seulement si UA vide).
        # IMPORTANT: on strip "@s.whatsapp.net" AVANT de retirer les points, sinon le
        # suffix devient "@swhatsappnet" et l'ancien regex ne matche plus.
        if account_col:
            account_vals = lf.select(pl.col(account_col).cast(pl.Utf8)).unique().collect().get_column(account_col).to_list()
            for account_val in account_vals:
                account_val = str(account_val or "").strip()
                if account_val:
                    clean_owner = re.sub(r'@s\.?whatsapp\.?net$', '', account_val, flags=re.IGNORECASE)
                    clean_owner = re.sub(r'[\s\-\.\+]', '', clean_owner).lower()
                    if clean_owner:
                        owner_identifiers.add(clean_owner)

        # Priorite 3 : "(owner)" dans From (Chats). Capte le token "username-like"
        # qui precede (owner) dans la chaine From.
        if from_col:
            owner_froms = lf.filter(
                pl.col(from_col).cast(pl.Utf8).str.to_lowercase().str.contains(r"\(owner\)")
            ).select(pl.col(from_col).cast(pl.Utf8).alias("__f")).unique().collect()
            for from_val in owner_froms.get_column("__f").to_list():
                from_val = from_val or ""
                lower = from_val.lower()
                pos = lower.find("(owner)")
                if pos < 0:
                    continue
                prefix = from_val[:pos].rstrip()
                tokens = prefix.split()
                owner_username = None
                for t in tokens:
                    if re.match(r'^[a-z0-9][a-z0-9._\-@]*$', t):
                        owner_username = t
                        break
                if owner_username:
                    owner_identifiers.add(re.sub(r'[\s\-\.\+]', '', owner_username).lower())

        # Priorite 4 : "(owner)" dans Participants (Chats). Pour chaque occurrence
        # de (owner) dans la chaine Participants, on capte le 1er token "username-like"
        # qui le precede. Cumule avec les 3 priorites precedentes.
        participants_col = next((c for c in ["Participants", "Parties", "Contact"] if c in cols), None)
        if participants_col:
            owner_parts = lf.filter(
                pl.col(participants_col).cast(pl.Utf8).str.to_lowercase().str.contains(r"\(owner\)")
            ).select(pl.col(participants_col).cast(pl.Utf8).alias("__p")).unique().collect()
            for participants in owner_parts.get_column("__p").to_list():
                participants = participants or ""
                lower = participants.lower()
                idx = 0
                while True:
                    pos = lower.find("(owner)", idx)
                    if pos < 0:
                        break
                    prefix = participants[:pos].rstrip()
                    tokens = prefix.split()
                    # Reculer pour trouver le dernier token "username-like" avant (owner)
                    owner_username = None
                    for t in reversed(tokens):
                        if re.match(r'^[a-z0-9][a-z0-9._\-@]*$', t):
                            owner_username = t
                            break
                    if owner_username:
                        owner_identifiers.add(re.sub(r'[\s\-\.\+]', '', owner_username).lower())
                    idx = pos + len("(owner)")

        source_lower = source_value.lower() if source_value else ""
        is_natif_source = source_lower in {"natif", "native messages"}
        is_snapchat = source_lower == "snapchat"
        is_phone_source = is_natif_source or source_lower in {"whatsapp", "whatsapp business"}

        # Construire le lookup Contacts pour toutes les sources non-telephoniques
        # (Signal, Snapchat, Instagram, TikTok...) afin de recuperer l'identifiant
        # utilisateur, comme le fait le tableau resume.
        needs_user_id = not is_phone_source
        contacts_lookup = self._build_contacts_user_id_lookup(source_value) if needs_user_id else {}

        result = []
        for row in agg_df.to_dicts():
            username = row.get("username", "") or ""
            count = row.get("count", 0)
            sample_from = row.get("sample_from", "") or ""

            if not username:
                continue

            # Nettoyer l'identifiant
            clean_id = re.sub(r'[\s\-\.\+]', '', username).lower()

            # Extraire le nom depuis sample_from
            tokens = sample_from.split()
            name = " ".join(tokens[1:]) if len(tokens) > 1 else ""
            name = re.sub(r'_x000d_', '', name, flags=re.IGNORECASE).strip()

            # Normalisation selon la source
            final_id = clean_id
            display_id = username

            if is_natif_source:
                normalized_root, is_phone, is_french = self._normalize_phone_natif(username)
                if is_phone:
                    final_id = normalized_root
                    if is_french:
                        display_id = f"+33{normalized_root}"
                    else:
                        display_id = f"+{normalized_root}" if not username.startswith('+') else username
                else:
                    continue  # Ignorer les non-téléphones pour Native Messages
            elif source_lower in ("whatsapp", "whatsapp business"):
                final_id = re.sub(r'@s\.?whatsapp\.?net$', '', clean_id, flags=re.IGNORECASE)
                phone_only = re.sub(r'@s\.whatsapp\.net$', '', username, flags=re.IGNORECASE)
                display_id = f"+{phone_only}" if not phone_only.startswith('+') else phone_only

            # Pour les sources telephoniques: exclure les numeros a moins de 8 chiffres
            # (short-codes operateurs type 38600, alertes bancaires, etc.)
            if is_natif_source or source_lower in ("whatsapp", "whatsapp business"):
                digit_count = sum(c.isdigit() for c in final_id)
                if digit_count < 8:
                    continue

            # Exclure l'owner
            if final_id in owner_identifiers:
                continue

            # Exclure si "(owner)" dans le nom ou identifiant
            if "(owner)" in display_id.lower() or "(owner)" in name.lower():
                continue

            # Recuperer l'identifiant utilisateur via le lookup Contacts
            # (Signal, Snapchat, Instagram, TikTok...). Cle = username (From) puis nom.
            identifiant_utilisateur = ""
            if needs_user_id:
                lookup_info = None
                if display_id and display_id.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[display_id.lower()]
                elif name and name.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[name.lower()]
                if lookup_info:
                    identifiant_utilisateur = lookup_info.get("identifiant_utilisateur") or ""

            result.append({
                "Identifier": display_id,
                "Identifiant_utilisateur": identifiant_utilisateur,
                "Name": name,
                "Nombre_messages": count
            })

            if len(result) >= 15:
                break

        elapsed_ms = (time.perf_counter() - start_time) * 1000
        print(f"[CHATS TOP15 FAST] {source_value}: {len(result)} contacts en {elapsed_ms:.1f}ms")

        return result

    def get_chats_summary(self, source_value: str = None) -> list:
        """Retourne le tableau résumé des chats par source
        Pour Snapchat/Instagram et autres sources sociales: Pseudonyme, Nom d'utilisateur, Identifiant utilisateur, Nombre de messages
        Pour Native Messages/Natif/WhatsApp: Pseudonyme, Numéro de téléphone, Nombre de messages
        Pour Signal: Pseudonyme, Identifiant utilisateur, Nombre de messages (UUID, pas numéro)
        """
        all_contacts = self._analyze_chats(source_value)

        source_lower = source_value.lower() if source_value else ""
        # Sources téléphoniques: format 3 colonnes (Pseudonyme, Numéro, Nombre)
        is_phone_source = source_lower in ("native messages", "natif", "whatsapp", "whatsapp business")
        # Signal: format spécial 2 colonnes d'identité (Pseudonyme, Identifiant utilisateur UUID)
        is_signal = source_lower == "signal"
        # Sources sociales (toutes les autres): format 4 colonnes (Pseudonyme, Nom utilisateur, Identifiant, Nombre)
        is_social_source = not is_phone_source and not is_signal

        # Construire le lookup depuis Contacts pour toutes les sources
        contacts_lookup = self._build_contacts_user_id_lookup(source_value)

        if is_signal:
            # Signal: format spécial (Pseudonyme + Identifiant utilisateur uniquement)
            # L'identifiant Signal est un UUID (ex: 0FD9CF43-C376-4D8E-8E58-108F337966F4 coub)
            aggregated_rows: Dict[str, Dict[str, Any]] = {}

            for contact in all_contacts:
                current_name = contact["Name"] or ""
                current_identifier = contact["Identifier"] or ""  # UUID + pseudonyme pour Signal

                # Parser l'identifiant Signal: "UUID pseudonyme"
                # UUID avec tirets (ex: "0FD9CF43-C376-4D8E-8E58-108F337966F4 coub")
                uuid_match = re.match(r'^([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})\s*(.*)$', current_identifier)
                if uuid_match:
                    identifiant_utilisateur = uuid_match.group(1)
                    pseudonyme_from_identifier = uuid_match.group(2).strip() if uuid_match.group(2) else ""
                else:
                    # Essayer UUID sans tirets (32 caractères hex) et reformater avec tirets
                    uuid_no_dash = re.match(r'^([0-9A-Fa-f]{32})\s*(.*)$', current_identifier)
                    if uuid_no_dash:
                        raw_uuid = uuid_no_dash.group(1)
                        # Reformater: XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
                        identifiant_utilisateur = f"{raw_uuid[:8]}-{raw_uuid[8:12]}-{raw_uuid[12:16]}-{raw_uuid[16:20]}-{raw_uuid[20:]}"
                        pseudonyme_from_identifier = uuid_no_dash.group(2).strip() if uuid_no_dash.group(2) else ""
                    else:
                        # Pas un UUID, utiliser tel quel
                        identifiant_utilisateur = current_identifier or "-"
                        pseudonyme_from_identifier = ""

                # Chercher dans le lookup Contacts par UUID ou name
                lookup_info = None
                if identifiant_utilisateur and identifiant_utilisateur.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[identifiant_utilisateur.lower()]
                elif current_name and current_name.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[current_name.lower()]

                # Déterminer le pseudonyme final
                if lookup_info:
                    pseudonyme = lookup_info.get("pseudonyme") or "-"
                    # Utiliser l'identifiant du lookup si disponible
                    if lookup_info.get("identifiant_utilisateur"):
                        identifiant_utilisateur = lookup_info["identifiant_utilisateur"]
                else:
                    # Pas trouvé dans Contacts: utiliser pseudonyme extrait de l'identifier, sinon Name
                    if pseudonyme_from_identifier:
                        pseudonyme = pseudonyme_from_identifier
                    elif current_name:
                        pseudonyme = current_name
                    else:
                        pseudonyme = "-"

                # Clé d'agrégation pour Signal
                key = f"{pseudonyme}|{identifiant_utilisateur}".lower()

                if key in aggregated_rows:
                    aggregated_rows[key]["Nombre de messages"] += contact["Nombre_messages"]
                else:
                    aggregated_rows[key] = {
                        "Pseudonyme": pseudonyme,
                        "Identifiant utilisateur": identifiant_utilisateur,
                        "Nombre de messages": contact["Nombre_messages"]
                    }

            result = sorted(aggregated_rows.values(), key=lambda x: x["Nombre de messages"], reverse=True)

        elif is_social_source:
            # Pour sources sociales: agréger les doublons après résolution des noms
            aggregated_rows: Dict[str, Dict[str, Any]] = {}

            for contact in all_contacts:
                current_name = contact["Name"] or ""
                current_identifier = contact["Identifier"]

                # Chercher dans le lookup par identifier ou name
                lookup_info = None
                if current_identifier and current_identifier.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[current_identifier.lower()]
                elif current_name and current_name.lower() in contacts_lookup:
                    lookup_info = contacts_lookup[current_name.lower()]

                # Déterminer les valeurs finales
                pseudonyme = "-"
                nom_utilisateur = current_name if current_name else ""
                identifiant_utilisateur = ""

                if lookup_info:
                    # Valeurs depuis Contacts
                    pseudonyme = lookup_info.get("pseudonyme") or "-"
                    if lookup_info.get("nom_utilisateur"):
                        nom_utilisateur = lookup_info["nom_utilisateur"]
                    identifiant_utilisateur = lookup_info.get("identifiant_utilisateur") or ""

                    # Si current_identifier == identifiant_utilisateur du lookup, c'était un ID pas un pseudo
                    if current_identifier and current_identifier.lower() == identifiant_utilisateur.lower():
                        pass  # pseudonyme reste celui du lookup
                    elif current_identifier and pseudonyme and current_identifier.lower() != pseudonyme.lower():
                        # current_identifier n'est ni le pseudo ni l'ID du lookup
                        if not identifiant_utilisateur:
                            identifiant_utilisateur = current_identifier
                else:
                    # Pas trouvé dans Contacts: règle de base pour sources sociales
                    # Name → Pseudonyme (carnet d'adresses), Identifier → Nom d'utilisateur (username app)
                    pseudonyme = current_name if current_name else "-"
                    nom_utilisateur = current_identifier
                    # Vérifier si nom_utilisateur ressemble à un ID (UUID ou numérique long)
                    if nom_utilisateur and self._looks_like_user_id(nom_utilisateur):
                        identifiant_utilisateur = nom_utilisateur
                        nom_utilisateur = "-"

                # Clé d'agrégation
                key = f"{pseudonyme}|{nom_utilisateur}|{identifiant_utilisateur}".lower()

                if key in aggregated_rows:
                    # Agréger les compteurs
                    aggregated_rows[key]["Nombre de messages"] += contact["Nombre_messages"]
                else:
                    # Sources sociales: 3 colonnes d'identité
                    aggregated_rows[key] = {
                        "Pseudonyme": pseudonyme,
                        "Nom d'utilisateur": nom_utilisateur,
                        "Identifiant utilisateur": identifiant_utilisateur,
                        "Nombre de messages": contact["Nombre_messages"]
                    }

            # Trier par nombre de messages et retourner
            result = sorted(aggregated_rows.values(), key=lambda x: x["Nombre de messages"], reverse=True)
        else:
            # Pour les sources téléphoniques: 3 colonnes (Pseudonyme, Numéro, Nombre)
            result = []
            for contact in all_contacts:
                current_name = contact["Name"] or "Inconnu"
                current_identifier = contact["Identifier"]

                # Nettoyer le numéro de téléphone pour WhatsApp Business
                # Format entrant: "590690196318@s.whatsapp.net" → "+590690196318"
                phone_number = current_identifier
                if phone_number:
                    # Supprimer le suffixe @s.whatsapp.net ou similaire
                    if "@" in phone_number:
                        phone_number = phone_number.split("@")[0]
                    # Ajouter le + si c'est un numéro et qu'il n'en a pas déjà
                    if phone_number and phone_number[0].isdigit():
                        phone_number = "+" + phone_number

                # Exclure les numeros a moins de 8 chiffres (short-codes operateurs,
                # alertes, numeros non resolus type +0). Aligne le resume sur le Top 15.
                if sum(c.isdigit() for c in (phone_number or "")) < 8:
                    continue

                row = {
                    "Pseudonyme": current_name,
                    "Numéro de téléphone": phone_number,
                    "Nombre de messages": contact["Nombre_messages"]
                }
                result.append(row)

        return result

# GESTION DU STOCKAGE


def cleanup_import(import_id: str) -> bool:
    """Supprime un import et ses fichiers Parquet"""
    import_path = DATA_DIR / import_id
    if import_path.exists():
        shutil.rmtree(import_path)
        return True
    return False


def list_imports() -> List[Dict[str, Any]]:
    """Liste tous les imports disponibles"""
    imports = []
    for path in DATA_DIR.iterdir():
        if path.is_dir():
            parquet_count = len(list(path.glob("*.parquet")))
            imports.append({
                "import_id": path.name,
                "path": str(path),
                "parquet_count": parquet_count
            })
    return imports


def get_import_path(import_id: str) -> Optional[Path]:
    """Retourne le chemin d'un import"""
    path = DATA_DIR / import_id
    return path if path.exists() else None

