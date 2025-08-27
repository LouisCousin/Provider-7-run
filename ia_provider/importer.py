from __future__ import annotations
import logging
from typing import Any, Dict, List, Optional, Tuple, Union

import docx
from docx.document import Document as DocumentObject
from docx.section import _Header, _Footer
from docx.opc.exceptions import OpcError
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
import fitz  # PyMuPDF

# Configuration de la journalisation
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def _extraire_style_run(run) -> Dict[str, Any]:
    """Extrait les informations de style d'un segment de texte (run)."""
    font = run.font
    color = font.color.rgb if font.color and font.color.rgb else None
    return {
        "text": run.text,
        "style": {
            "font_name": font.name,
            "font_size": font.size.pt if font.size else None,
            "is_bold": font.bold,
            "is_italic": font.italic,
            "font_color_rgb": str(color) if color else None,
        },
    }


def _analyser_contenu_block(parent: Union[DocumentObject, _Header, _Footer, _Cell]) -> List[Dict[str, Any]]:
    """Analyse un conteneur (document, header, cell, etc.) et retourne la structure des blocs."""

    # Utilise une fonction interne pour itérer sur les paragraphes et tableaux
    def iter_block_items(parent_item):
        if isinstance(parent_item, _Cell):
            parent_element = parent_item._tc
        else:
            parent_element = parent_item.element.body

        for child in parent_element.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent_item)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent_item)

    contenu_structure: List[Dict[str, Any]] = []
    for block in iter_block_items(parent):
        if isinstance(block, Paragraph):
            if not block.text.strip():
                continue

            style_name = block.style.name.lower() if block.style and block.style.name else ""

            # Gestion des listes
            if "list" in style_name or "liste" in style_name:
                # Si le dernier bloc était déjà une liste, on y ajoute l'item
                if contenu_structure and contenu_structure[-1]["type"] == "list":
                    contenu_structure[-1]["items"].append(block.text)
                else:
                    contenu_structure.append({"type": "list", "items": [block.text]})
                continue

            # Gestion des titres et paragraphes
            block_type = "paragraph"
            if style_name.startswith("heading 1") or style_name.startswith("titre 1"):
                block_type = "heading_1"
            elif style_name.startswith("heading 2") or style_name.startswith("titre 2"):
                block_type = "heading_2"
            # ... (ajouter d'autres niveaux de titre si nécessaire)

            runs_data = [_extraire_style_run(run) for run in block.runs if run.text.strip()]
            if runs_data:
                contenu_structure.append({"type": block_type, "runs": runs_data})

        elif isinstance(block, Table):
            table_data: List[List[Dict[str, Any]]] = []
            for row in block.rows:
                row_data = [_analyser_contenu_block(cell) for cell in row.cells]
                table_data.append(row_data)
            if table_data:
                contenu_structure.append({"type": "table", "rows": table_data})

    return contenu_structure


def analyser_docx(
    file_stream,
) -> Tuple[Dict[str, List[Dict[str, Any]]], None]:
    """Extrait le contenu structuré d'un DOCX, y compris en-têtes et pieds de page."""
    try:
        file_stream.seek(0)
        document = docx.Document(file_stream)

        # 1. Analyser le corps du document
        corps_structure = _analyser_contenu_block(document)

        # 2. Analyser l'en-tête et le pied de page (simplifié à la première section)
        header_structure = []
        footer_structure = []
        if document.sections:
            section = document.sections[0]
            if section.header:
                header_structure = _analyser_contenu_block(section.header)
            if section.footer:
                footer_structure = _analyser_contenu_block(section.footer)

        document_complet = {
            "header": header_structure,
            "body": corps_structure,
            "footer": footer_structure,
        }
        return document_complet, None

    except OpcError as e:
        logging.error(f"Fichier DOCX corrompu : {e}")
        return {"header": [], "body": [], "footer": []}, None
    except Exception as e:
        logging.error(f"Erreur inattendue sur DOCX : {e}", exc_info=True)
        return {"header": [], "body": [], "footer": []}, None


def analyser_pdf(file_stream) -> Tuple[str, None]:
    """Extrait le contenu textuel brut d'un PDF."""
    try:
        file_stream.seek(0)
        with fitz.open(stream=file_stream.read(), filetype="pdf") as doc:
            full_text = "".join(page.get_text() for page in doc)
        return full_text, None
    except Exception as e:
        logging.error(f"Erreur inattendue sur PDF : {e}", exc_info=True)
        return "", None


def analyser_document(
    fichier,
) -> Tuple[Union[str, Dict[str, List[Dict[str, Any]]]], None]:
    """Analyse un fichier importé et choisit la méthode appropriée."""
    filename = fichier.name.lower()
    if filename.endswith(".docx"):
        return analyser_docx(fichier)
    if filename.endswith(".pdf"):
        return analyser_pdf(fichier)
    return "", None


def extraire_texte_de_structure(document_structure: Dict[str, List[Dict[str, Any]]]) -> str:
    """Extrait et concatène tout le texte d'une structure de document."""
    logging.info("Début de l'extraction de texte depuis la structure.")

    nb_blocs_body = len(document_structure.get("body", []))
    logging.info(
        f"Nombre de blocs principaux trouvés dans le corps : {nb_blocs_body}"
    )

    if nb_blocs_body == 0:
        logging.warning(
            "Aucun bloc de contenu trouvé dans la structure du document. L'extraction retournera une chaîne vide."
        )

    texte_complet: List[str] = []

    def extraire_runs(blocs: List[Dict[str, Any]], niveau: int = 0) -> None:
        for i, bloc in enumerate(blocs):
            type_bloc = bloc.get("type", "inconnu")
            logging.info(
                f"{'  ' * niveau}Traitement du bloc {i+1}/{len(blocs)} de type : {type_bloc}"
            )

            if type_bloc == "table":
                for row in bloc.get("rows", []):
                    for cell in row:
                        extraire_runs(cell, niveau + 1)
            elif type_bloc == "list":
                items = bloc.get("items", [])
                logging.info(
                    f"{'  ' * (niveau+1)}-> Trouvé {len(items)} élément(s) de liste."
                )
                for item in items:
                    if item.strip():
                        texte_complet.append(item)
            elif bloc.get("runs"):
                paragraphe = "".join(
                    run.get("text", "") for run in bloc.get("runs", [])
                )
                if paragraphe.strip():
                    logging.info(
                        f"{'  ' * (niveau+1)}-> Ajout de {len(paragraphe)} caractères."
                    )
                    texte_complet.append(paragraphe)
            else:
                logging.warning(
                    f"{'  ' * (niveau+1)}-> Bloc de type '{type_bloc}' ignoré car il ne contient ni 'runs', ni 'items', ni n'est une table."
                )

    extraire_runs(document_structure.get("header", []))
    extraire_runs(document_structure.get("body", []))
    extraire_runs(document_structure.get("footer", []))

    texte_final = "\n\n".join(texte_complet)
    logging.info(
        f"Extraction terminée. Longueur totale du texte : {len(texte_final)} caractères."
    )
    return texte_final
