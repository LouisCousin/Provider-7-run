from __future__ import annotations

"""Module d'analyse de documents pour l'import."""

from typing import Any, Dict, List, Optional, Tuple, Union

import logging

import docx
from docx.opc.exceptions import OpcError
import fitz  # PyMuPDF
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

# Configuration simple pour la journalisation
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def _extraire_style_run(run) -> Dict[str, Any]:
    """Extrait les informations de style d'un segment de texte (run)."""

    font = run.font
    couleur_rgb = font.color.rgb if font.color and font.color.rgb else None
    return {
        "font_name": font.name,
        "font_size": font.size.pt if font.size else None,
        "is_bold": font.bold,
        "is_italic": font.italic,
        "font_color_rgb": str(couleur_rgb) if couleur_rgb else None,
    }


def analyser_docx(
    file_stream,
) -> Tuple[Dict[str, Any], Optional[Dict[str, Any]]]:
    """Extrait le contenu structuré d'un DOCX avec les styles associés.

    Retourne ``(structure, None)`` où ``structure`` contient trois clés :
    ``body`` pour la structure principale du document, ``header`` et ``footer``
    pour les textes des en-têtes et pieds de page.
    """
    try:
        file_stream.seek(0)
        document = docx.Document(file_stream)

        def iter_block_items(parent):
            """Yield paragraph and table objects in *parent* in document order."""
            for child in parent.element.body.iterchildren():
                if isinstance(child, CT_P):
                    yield Paragraph(child, parent)
                elif isinstance(child, CT_Tbl):
                    yield Table(child, parent)

        contenu_structure: List[Dict[str, Any]] = []
        for block in iter_block_items(document):
            if isinstance(block, Paragraph):
                style_name = (
                    block.style.name.lower() if block.style and block.style.name else ""
                )
                if "list" in style_name or "liste" in style_name:
                    if block.text.strip():
                        if contenu_structure and contenu_structure[-1]["type"] == "list":
                            contenu_structure[-1]["items"].append(block.text)
                        else:
                            contenu_structure.append({"type": "list", "items": [block.text]})
                    continue

                block_type = "paragraph"
                if style_name.startswith("heading 1") or style_name.startswith("titre 1"):
                    block_type = "heading_1"
                elif style_name.startswith("heading 2") or style_name.startswith("titre 2"):
                    block_type = "heading_2"
                elif style_name.startswith("heading 3") or style_name.startswith("titre 3"):
                    block_type = "heading_3"
                elif style_name.startswith("heading 4") or style_name.startswith("titre 4"):
                    block_type = "heading_4"
                elif style_name.startswith("heading 5") or style_name.startswith("titre 5"):
                    block_type = "heading_5"
                elif style_name.startswith("heading 6") or style_name.startswith("titre 6"):
                    block_type = "heading_6"

                if block.text.strip():
                    runs_data = []
                    for run in block.runs:
                        if run.text.strip():
                            runs_data.append(
                                {
                                    "text": run.text,
                                    "style": _extraire_style_run(run),
                                }
                            )

                    if runs_data:
                        contenu_structure.append({"type": block_type, "runs": runs_data})

            elif isinstance(block, Table):
                table_data: List[List[str]] = []
                for row in block.rows:
                    row_data = [cell.text for cell in row.cells]
                    table_data.append(row_data)
                if table_data:
                    contenu_structure.append({"type": "table", "rows": table_data})

        contenu_header = ""
        contenu_footer = ""
        for section in document.sections:
            if section.header:
                contenu_header += "\n".join(
                    p.text for p in section.header.paragraphs if p.text
                )
            if section.footer:
                contenu_footer += "\n".join(
                    p.text for p in section.footer.paragraphs if p.text
                )

        return {
            "body": contenu_structure,
            "header": contenu_header,
            "footer": contenu_footer,
        }, None

    except OpcError as e:
        logging.error(
            f"Erreur de parsing du fichier DOCX (potentiellement corrompu) : {e}"
        )
        return {"body": [], "header": "", "footer": ""}, None
    except Exception as e:  # Garde un filet de sécurité
        logging.error(
            f"Erreur inattendue lors de l'analyse du DOCX : {e}", exc_info=True
        )
        return {"body": [], "header": "", "footer": ""}, None


def analyser_pdf(file_stream) -> Tuple[str, None]:
    """Extrait le contenu textuel brut d'un PDF."""
    file_stream.seek(0)
    with fitz.open(stream=file_stream.read(), filetype="pdf") as doc:
        full_text = "".join(page.get_text() for page in doc)
    return full_text, None


def analyser_document(
    fichier,
) -> Tuple[Union[str, Dict[str, Any]], Optional[Dict[str, Any]]]:
    """Analyse un fichier importé et choisit la méthode appropriée."""
    filename = fichier.name.lower()
    if filename.endswith(".docx"):
        return analyser_docx(fichier)
    if filename.endswith(".pdf"):
        return analyser_pdf(fichier)
    return "", None
