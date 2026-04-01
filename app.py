# -*- coding: utf-8 -*-
"""
Générateur de cartels musée
"""

import io
import re
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# -------------------------
# CONFIG
# -------------------------
st.set_page_config(page_title="Générateur de cartels", page_icon="🖼️")

st.title("🖼️ Générateur de cartels")

uploaded = st.file_uploader("Importer un fichier Excel", type=["xlsx"])


# -------------------------
# CONSTANTES
# -------------------------
ALL_FIELDS = [
    "Auteur / Exécutant",
    "Date auteur/exécutant",
    "Editeur",
    "Titre",
    "Provenance",
    "Date",
    "Technique(s) de l'œuvre originale",
    "Mention Collection pour le cartel",
    "Information sur l'acquisition",
]

REQUIRED_FIELDS = [
    "Titre",
    "Technique(s) de l'œuvre originale",
    "Mention Collection pour le cartel",
]

EXPECTED_SHEETS = ["Liste", "Repro"]


# -------------------------
# OUTILS
# -------------------------
def add_horizontal_rule(doc):
    p = doc.add_paragraph()
    p_par = p._p
    pPr = p_par.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "auto")
    pBdr.append(bottom)
    pPr.append(pBdr)


def clean_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", str(name).strip())


def make_unique_columns(columns):
    seen = {}
    new_cols = []
    for col in columns:
        col = str(col).strip()
        if col in seen:
            seen[col] += 1
            new_cols.append(f"{col}__{seen[col]}")
        else:
            seen[col] = 1
            new_cols.append(col)
    return new_cols


def normalize_columns(df):
    df.columns = make_unique_columns(df.columns)
    return df


def safe(val):
    if val is None:
        return ""
    if isinstance(val, pd.Series):
        return " | ".join([str(v) for v in val if pd.notna(v)])
    if pd.isna(val):
        return ""
    return str(val).strip()


def get_cell_value(row, col):
    return safe(row[col]) if col in row.index else ""


def read_sheet(file, sheet):
    df = pd.read_excel(file, sheet_name=sheet, header=1)
    return normalize_columns(df)


def drop_empty_rows(df):
    mask = df[REQUIRED_FIELDS].apply(lambda r: any(safe(v) != "" for v in r), axis=1)
    return df[mask].reset_index(drop=True)


# -------------------------
# FORMAT DATE (FIX BUG)
# -------------------------
def add_date_with_superscript(paragraph, text, font_size=Pt(10.5)):
    """
    Met en exposant le 'e' uniquement pour les siècles (XIXe, XXe…)
    sans toucher au 'e' de 'siècle'
    """
    pattern = re.compile(r'(?i)\b([IVXLCDM]+)(e)(?=\s|[-–—,;:.!?)]|$)')
    last_idx = 0

    for match in pattern.finditer(text):
        start, end = match.span()

        if start > last_idx:
            run = paragraph.add_run(text[last_idx:start])
            run.font.size = font_size

        # chiffres romains
        run1 = paragraph.add_run(match.group(1))
        run1.font.size = font_size

        # "e" en exposant
        run2 = paragraph.add_run(match.group(2))
        run2.font.size = font_size
        run2.font.superscript = True

        last_idx = end

    if last_idx < len(text):
        run = paragraph.add_run(text[last_idx:])
        run.font.size = font_size


# -------------------------
# CARTEL
# -------------------------
def add_cartel(doc, row, fields):
    titre = get_cell_value(row, "Titre")

    auteur = get_cell_value(row, "Auteur / Exécutant")
    date_auteur = get_cell_value(row, "Date auteur/exécutant")
    editeur = get_cell_value(row, "Editeur")
    provenance = get_cell_value(row, "Provenance")
    date = get_cell_value(row, "Date")
    technique = get_cell_value(row, "Technique(s) de l'œuvre originale")
    mention = get_cell_value(row, "Mention Collection pour le cartel")
    acquisition = get_cell_value(row, "Information sur l'acquisition")

    # Titre
    p = doc.add_paragraph()
    r = p.add_run(titre)
    r.bold = True
    r.font.size = Pt(14)

    # Auteur + date auteur
    parts = []
    if "Auteur / Exécutant" in fields and auteur:
        parts.append(auteur)
    if "Date auteur/exécutant" in fields and date_auteur:
        parts.append(date_auteur)

    if parts:
        p = doc.add_paragraph(", ".join(parts))

    # Editeur
    if "Editeur" in fields and editeur:
        doc.add_paragraph(editeur)

    # Provenance
    if "Provenance" in fields and provenance:
        doc.add_paragraph(provenance)

    # Date (avec exposant)
    if "Date" in fields and date:
        p = doc.add_paragraph()
        add_date_with_superscript(p, date)

    # Technique
    if "Technique(s) de l'œuvre originale" in fields and technique:
        doc.add_paragraph(technique)

    # Mention
    if "Mention Collection pour le cartel" in fields and mention:
        doc.add_paragraph(mention)

    # Acquisition
    if "Information sur l'acquisition" in fields and acquisition:
        doc.add_paragraph(acquisition)


# -------------------------
# DOC
# -------------------------
def create_doc(data, fields, title):
    doc = Document()

    for sec in doc.sections:
        sec.top_margin = Cm(2)
        sec.bottom_margin = Cm(2)

    doc.add_paragraph(title)

    for i, (_, row) in enumerate(data.iterrows()):
        add_cartel(doc, row, fields)

        if i < len(data) - 1:
            doc.add_paragraph()
            add_horizontal_rule(doc)
            doc.add_paragraph()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# -------------------------
# APP
# -------------------------
if uploaded:

    xl = pd.ExcelFile(uploaded)
    sheets = [s for s in EXPECTED_SHEETS if s in xl.sheet_names]

    selected_sheet = st.selectbox("Onglet", sheets)

    df = read_sheet(uploaded, selected_sheet)

    fields = st.multiselect(
        "Champs",
        ALL_FIELDS,
        default=REQUIRED_FIELDS
    )

    name = st.text_input("Nom fichier")
    title = st.text_input("Titre doc", value="Cartels")

    if st.button("Générer"):
        df = drop_empty_rows(df)

        buffer = create_doc(df, fields, title)

        st.download_button(
            "Télécharger",
            buffer,
            file_name=f"{clean_filename(name)}.docx"
        )
