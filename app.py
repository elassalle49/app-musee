# -*- coding: utf-8 -*-
"""
Created on Thu Nov  6 10:43:56 2025
@author: EvaLa
"""

import io
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- Configuration de la page ---
st.set_page_config(page_title="Générateur de cartels", page_icon="🖼️", layout="centered")

st.title("🖼️ Générateur de cartels")
st.write("Cette interface permet de générer les cartels à partir de votre fichier Excel qui répertorie les œuvres d'art d'une exposition.")

# Légende centrée au-dessus de l’uploader
st.markdown(
    """
    <h4 style='text-align: center;'>📂 Insérer votre fichier Excel</h4>
    """,
    unsafe_allow_html=True
)

# --- Uploader du fichier Excel ---
uploaded = st.file_uploader("", type=["xlsx", "xls"])

# --- Fonctions utilitaires ---
def add_horizontal_rule(doc):
    """Ajoute une ligne horizontale comme séparateur."""
    p = doc.add_paragraph()
    p_par = p._p
    pPr = p_par.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')  # style de ligne
    bottom.set(qn('w:sz'), '12')       # épaisseur
    bottom.set(qn('w:space'), '1')     # espace
    bottom.set(qn('w:color'), 'auto')  # couleur
    pBdr.append(bottom)
    pPr.append(pBdr)

def safe(val):
    return "" if pd.isna(val) else str(val)

required_cols = ["Titre de l'œuvre", "Artiste", "Date de création", "Description"]

# --- Si un fichier a été chargé ---
if uploaded:
    try:
        df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Erreur de lecture du fichier Excel : {e}")
        st.stop()

    if df.empty:
        st.warning("Le fichier est vide.")
        st.stop()

    st.subheader("Aperçu des données")
    st.dataframe(df.head(10), use_container_width=True)

    # Vérification des colonnes nécessaires
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Colonnes manquantes : {', '.join(missing)}")
        st.stop()

    # --- Options (marge fixe + nom obligatoire) ---
    st.markdown(
        "<div style='font-weight:600; margin-bottom:4px;'>"
        "Nom du fichier Word (sans extension) <span style='color:#d00'>*</span>"
        "</div>",
        unsafe_allow_html=True
    )
    nom_fichier = st.text_input(
        label="",  # on masque le label natif
        placeholder="Indiquer le nom du document",
        key="nom_fichier_input",
        label_visibility="collapsed"
    )
    marge_cm = 2.0  # marge fixe

    # --- Bouton de génération ---
    if st.button("🪄 Transformer"):
        # Validation : nom de fichier obligatoire
        if not nom_fichier.strip():
            st.error("Veuillez indiquer le nom du document.")
            # Mise en rouge du champ
            st.markdown("""
                <style>
                div[data-testid="stTextInput"] input {
                    border: 1px solid #d00 !important;
                    box-shadow: 0 0 0 1px rgba(221,0,0,.25) !important;
                }
                </style>
            """, unsafe_allow_html=True)
            st.stop()

        # Création du document Word
        doc = Document()
        doc.core_properties.title = "Cartels - Expo A"

        for sec in doc.sections:
            sec.top_margin = Cm(marge_cm)
            sec.bottom_margin = Cm(marge_cm)
            sec.left_margin = Cm(marge_cm)
            sec.right_margin = Cm(marge_cm)

        # Titre principal
        title_p = doc.add_paragraph()
        title_run = title_p.add_run("Cartels - Expo A")
        title_run.bold = True
        title_run.font.size = Pt(20)
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()

        # Boucle sur les œuvres
        for _, row in df.iterrows():
            titre = safe(row["Titre de l'œuvre"])
            artiste = safe(row["Artiste"])
            date = safe(row["Date de création"])
            desc = safe(row["Description"])

            # Titre
            p_titre = doc.add_paragraph()
            r_titre = p_titre.add_run(titre if titre else "Sans titre")
            r_titre.bold = True
            r_titre.font.size = Pt(14)

            # Artiste + Date
            p_meta = doc.add_paragraph()
            r_meta = p_meta.add_run(f"{artiste} — {date}".strip(" —"))
            r_meta.italic = True
            r_meta.font.size = Pt(11)

            # Description
            if desc:
                p_desc = doc.add_paragraph(desc)
                for r in p_desc.runs:
                    r.font.size = Pt(11)

            # Séparateur
            doc.add_paragraph()
            add_horizontal_rule(doc)
            doc.add_paragraph()

        # Sauvegarde en mémoire
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.success("✅ Document généré avec succès !")
        st.download_button(
            label="📥 Télécharger le fichier Word",
            data=buffer,
            file_name=f"{nom_fichier}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
