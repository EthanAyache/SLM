"""Générateur de lettres de mission — formulaire de saisie et génération
du fichier Word correspondant.

Lancement : streamlit run app.py
"""

import streamlit as st
from dotenv import load_dotenv

from generator.docx_engine import render_letter
from generator.missions import MISSIONS
from generator.pappers_client import (
    CleApiInvalide,
    PappersError,
    QuotaDepasse,
    SirenIntrouvable,
    extract_form_data,
    fetch_entreprise,
)

load_dotenv()

st.set_page_config(page_title="Lettres de mission", page_icon="📄")

st.title("Générateur de lettres de mission")

mission_label_to_key = {m.label: key for key, m in MISSIONS.items()}
mission_label = st.selectbox("Type de mission", list(mission_label_to_key.keys()))
mission = MISSIONS[mission_label_to_key[mission_label]]

st.caption(
    "Renseignez les informations ci-dessous puis cliquez sur "
    "\"Générer la lettre\" pour télécharger le fichier Word."
)

# --- Pré-remplissage via l'API Pappers (hors formulaire : un st.form ne se
# rafraîchit qu'à la soumission, donc la recherche doit être en dehors pour
# pouvoir préremplir les champs avant que l'utilisateur les remplisse). ---
st.subheader("Pré-remplissage automatique via Pappers (optionnel)")
col_siren, col_button = st.columns([3, 1])
with col_siren:
    siren_input = st.text_input("Numéro SIREN", key="siren_input", max_chars=9)
with col_button:
    st.write("")
    st.write("")
    rechercher = st.button("Rechercher via Pappers")

if rechercher:
    try:
        raw = fetch_entreprise(siren_input)
        data = extract_form_data(raw)
        candidats = data.pop("representants_candidats")
        for field_name, value in data.items():
            st.session_state[field_name] = value
        if len(candidats) == 1:
            c = candidats[0]
            st.session_state["civilite_representant"] = c["civilite_representant"]
            st.session_state["nom_representant"] = c["nom_representant"]
            st.session_state["fonction_representant"] = c["fonction_representant"]
            st.session_state.pop("_representants_candidats", None)
        elif len(candidats) > 1:
            st.session_state["_representants_candidats"] = candidats
        else:
            st.session_state.pop("_representants_candidats", None)
        st.success(f"Informations récupérées pour « {data['nom_du_dossier']} ».")
    except SirenIntrouvable as e:
        st.error(str(e))
    except CleApiInvalide as e:
        st.error(str(e))
    except QuotaDepasse as e:
        st.error(str(e))
    except PappersError as e:
        st.error(str(e))

candidats = st.session_state.get("_representants_candidats")
if candidats:
    labels = [c["label"] for c in candidats]
    choix = st.selectbox(
        "Plusieurs dirigeants trouvés — choisir le représentant pour la lettre",
        labels,
        key="choix_representant",
    )
    chosen = next(c for c in candidats if c["label"] == choix)
    st.session_state["civilite_representant"] = chosen["civilite_representant"]
    st.session_state["nom_representant"] = chosen["nom_representant"]
    st.session_state["fonction_representant"] = chosen["fonction_representant"]

st.divider()

with st.form(key=f"form_{mission.key}"):
    raw_values = {}
    for f in mission.fields:
        widget_label = f.label + ("" if f.required else " (optionnel)")
        if f.name not in st.session_state:
            st.session_state[f.name] = f.default
        if f.type == "select":
            raw_values[f.name] = st.selectbox(widget_label, f.options, key=f.name, help=f.help or None)
        elif f.type == "checkbox":
            raw_values[f.name] = st.checkbox(widget_label, key=f.name, help=f.help or None)
        elif f.type == "date":
            raw_values[f.name] = st.date_input(widget_label, key=f.name, help=f.help or None)
        elif f.type == "textarea":
            raw_values[f.name] = st.text_area(widget_label, key=f.name, help=f.help or None)
        else:
            raw_values[f.name] = st.text_input(widget_label, key=f.name, help=f.help or None)

    submitted = st.form_submit_button("Générer la lettre")

if submitted:
    missing = [
        f.label for f in mission.fields
        if f.required and not str(raw_values.get(f.name, "")).strip()
    ]
    if missing:
        st.error("Champs obligatoires manquants : " + ", ".join(missing))
    else:
        buffer = render_letter(mission, raw_values)
        filename = f"Lettre de mission - {raw_values.get('nom_du_dossier', 'client')} - {mission.label}.docx"
        st.success("Lettre générée.")
        st.download_button(
            label="Télécharger le fichier Word",
            data=buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
