"""Rendu des lettres de mission : remplit un template prepare (docxtpl)
avec les valeurs saisies dans le formulaire et retourne le .docx en memoire.
"""

import io

from docxtpl import DocxTemplate

from generator.missions import Mission


def build_context(mission: Mission, raw_values: dict) -> dict:
    """Transforme les valeurs brutes du formulaire en contexte Jinja.

    - Les dates sont converties en texte au format jour/mois/annee en toutes
      lettres n'est pas fait ici (laisse a l'utilisateur, champ texte libre
      pour date_de_cloture) ; seule `aujourdhui`, saisie via un date picker,
      est formatee en français.
    - `mention_intermediaire` (case a cocher) devient le texte "ou
      intermédiaires" ou une chaine vide.
    - `raison_sociale`, si laissee vide, reprend `nom_du_dossier`.
    """
    context = dict(raw_values)

    if "aujourdhui" in context and hasattr(context["aujourdhui"], "strftime"):
        context["aujourdhui"] = format_date_fr(context["aujourdhui"])

    if "mention_intermediaire" in context:
        context["mention_intermediaire"] = (
            "ou intermédiaires" if context["mention_intermediaire"] else ""
        )

    if "raison_sociale" in context and not context["raison_sociale"]:
        context["raison_sociale"] = context.get("nom_du_dossier", "")

    return context


MOIS_FR = [
    "janvier", "février", "mars", "avril", "mai", "juin",
    "juillet", "août", "septembre", "octobre", "novembre", "décembre",
]


def format_date_fr(d) -> str:
    return f"{d.day} {MOIS_FR[d.month - 1]} {d.year}"


def render_letter(mission: Mission, raw_values: dict) -> io.BytesIO:
    """Genere la lettre de mission et retourne un buffer .docx pret a
    telecharger."""
    context = build_context(mission, raw_values)
    doc = DocxTemplate(str(mission.template_path))
    doc.render(context)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
