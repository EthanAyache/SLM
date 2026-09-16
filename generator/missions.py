"""Configuration declarative des types de mission geres par l'application.

Chaque mission pointe vers un template prepare (templates/<key>.docx, tags
Jinja {{ champ }}) et declare la liste des champs a afficher dans le
formulaire, dans l'ordre d'affichage souhaite.
"""

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates"

CIVILITE_OPTIONS = ["Monsieur", "Madame"]
FORME_JURIDIQUE_OPTIONS = [
    "SAS", "SASU", "SARL", "EURL", "SA", "SCI", "SNC", "EI", "Association", "Autre",
]
PERIODE_REVISION_OPTIONS = ["Janvier à Avril", "Février à Mai", "Mars à Juin"]
PERIODE_LIASSE_OPTIONS = ["Avril-Mai", "Mai-Juin", "Juin-Juillet"]
PERIODE_JURIDIQUE_OPTIONS = ["Avril-Mai-Juin", "Mai-Juin-Juillet", "Juin-Juillet-Août"]


@dataclass
class FieldDef:
    name: str  # doit correspondre au tag {{ name }} dans le template
    label: str
    type: str = "text"  # text | textarea | select | checkbox | date
    options: list = field(default_factory=list)
    default: object = ""
    required: bool = True
    help: str = ""


# --- Blocs de champs reutilisables entre missions -------------------------

CLIENT_FIELDS = [
    FieldDef("nom_du_dossier", "Nom du dossier / client"),
    FieldDef("forme_juridique", "Forme juridique", type="select", options=FORME_JURIDIQUE_OPTIONS),
    FieldDef("adresse1", "Adresse (ligne 1)"),
    FieldDef("adresse2", "Adresse (ligne 2)", required=False),
    FieldDef("code_postal", "Code postal"),
    FieldDef("ville", "Ville"),
]

REPRESENTANT_FIELDS = [
    FieldDef("civilite_representant", "Civilité du représentant", type="select", options=CIVILITE_OPTIONS),
    FieldDef("nom_representant", "Nom du représentant"),
    FieldDef("fonction_representant", "Fonction du représentant", help="Ex: Président, Gérant"),
]

SIGNATURE_FIELDS = [
    FieldDef("lieu_signature", "Lieu de signature", default="Paris"),
    FieldDef("aujourdhui", "Date de la lettre", type="date", default=date.today()),
    FieldDef("expert_comptable", "Expert-comptable signataire"),
]

# Variante sans "expert_comptable" : ce champ n'existe pas dans les
# templates procedures_convenues et procedure_collective.
SIGNATURE_FIELDS_SANS_EC = [
    FieldDef("lieu_signature", "Lieu de signature", default="Paris"),
    FieldDef("aujourdhui", "Date de la lettre", type="date", default=date.today()),
]

CLOTURE_FIELD = FieldDef("date_de_cloture", "Date de clôture d'exercice", help="Ex: 31 décembre 2025")

HONORAIRES_CTA_FIELD = FieldDef("hono_cta", "Honoraires tenue/révision (€)", required=False)
HONORAIRES_JUR_FIELD = FieldDef("hono_jur", "Honoraires juridique (€)", required=False)

PERIODE_FIELDS = [
    FieldDef("periode_revision", "Période révision des comptes", type="select", options=PERIODE_REVISION_OPTIONS),
    FieldDef("periode_liasse", "Période établissement comptes annuels / liasses", type="select", options=PERIODE_LIASSE_OPTIONS),
    FieldDef("periode_juridique", "Période secrétariat juridique", type="select", options=PERIODE_JURIDIQUE_OPTIONS),
    FieldDef(
        "mention_intermediaire",
        "Ajouter la mention \"ou intermédiaires\"",
        type="checkbox",
        default=False,
        required=False,
    ),
]

RAISON_SOCIALE_FIELD = FieldDef(
    "raison_sociale", "Raison sociale (si différente du nom du dossier)", required=False,
    help="Laisser vide pour reprendre le nom du dossier",
)

# Champs specifiques a la mission Sociale (assistance en gestion sociale) :
# SIREN, activite (libelle NAF) et convention collective, pre-remplissables
# via la recherche Pappers au meme titre que les autres champs client.
SOCIALE_FIELDS = [
    FieldDef("siren", "Numéro SIREN"),
    FieldDef("naf_lib", "Activité principale (libellé code NAF/APE)", required=False),
    FieldDef(
        "indice_ccn", "Convention collective appliquée", required=False,
        help=(
            "Pré-remplie via Pappers. Si Pappers ne l'a pas confirmée "
            "(déduite du seul code NAF), la mention \"(supposée)\" est "
            "ajoutée : à vérifier avant envoi."
        ),
    ),
]

# expert_comptable a une valeur par defaut dans ce modele (signature deja
# associee a Serge AYACHE dans le document d'origine).
SIGNATURE_FIELDS_SOCIALE = [
    FieldDef("lieu_signature", "Lieu de signature", default="Paris"),
    FieldDef("aujourdhui", "Date de la lettre", type="date", default=date.today()),
    FieldDef("expert_comptable", "Expert-comptable signataire", default="Serge AYACHE"),
]


@dataclass
class Mission:
    key: str
    label: str
    template_file: str
    fields: list

    @property
    def template_path(self) -> Path:
        return TEMPLATES_DIR / self.template_file


MISSIONS = {
    "tenue": Mission(
        key="tenue",
        label="Tenue",
        template_file="tenue.docx",
        fields=(
            CLIENT_FIELDS
            + [RAISON_SOCIALE_FIELD]
            + REPRESENTANT_FIELDS
            + [CLOTURE_FIELD]
            + PERIODE_FIELDS
            + [HONORAIRES_CTA_FIELD, HONORAIRES_JUR_FIELD]
            + SIGNATURE_FIELDS
        ),
    ),
    "revision": Mission(
        key="revision",
        label="Révision",
        template_file="revision.docx",
        fields=(
            CLIENT_FIELDS
            + REPRESENTANT_FIELDS
            + [CLOTURE_FIELD]
            + PERIODE_FIELDS
            + [HONORAIRES_CTA_FIELD, HONORAIRES_JUR_FIELD]
            + [f for f in SIGNATURE_FIELDS if f.name != "aujourdhui"]
        ),
    ),
    "bnc": Mission(
        key="bnc",
        label="BNC",
        template_file="bnc.docx",
        fields=(
            CLIENT_FIELDS
            + [RAISON_SOCIALE_FIELD]
            + REPRESENTANT_FIELDS
            + [CLOTURE_FIELD]
            + [HONORAIRES_CTA_FIELD]
            + SIGNATURE_FIELDS
        ),
    ),
    "procedures_convenues": Mission(
        key="procedures_convenues",
        label="Procédures convenues",
        template_file="procedures_convenues.docx",
        fields=(
            CLIENT_FIELDS
            + REPRESENTANT_FIELDS
            + [CLOTURE_FIELD]
            + SIGNATURE_FIELDS_SANS_EC
        ),
    ),
    "procedure_collective": Mission(
        key="procedure_collective",
        label="Procédure collective",
        template_file="procedure_collective.docx",
        fields=(
            CLIENT_FIELDS
            + REPRESENTANT_FIELDS
            + SIGNATURE_FIELDS_SANS_EC
        ),
    ),
    "sociale": Mission(
        key="sociale",
        label="Sociale",
        template_file="sociale.docx",
        fields=(
            CLIENT_FIELDS
            + REPRESENTANT_FIELDS
            + SOCIALE_FIELDS
            + SIGNATURE_FIELDS_SOCIALE
        ),
    ),
}
