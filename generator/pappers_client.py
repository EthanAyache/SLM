"""Client minimal pour l'API entreprise de Pappers (https://api.pappers.fr).

Permet de retrouver, a partir d'un numero de SIREN, les informations a
prereplir dans le formulaire (nom, forme juridique, adresse du siege,
dirigeants).
"""

import os
import re

import requests

from generator.missions import FORME_JURIDIQUE_OPTIONS

API_URL = "https://api.pappers.fr/v2/entreprise"
TIMEOUT = 10

# Qualites de representants consideres comme "dirigeant" pour la lettre de
# mission (on exclut les commissaires aux comptes, censeurs, personnes
# morales, etc.)
QUALITES_DIRIGEANT_RE = re.compile(
    r"président|gérant|directeur général|directrice générale|"
    r"administrateur|liquidateur",
    re.IGNORECASE,
)


class PappersError(Exception):
    """Erreur de base pour les appels a l'API Pappers."""


class SirenIntrouvable(PappersError):
    pass


class CleApiInvalide(PappersError):
    pass


class QuotaDepasse(PappersError):
    pass


class ErreurReseau(PappersError):
    pass


def _get_api_key() -> str:
    key = os.environ.get("PAPPERS_API_KEY", "").strip()
    if not key:
        raise CleApiInvalide(
            "Aucune clé API Pappers configurée (variable PAPPERS_API_KEY "
            "absente du fichier .env)."
        )
    return key


def fetch_entreprise(siren: str) -> dict:
    """Interroge l'API Pappers pour un SIREN donne et retourne le JSON brut."""
    siren = re.sub(r"\D", "", siren or "")
    if len(siren) != 9:
        raise PappersError("Le SIREN doit comporter 9 chiffres.")

    api_key = _get_api_key()

    try:
        response = requests.get(
            API_URL,
            params={"siren": siren},
            headers={"api-key": api_key},
            timeout=TIMEOUT,
        )
    except requests.RequestException as exc:
        raise ErreurReseau(f"Impossible de contacter l'API Pappers : {exc}") from exc

    if response.status_code == 200:
        return response.json()
    if response.status_code in (401, 403):
        raise CleApiInvalide("Clé API Pappers invalide ou refusée.")
    if response.status_code == 404:
        raise SirenIntrouvable(f"Aucune entreprise trouvée pour le SIREN {siren}.")
    if response.status_code in (402, 429):
        raise QuotaDepasse("Quota de requêtes Pappers atteint pour ce mois.")
    raise PappersError(f"Erreur Pappers inattendue (HTTP {response.status_code}).")


def _normaliser_forme_juridique(raw: str) -> str:
    if not raw:
        return "Autre"
    sigle = raw.split(",")[0].strip().upper()
    for option in FORME_JURIDIQUE_OPTIONS:
        if option.upper() == sigle:
            return option
    return "Autre"


def _civilite_depuis_sexe(sexe: str) -> str:
    return "Madame" if sexe == "F" else "Monsieur"


def _representants_candidats(raw: dict) -> list:
    candidats = []
    for rep in raw.get("representants") or []:
        if rep.get("personne_morale"):
            continue
        qualite = rep.get("qualite") or ""
        if not QUALITES_DIRIGEANT_RE.search(qualite):
            continue
        candidats.append(
            {
                "label": f"{rep.get('nom_complet', '')} — {qualite}",
                "civilite_representant": _civilite_depuis_sexe(rep.get("sexe")),
                "nom_representant": rep.get("nom_complet", ""),
                "fonction_representant": qualite,
            }
        )
    return candidats


def _libelle_convention_collective(conventions: list) -> str:
    """Formate le nom de la premiere convention collective renvoyee par
    Pappers. Si elle n'est pas "confirmee" (i.e. deduite du code NAF plutot
    que declaree officiellement par l'entreprise en DSN), on l'indique par
    "(supposée)" pour que l'expert-comptable sache qu'elle est a verifier."""
    if not conventions:
        return ""
    convention = conventions[0]
    nom = convention.get("nom", "") or ""
    if nom and not convention.get("confirmee"):
        nom += " (supposée)"
    return nom


def extract_form_data(raw: dict) -> dict:
    """Transforme la reponse brute de l'API en valeurs prêtes a injecter
    dans le formulaire (clefs = noms des champs du formulaire)."""
    siege = raw.get("siege") or {}
    denomination = raw.get("denomination") or raw.get("nom_entreprise") or ""
    conventions = raw.get("conventions_collectives") or []

    return {
        "nom_du_dossier": denomination,
        "raison_sociale": denomination,
        "forme_juridique": _normaliser_forme_juridique(raw.get("forme_juridique", "")),
        "adresse1": siege.get("adresse_ligne_1", "") or "",
        "adresse2": siege.get("adresse_ligne_2", "") or "",
        "code_postal": siege.get("code_postal", "") or "",
        "ville": siege.get("ville", "") or "",
        "siren": raw.get("siren", "") or "",
        "naf_lib": raw.get("libelle_code_naf", "") or "",
        "indice_ccn": _libelle_convention_collective(conventions),
        "representants_candidats": _representants_candidats(raw),
    }
