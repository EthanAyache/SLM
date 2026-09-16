"""
Script one-shot : convertit les modeles .dotx source de LmMatrice/ (champs de
fusion Word MERGEFIELD + libelles de periode entre crochets) en fichiers .docx
"propres" utilisant des tags Jinja {{ champ }}, exploitables par docxtpl.

Usage : python scripts/prepare_templates.py
"""

import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
SOURCE_DIR = ROOT / "LmMatrice"
OUTPUT_DIR = ROOT / "templates"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


# Normalisation des noms de champs MERGEFIELD (tels que trouves dans les .dotx)
# vers des noms de variables Python (snake_case) utilises dans le formulaire.
FIELD_NAME_MAP = {
    "nom_du_dossier": "nom_du_dossier",
    "raison_sociale": "raison_sociale",
    "forme_juridique": "forme_juridique",
    "adresse1": "adresse1",
    "adresse2": "adresse2",
    "code_postal": "code_postal",
    "ville": "ville",
    "civilite_representant": "civilite_representant",
    "nom_representant": "nom_representant",
    "fonction_representant": "fonction_representant",
    "date_de_cloture": "date_de_cloture",
    "Lieu_signature": "lieu_signature",
    "expert_comptable": "expert_comptable",
    "aujourdhui": "aujourdhui",
    "Hono_CTA": "hono_cta",
    "HONO_JUR": "hono_jur",
}

# Libelles bruts entre crochets (texte litteral, pas des champs de fusion)
# a remplacer par des tags Jinja.
BRACKET_TEXT_MAP = {
    "[Janvier à Avril]": "{{ periode_revision }}",
    "[Avril-Mai]": "{{ periode_liasse }}",
    "[Avril-Mai-Juin]": "{{ periode_juridique }}",
    "[ou intermédiaires]": "{{ mention_intermediaire }}",
    # Restes d'un ancien systeme de publipostage (pages "Annexe" des
    # en-tetes), reperes dans MODELE - LM TENUE/PROCEDURES CONVENUES/
    # Procédure collective. nom_du_dossier plutôt que raison_sociale : ce
    # dernier n'existe pas dans tous les modeles.
    "<dossier.RaisonSociale>": "{{ nom_du_dossier }}",
    # Restes de saisie manuelle dans l'en-tete "Exercice clos le..." de
    # MODELE - LM BNC.
    "[Nom société]": "{{ nom_du_dossier }}",
    "[31 décembre 2020]": "{{ date_de_cloture }}",
    # MODELE - LM SOCIALE : date/lieu et signataire tapes en texte libre
    # (pas de champ de fusion pour ceux-la).
    "Paris/Orléans/Versailles, le 00/00/00": "{{ lieu_signature }}, le {{ aujourdhui }}",
    "Fait à Paris/Versailles/Orléans,": "Fait à {{ lieu_signature }},",
    "Serge AYACHE": "{{ expert_comptable }}",
}

# MODELE - LM SOCIALE utilise une troisieme convention de placeholder,
# distincte des MERGEFIELD et des chevrons : des variables "$PREFIXE_NOM"
# issues d'un autre systeme de publipostage. Les combinaisons de plusieurs
# variables sur une meme ligne doivent etre listees AVANT les variables
# seules, sinon le remplacement partiel casserait la combinaison.
DOLLAR_VAR_MAP = {
    "$STE_RAISONSOCIALE $STE_FORMEJURIDIQUE": "{{ nom_du_dossier }} {{ forme_juridique }}",
    "$CLI_EMPLOYEURCIVILITE2 $CLI_EMPLOYEURPRENOM $CLI_EMPLOYEURNOM": "{{ civilite_representant }} {{ nom_representant }}",
    "$CLI_EMPLOYEURCIVILITE2 $CLI_EMPLOYEURNOM": "{{ civilite_representant }} {{ nom_representant }}",
    "$STE_NUMVOIE $STE_NOMVOIE $STE_COMPLEMENTADRESSE": "{{ adresse1 }} {{ adresse2 }}",
    "$STE_CODEPOSTAL $STE_NOMVILLE": "{{ code_postal }} {{ ville }}",
    # Le prenom seul (colonne signature, sous le nom) devient superflu
    # puisque nom_representant contient deja le nom complet ci-dessus.
    "$CLI_EMPLOYEURPRENOM": "",
    "$CLI_EMPLOYEURQUALITE": "{{ fonction_representant }}",
    "$STE_SIREN": "{{ siren }}",
    "$STE_NAFLIB": "{{ naf_lib }}",
    "$STE_INDICECCNX": "{{ indice_ccn }}",
    # Filets de securite si une variable apparaissait seule ailleurs.
    "$STE_RAISONSOCIALE": "{{ nom_du_dossier }}",
    "$STE_FORMEJURIDIQUE": "{{ forme_juridique }}",
}

# Dans certains en-têtes (ex: "«forme_juridique» «nom_du_dossier»" en tête
# de page), le texte «champ» a été tapé littéralement plutôt que via un
# vrai champ de fusion Word (pas de fldChar/instrText associé) : ce ne sont
# donc pas des champs "vivants" detectables par find_all_field_spans, juste
# du texte statique qui ressemble a un champ merge. On les remplace aussi
# par simple substitution de texte, comme pour BRACKET_TEXT_MAP.
CHEVRON_TEXT_MAP = {
    f"«{raw_name}»": "{{ " + var_name + " }}"
    for raw_name, var_name in FIELD_NAME_MAP.items()
}

MISSIONS = {
    "tenue": "MODELE - LM TENUE.dotx",
    "revision": "MODELE - LM REVISION.dotx",
    "bnc": "MODELE - LM BNC.dotx",
    "procedures_convenues": "MODELE - LM PROCEDURES CONVENUES.dotx",
    "procedure_collective": "MODELE - LM Procédure collective.dotx",
    "sociale": "MODELE - LM SOCIALE MODELE 2025.12.docx",
}


def find_all_field_spans(paragraph):
    """
    Parcourt un paragraphe et repere toutes les sequences de champs Word
    (fldChar begin -> instrText (eventuellement fragmente sur plusieurs
    runs) -> fldChar separate -> runs de resultat en cache -> fldChar end).
    Le texte d'instruction (ex: " MERGEFIELD nom_du_dossier ") peut etre
    reparti sur plusieurs runs w:instrText a cause de l'historique
    d'edition Word (rsid) ; on les concatene avant d'extraire le nom du
    champ.
    Retourne une liste de tuples (start_index, end_index, field_name).
    """
    children = list(paragraph)
    n = len(children)
    spans = []
    i = 0
    while i < n:
        el = children[i]
        fld = el.find(w("fldChar")) if el.tag == w("r") else None
        if fld is None or fld.get(w("fldCharType")) != "begin":
            i += 1
            continue
        start = i
        instr_parts = []
        j = i + 1
        end = None
        while j < n:
            cur = children[j]
            if cur.tag != w("r"):
                j += 1
                continue
            cur_fld = cur.find(w("fldChar"))
            if cur_fld is not None and cur_fld.get(w("fldCharType")) == "end":
                end = j
                break
            instr = cur.find(w("instrText"))
            if instr is not None and instr.text:
                instr_parts.append(instr.text)
            j += 1
        if end is None:
            i += 1
            continue
        instruction = "".join(instr_parts)
        m = re.search(r'MERGEFIELD\s+"?([^\s"]+)"?', instruction)
        field_name = m.group(1) if m else None
        if field_name:
            spans.append((start, end, field_name))
        i = end + 1
    return spans


def replace_field_with_text(paragraph, start, end, text):
    """Remplace les runs [start, end] (inclusifs) par un run unique
    contenant `text`, en reprenant le rPr du run affichant la valeur en
    cache (le dernier run avant fldChar end qui contient un w:t)."""
    children = list(paragraph)
    rpr = None
    for idx in range(end - 1, start, -1):
        cand = children[idx]
        if cand.tag == w("r") and cand.find(w("t")) is not None:
            rpr_el = cand.find(w("rPr"))
            if rpr_el is not None:
                rpr = rpr_el
            break
    if rpr is None:
        rpr_el = children[start].find(w("rPr"))
        rpr = rpr_el

    new_run = etree.Element(w("r"))
    if rpr is not None:
        new_run.append(etree.fromstring(etree.tostring(rpr)))
    t_el = etree.SubElement(new_run, w("t"))
    t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_el.text = text

    # Inserer new_run juste avant le premier run du champ, puis
    # supprimer tous les anciens runs [start, end].
    children[start].addprevious(new_run)
    for idx in range(start, end + 1):
        paragraph.remove(children[idx])


def convert_mergefields(root):
    paragraphs = root.findall(".//w:p", NSMAP)
    for p in paragraphs:
        spans = find_all_field_spans(p)
        if not spans:
            continue
        # Traiter en ordre inverse pour ne pas invalider les indices des
        # spans precedents lors des suppressions.
        for start, end, field_name in reversed(spans):
            var_name = FIELD_NAME_MAP.get(field_name)
            if var_name is None:
                continue
            replace_field_with_text(p, start, end, "{{ " + var_name + " }}")


BREAK_TAGS = {w("tab"), w("br"), w("cr")}


def _iter_text_segments(paragraph):
    """Decoupe un paragraphe en segments de w:t consecutifs, coupes a
    chaque tabulation/saut de ligne (w:tab, w:br, w:cr). Necessaire pour
    les paragraphes a colonnes alignees par tabulations (ex: bloc de
    signature "Nom \t Serge AYACHE") : il ne faut jamais fusionner du
    texte de part et d'autre d'une tabulation, sous peine de casser
    l'alignement des colonnes."""
    segment = []
    for el in paragraph.iter():
        if el.tag == w("t"):
            segment.append(el)
        elif el.tag in BREAK_TAGS:
            if segment:
                yield segment
                segment = []
    if segment:
        yield segment


def convert_bracket_text(root):
    """Remplace le texte litteral connu (crochets d'instruction, chevrons de
    champ non lies a un vrai champ de fusion, variables $PREFIXE_NOM) par
    des tags Jinja.

    Fait au niveau du segment de texte (suite de w:t entre deux
    tabulations/sauts de ligne, ou pour tout le paragraphe s'il n'y en a
    pas) et non run par run, car Word scinde parfois un texte visuellement
    continu sur plusieurs runs w:r/w:t (rsid, correction orthographique...)
    : un remplacement run-par-run manquerait alors les occurrences dont le
    motif est a cheval sur deux runs. On concatene donc le texte du
    segment, on remplace, puis on reinjecte le resultat dans le premier
    run textuel du segment (les runs suivants sont vides) : le segment
    garde le formatage de son premier run pour la portion remplacee, ce
    qui est un compromis acceptable pour ce texte de type "placeholder".
    """
    replacements = {**BRACKET_TEXT_MAP, **CHEVRON_TEXT_MAP, **DOLLAR_VAR_MAP}
    for p in root.findall(".//w:p", NSMAP):
        for t_elements in _iter_text_segments(p):
            full_text = "".join(t.text or "" for t in t_elements)
            if not any(literal in full_text for literal in replacements):
                continue
            new_text = full_text
            for literal, tag in replacements.items():
                new_text = new_text.replace(literal, tag)
            # Certains en-têtes contenaient a la fois un texte litteral
            # "«champ»" et, juste a cote, un vrai champ de fusion pour la
            # même donnee (doublon deja present dans le modele Word
            # d'origine) : une fois les deux convertis, on se retrouve
            # avec le même tag repete collé
            # ("{{ nom_du_dossier }}{{ nom_du_dossier }}"). On dedoublonne.
            previous = None
            while previous != new_text:
                previous = new_text
                new_text = re.sub(r"(\{\{ \w+ \}\})\1", r"\1", new_text)
            t_elements[0].text = new_text
            t_elements[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            for t in t_elements[1:]:
                t.text = ""


def convert_template(src_path: Path, dst_path: Path):
    with zipfile.ZipFile(src_path) as zin:
        names = zin.namelist()
        contents = {name: zin.read(name) for name in names}

    # Les champs de fusion peuvent se trouver dans le corps du document
    # mais aussi dans les en-tetes et pieds de page (word/headerN.xml,
    # word/footerN.xml), qui sont des parties XML separees.
    target_parts = [
        name for name in names
        if name == "word/document.xml" or re.match(r"word/(header|footer)\d*\.xml", name)
    ]

    for part_name in target_parts:
        root = etree.fromstring(contents[part_name])
        convert_mergefields(root)
        convert_bracket_text(root)
        contents[part_name] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    # Corriger le content-type : template -> document, pour que le fichier
    # de sortie (.docx) soit reconnu comme un document Word normal.
    ct_xml = contents["[Content_Types].xml"].decode("utf-8")
    ct_xml = ct_xml.replace(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    )
    contents["[Content_Types].xml"] = ct_xml.encode("utf-8")

    dst_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, contents[name])


def main():
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True)

    for key, filename in MISSIONS.items():
        src = SOURCE_DIR / filename
        dst = OUTPUT_DIR / f"{key}.docx"
        print(f"Conversion de {filename} -> templates/{key}.docx")
        convert_template(src, dst)

    print("\nTermine. Verifiez les fichiers dans templates/.")


if __name__ == "__main__":
    main()
