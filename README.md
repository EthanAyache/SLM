# Générateur de lettres de mission

Application locale (Streamlit) pour générer automatiquement des lettres de
mission au format Word à partir d'un formulaire de saisie.

## Installation

```
pip install -r requirements.txt
```

Copier `.env.example` en `.env` et y renseigner votre clé API Pappers
(`PAPPERS_API_KEY=...`, disponible sur pappers.fr/api) si vous voulez
utiliser le pré-remplissage automatique par SIREN. Le fichier `.env` n'est
jamais versionné (voir `.gitignore`).

## Lancement

```
streamlit run app.py
```

(Python 3.12 vient d'être installé sur ce poste. Si la commande `streamlit`
n'est pas reconnue, ouvrez un **nouveau** terminal — le PATH doit être
rafraîchi — ou utilisez `python -m streamlit run app.py`.)

Cela ouvre une page dans le navigateur avec le formulaire. Sélectionnez le
type de mission, remplissez les champs, puis cliquez sur "Générer la
lettre" pour télécharger le fichier Word correspondant.

## Pré-remplissage automatique par SIREN (Pappers)

Saisissez un numéro de SIREN (9 chiffres) puis cliquez sur "Rechercher via
Pappers" : le nom du dossier, la forme juridique et l'adresse du siège se
pré-remplissent automatiquement. Si plusieurs dirigeants sont trouvés
(Président, Directeur général, Gérant...), un menu déroulant permet de
choisir celui qui doit apparaître dans la lettre ; sinon il est sélectionné
automatiquement. Tous les champs pré-remplis restent modifiables avant
génération.

Pour la mission **Sociale**, le libellé du code NAF/APE et la convention
collective sont également récupérés. La convention collective n'est pas
toujours déclarée officiellement par l'entreprise (DSN) : quand Pappers ne
l'a pas confirmée, elle est déduite du seul code NAF et la mention
« (supposée) » est ajoutée automatiquement — à vérifier avant l'envoi de
la lettre.

Chaque recherche consomme un crédit de votre abonnement Pappers (offre
gratuite limitée par mois). La date de clôture d'exercice n'est pas
déduite automatiquement (Pappers ne renvoie pas l'année) : à saisir à la
main.

## Types de mission disponibles (V1)

- Tenue
- Révision
- BNC
- Procédures convenues
- Procédure collective

Ce sont les 5 modèles du dossier `LmMatrice/` qui contenaient déjà de vrais
champs de fusion Word (publipostage). Les 12 autres modèles (ECF,
Évaluation société, IR-IFI, Sociale, Transmission d'entreprise, mandats,
etc.) n'utilisent pas une convention homogène de champs et seront ajoutés
dans une prochaine version.

## Structure du projet

```
app.py                          Application Streamlit (formulaire + génération)
generator/
  missions.py                   Configuration des types de mission et de leurs champs
  docx_engine.py                Rendu du .docx à partir du contexte saisi
templates/                      Modèles préparés (tags {{ champ }}), générés par le script ci-dessous
scripts/
  prepare_templates.py          Script one-shot : convertit les .dotx de LmMatrice/ en templates/*.docx
LmMatrice/                      Modèles Word originaux du cabinet (non modifiés)
```

## Régénérer les templates

Si les modèles source dans `LmMatrice/` sont modifiés (nouveau champ de
fusion, texte différent), relancer :

```
python scripts/prepare_templates.py
```

Cela régénère les fichiers dans `templates/` à partir des `.dotx` sources.

## Limites connues (V1)

- Dans les lettres Tenue/Révision, les cases à cocher (fréquence des
  situations intermédiaires, etc.) ne sont pas automatisées : elles restent
  à cocher manuellement dans le Word généré.
- Pas de base clients / historique des lettres déjà générées.
