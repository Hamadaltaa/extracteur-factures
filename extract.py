"""Extrait les données de factures PDF (natifs ou scannés) vers un fichier Excel.

Usage : python extract.py [dossier_factures] [fichier_sortie.xlsx]
"""

import json
import sys
import time
from pathlib import Path
from google import genai
from google.genai import types
from openpyxl import Workbook

MODEL = "gemini-3.8-flash"
TOLERANCE = 0.02

PROMPT = """Analyse cette facture. Réponds UNIQUEMENT avec un objet JSON
(sans texte autour) avec ces clés :
fournisseur (texte), date (AAAA-MM-JJ), total_ht (nombre), tva (nombre),
total_ttc (nombre), devise (texte).
Si une valeur est introuvable, mets null. Ne devine pas.
"""

COLONNES = [
    "fichier",
    "fournisseur",
    "date",
    "total_ht",
    "tva",
    "total_ttc",
    "devise",
    "a_verifier",
]


def extraire_pdf_direct(client: genai.Client, chemin_pdf: Path) -> dict:
    fichier_api = client.files.upload(file=str(chemin_pdf))
    try:
        for tentative in range(3):
            try:
                reponse = client.models.generate_content(
                    model=MODEL,
                    contents=[fichier_api, PROMPT],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json"
                    ),
                )
                brut = reponse.text.strip()
                brut = (
                    brut.removeprefix("```json")
                    .removeprefix("```")
                    .removesuffix("```")
                    .strip()
                )
                return json.loads(brut)
            except Exception as e:
                if "503" in str(e) or "UNAVAILABLE" in str(e):
                    print(
                        "  -> Serveur occupé (503), nouvelle tentative dans 5s"
                        f" ({tentative + 1}/3)..."
                    )
                    time.sleep(5)
                else:
                    raise e
        raise RuntimeError(
            "Le serveur Gemini est resté indisponible après 3 tentatives."
        )
    finally:
        client.files.delete(name=fichier_api.name)


def verifier(d: dict) -> str:
    """Retourne la raison du doute, ou une chaîne vide si tout est cohérent."""
    problemes = []
    for cle in ("fournisseur", "date", "total_ht", "tva", "total_ttc"):
        if d.get(cle) is None:
            problemes.append(f"{cle} manquant")
    if not problemes:
        ecart = abs(d["total_ht"] + d["tva"] - d["total_ttc"])
        if ecart > TOLERANCE:
            problemes.append(f"HT + TVA différent du TTC (écart {ecart:.2f})")
    return "; ".join(problemes)


def main() -> None:
    dossier = Path(sys.argv[1] if len(sys.argv) > 1 else "factures")
    sortie = sys.argv[2] if len(sys.argv) > 2 else "resultat.xlsx"
    client = genai.Client()

    classeur = Workbook()
    feuille = classeur.active
    feuille.append(COLONNES)

    pdfs = sorted(dossier.glob("*.pdf"))
    if not pdfs:
        sys.exit(f"Aucun PDF trouvé dans {dossier}/")

    for pdf in pdfs:
        print(f"Traitement : {pdf.name}")
        try:
            donnees = extraire_pdf_direct(client, pdf)
        except Exception as erreur:
            print(f"  -> Erreur sur {pdf.name} : {erreur}")
            feuille.append([pdf.name] + [None] * 6 + [f"erreur : {erreur}"])
            continue

        feuille.append(
            [pdf.name]
            + [donnees.get(c) for c in COLONNES[1:7]]
            + [verifier(donnees)]
        )

    classeur.save(sortie)
    print(f"Terminé : {sortie}")


if __name__ == "__main__":
    main()
