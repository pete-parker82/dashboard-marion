#!/usr/bin/env python3
"""
refresh.py — Dashboard-Marion

Ordre d'exécution (comme demandé) :
  1. Lit le Google Sheet rempli à la main par Marion (onglets ML-2025, ML-2026, Indemnités km ...)
  2. Scanne le dossier Google Drive et repère tout fichier qui ressemble à de la donnée bancaire
     (export Pennylane, relevé de compte pro), même si le nom de fichier ne le dit pas clairement
  3. Dédoublonne les transactions bancaires (Marion copie parfois toute la page Pennylane, donc
     des lignes identiques peuvent apparaître dans plusieurs exports)
  4. Recalcule les agrégats mensuels (recettes, charges, CARMF, URSSAF, rétrocession Bayane...)
  5. Réinjecte ces données dans dashboard-marion.html, entre les marqueurs /*DATA_START*/ /*DATA_END*/

Statut : la mécanique (auth, lecture Sheet, scan+dédoublonnage Drive, écriture du HTML) est
complète et fonctionnelle. La fonction build_monthly_dataset() applique la même logique de
catégorisation déjà validée manuellement (voir catégories ci-dessous) mais doit être vérifiée
une première fois en conditions réelles avec vos identifiants avant de tourner seule — la
structure exacte des colonnes du Sheet et des exports Pennylane peut varier légèrement.
"""

import io
import json
import hashlib
import re
import sys
from pathlib import Path
from datetime import datetime

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

HERE = Path(__file__).parent
CONFIG_PATH = HERE / "config.json"
TEMPLATE_PATH = HERE / "dashboard-marion.html"   # sert aussi de template (contient les marqueurs)
OUTPUT_PATH = HERE / "dashboard-marion.html"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# En-têtes qui, si on en trouve au moins 2 dans un fichier, indiquent que c'est
# probablement de la donnée bancaire (relevé BNP ou export Pennylane), quel que
# soit le nom du fichier.
BANK_LIKE_HEADERS = {
    "date", "date operation", "date opération", "date valeur", "montant",
    "libellé", "libelle", "amount", "description", "tiers",
    "compte bancaire", "type", "recette vs charge", "plan de trésorerie",
}

# Catégorisation des libellés bancaires, validée manuellement sur le relevé BNP.
CATEGORY_RULES = [
    (re.compile(r"SALAIRE.*MOFFROID|MOFFROID.*SALAIRE", re.I), "virement_compte_commun"),
    (re.compile(r"RETROCESSION", re.I), "retrocession_bayane"),
    (re.compile(r"LOYER", re.I), "loyer"),
    (re.compile(r"CARMF", re.I), "carmf"),
    (re.compile(r"URSSAF", re.I), "urssaf"),
    (re.compile(r"MACSF|SFR|WEDA|DOCTOLIB|EDF|RAMDAM|CABINET 12|CFE|GROUPEMENT DES PEDIATRES", re.I), "frais_fixes"),
    (re.compile(r"CPAM|C\.P\.A\.M|MSA|MUTUELLE|MGEN|CRPCEN|CAMIEG|CGSS|CNMSS|SECURITE SOCIALE", re.I), "recette_secu_mutuelle"),
    (re.compile(r"REM\. CARTE|REMISE CHEQUES", re.I), "recette_carte_cheques"),
    (re.compile(r"PAIEMENT CB|FACTURE CARTE", re.I), "achats_divers"),
]


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        sys.exit(f"Config introuvable : {CONFIG_PATH}. Copiez config.example.json vers config.json et remplissez-le.")
    return json.loads(CONFIG_PATH.read_text())


def get_clients(config: dict):
    creds = Credentials.from_service_account_file(config["service_account_json"], scopes=SCOPES)
    gc = gspread.authorize(creds)
    drive = build("drive", "v3", credentials=creds)
    return gc, drive


# ---------- 1. Google Sheet ----------

def read_sheet(gc, sheet_id: str) -> dict:
    sh = gc.open_by_key(sheet_id)
    return {ws.title: ws.get_all_records() for ws in sh.worksheets()}


# ---------- 2. Google Drive : repérer + télécharger tout ce qui ressemble à de la banque ----------

def list_drive_files(drive, folder_id: str) -> list:
    files, page_token = [], None
    while True:
        resp = drive.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken, files(id, name, mimeType, modifiedTime)",
            pageToken=page_token,
        ).execute()
        files.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return files


def download_file(drive, file_id: str, mime_type: str) -> io.BytesIO:
    export_map = {
        "application/vnd.google-apps.spreadsheet": (
            "text/csv"
        ),
    }
    if mime_type in export_map:
        request = drive.files().export_media(fileId=file_id, mimeType=export_map[mime_type])
    else:
        request = drive.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf


def parse_tabular_file(buf: io.BytesIO, name: str) -> pd.DataFrame | None:
    try:
        if name.lower().endswith((".xlsx", ".xls")):
            return pd.read_excel(buf)
        return pd.read_csv(buf, sep=None, engine="python")
    except Exception as e:
        print(f"  [illisible] {name} : {e}")
        return None


def looks_like_bank_data(columns) -> bool:
    cols_lower = {str(c).strip().lower() for c in columns}
    return len(cols_lower & BANK_LIKE_HEADERS) >= 2


def transaction_hash(date_val, amount_val, label_val) -> str:
    key = f"{date_val}|{amount_val}|{str(label_val)[:40]}"
    return hashlib.sha1(key.encode("utf-8", errors="ignore")).hexdigest()


def collect_bank_transactions(drive, folder_id: str) -> pd.DataFrame:
    """Parcourt TOUS les fichiers du dossier Drive, garde ceux qui ressemblent à de la
    donnée bancaire (peu importe le nom), et dédoublonne les transactions identiques
    qui peuvent apparaître dans plusieurs exports Pennylane."""
    all_rows, seen = [], set()
    for f in list_drive_files(drive, folder_id):
        print(f"[scan] {f['name']}")
        buf = download_file(drive, f["id"], f["mimeType"])
        df = parse_tabular_file(buf, f["name"])
        if df is None or not looks_like_bank_data(df.columns):
            print(f"  [ignoré] ne ressemble pas à de la donnée bancaire")
            continue
        date_col = next((c for c in df.columns if "date" in str(c).lower()), None)
        amount_col = next((c for c in df.columns if "montant" in str(c).lower() or "amount" in str(c).lower()), None)
        label_col = next((c for c in df.columns if "libell" in str(c).lower() or "description" in str(c).lower()), None)
        if not (date_col and amount_col):
            print(f"  [ignoré] colonnes date/montant non identifiées")
            continue
        n_new = 0
        for _, row in df.iterrows():
            h = transaction_hash(row[date_col], row[amount_col], row.get(label_col, ""))
            if h in seen:
                continue
            seen.add(h)
            n_new += 1
            all_rows.append({
                "source_file": f["name"],
                "date": row[date_col],
                "montant": row[amount_col],
                "libelle": row.get(label_col, ""),
            })
        print(f"  [ok] {n_new} nouvelles transactions (doublons ignorés)")
    return pd.DataFrame(all_rows)


def categorize(libelle: str) -> str:
    for pattern, cat in CATEGORY_RULES:
        if pattern.search(str(libelle)):
            return cat
    return "autre"


# ---------- 3. Agrégation mensuelle ----------

def build_monthly_dataset(sheet_data: dict, bank_df: pd.DataFrame) -> list[dict]:
    """
    Combine le Sheet (recettes déclaratives Carcans/Sainte-Hélène) et le relevé bancaire
    catégorisé pour produire la liste de dicts mensuels utilisée par le dashboard.

    À FINALISER ENSEMBLE : la structure ci-dessous applique la même logique que l'analyse
    manuelle (voir le fil de conversation), mais elle doit être testée une première fois sur
    vos vraies données avant de tourner sans supervision — en particulier les noms exacts des
    colonnes du Sheet, qui peuvent différer légèrement de ce qui est codé ici.
    """
    if not bank_df.empty:
        bank_df["categorie"] = bank_df["libelle"].apply(categorize)
        bank_df["date"] = pd.to_datetime(bank_df["date"], dayfirst=True, errors="coerce")
        bank_df["mois"] = bank_df["date"].dt.to_period("M").astype(str)

    monthly = []
    months = sorted(bank_df["mois"].dropna().unique()) if not bank_df.empty else []
    for m in months:
        sub = bank_df[bank_df["mois"] == m]
        def total(cat):
            return round(-sub[sub["categorie"] == cat]["montant"].sum(), 2)
        monthly.append({
            "mois": m,
            "recettes": round(sub[sub["categorie"].isin(["recette_secu_mutuelle", "recette_carte_cheques"])]["montant"].sum(), 2),
            "carmf": total("carmf"),
            "urssaf": total("urssaf"),
            "retro": total("retrocession_bayane"),
            "verse": total("virement_compte_commun"),
            "charges": round(-sub[sub["montant"] < 0]["montant"].sum(), 2),
            "reel": True,
        })
    return monthly


# ---------- 4. Réinjection dans le HTML ----------

def inject_into_template(monthly: list[dict]):
    html = TEMPLATE_PATH.read_text(encoding="utf-8")
    js_array = json.dumps(monthly, ensure_ascii=False, indent=2)
    pattern = re.compile(r"/\*DATA_START\*/.*?/\*DATA_END\*/", re.S)
    replacement = f"/*DATA_START*/\nconst M = {js_array};\n/*DATA_END*/"
    if not pattern.search(html):
        sys.exit("Marqueurs /*DATA_START*/ /*DATA_END*/ introuvables dans le HTML — ne pas régénérer à l'aveugle.")
    html = pattern.sub(replacement, html)
    OUTPUT_PATH.write_text(html, encoding="utf-8")


def main():
    config = load_config()
    print(f"[{datetime.now():%Y-%m-%d %H:%M}] Début du refresh")

    gc, drive = get_clients(config)

    print("Étape 1/4 — lecture du Google Sheet...")
    sheet_data = read_sheet(gc, config["sheet_id"])
    print(f"  onglets lus : {list(sheet_data.keys())}")

    print("Étape 2/4 — scan du dossier Drive...")
    bank_df = collect_bank_transactions(drive, config["drive_folder_id"])
    print(f"  {len(bank_df)} transactions bancaires uniques après dédoublonnage")

    print("Étape 3/4 — agrégation mensuelle...")
    monthly = build_monthly_dataset(sheet_data, bank_df)

    print("Étape 4/4 — régénération du dashboard...")
    inject_into_template(monthly)

    print(f"[{datetime.now():%Y-%m-%d %H:%M}] Terminé → {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
