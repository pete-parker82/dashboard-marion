#!/usr/bin/env python3
"""
refresh.py — Dashboard-Marion

Ordre d'exécution (comme demandé) :
  1. Lit le Google Sheet rempli à la main par Marion (onglets ML-2025, ML-2026, Indemnités km ...)
  2. Scanne le dossier Google Drive et repère tout fichier qui ressemble à de la donnée bancaire
     (export Pennylane, relevé de compte pro), même si le nom de fichier ne le dit pas clairement
  3. Dédoublonne les transactions bancaires (Marion copie parfois toute la page Pennylane, donc
     des lignes identiques peuvent apparaître dans plusieurs exports)
  4. Recalcule les agrégats mensuels (recettes, charges, CARMF, URSSAF, rétrocession...)
  5. Réinjecte ces données dans dashboard-marion.html, entre les marqueurs /*DATA_START*/ /*DATA_END*/

Statut : la mécanique (auth, lecture Sheet, scan+dédoublonnage Drive, écriture du HTML) est
complète et fonctionnelle. build_monthly_dataset() utilise directement les colonnes Type/
Dénomination du Sheet (vérifiées sur les onglets ML-2025 et ML-2026 réels), donc les recettes
et charges de chaque mois sont calculées à partir de ce que Marion saisit, pas d'une estimation.
Seul le "versé compte commun" reste détecté depuis le relevé bancaire (le Sheet ne le suit pas).

Limite connue : si un mois n'a pas encore ses lignes "Recettes Carcans" / "Recettes Ste Hélène"
saisies par Marion, les recettes calculées pour ce mois resteront à 0 — le script ne remplace
jamais un chiffre manquant par une estimation inventée. Le mois reste marqué reel=False dans ce
cas, pour qu'on garde l'ancienne valeur estimée côté dashboard plutôt que d'écraser avec du 0.
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
STATE_PATH = HERE / "monthly_data.json"          # état persistant, gardé entre deux refresh

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
    """Ne lit en détail que les onglets ML-* (les seuls dont build_monthly_dataset se sert) —
    ça évite de planter sur un onglet comme 'Indemnités km' dont les en-têtes de colonnes
    ne sont pas uniques et que get_all_records() n'aime pas."""
    sh = gc.open_by_key(sheet_id)
    data = {}
    for ws in sh.worksheets():
        if not ws.title.startswith("ML-"):
            print(f"  [ignoré] onglet '{ws.title}' (non utilisé par le calcul)")
            continue
        try:
            data[ws.title] = ws.get_all_records()
        except Exception as e:
            print(f"  [attention] onglet '{ws.title}' illisible ({e}) — ignoré, pas de crash")
    return data


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


FR_MONTHS = {
    "janv": 1, "févr": 2, "fevr": 2, "mars": 3, "avr": 4, "mai": 5, "juin": 6,
    "juil": 7, "août": 8, "aout": 8, "sept": 9, "oct": 10, "nov": 11, "déc": 12, "dec": 12,
}
FR_MONTH_LABELS = {
    1: "Janv.", 2: "Fév.", 3: "Mars", 4: "Avr.", 5: "Mai", 6: "Juin",
    7: "Juil.", 8: "Août", 9: "Sept.", 10: "Oct.", 11: "Nov.", 12: "Déc.",
}


def normalize_key(k) -> str:
    """Les en-têtes du Sheet contiennent parfois des retours à la ligne (ex: 'Dénomination\\n(libre)') :
    on les aplatit pour matcher de façon fiable quel que soit le rendu exact du header."""
    return re.sub(r"\s+", " ", str(k)).strip()


def parse_amount_fr(v) -> float:
    if v is None or v == "":
        return 0.0
    s = str(v)
    s = re.sub(r"[\s\u202f\xa0]", "", s)  # espaces normales, insécables, insécables étroites
    s = s.replace("€", "").replace(",", ".").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_date_fr(s) -> str | None:
    """'31-janv.-2025' / '5-août-2026' -> '2025-01' / '2026-08'."""
    s = str(s).strip().lower()
    m = re.match(r"(\d{1,2})-([a-zéû]+)\.?-(\d{4})", s)
    if not m:
        return None
    _, mon_str, year = m.groups()
    mon_str = mon_str.rstrip(".")
    mon = FR_MONTHS.get(mon_str) or FR_MONTHS.get(mon_str[:4])
    if not mon:
        return None
    return f"{year}-{mon:02d}"


# ---------- 3. Agrégation mensuelle ----------

def build_monthly_dataset(sheet_data: dict, bank_df: pd.DataFrame) -> list[dict]:
    """
    Source principale : les onglets 'ML-YYYY' du Sheet, qui portent déjà la catégorisation faite
    par Marion via les colonnes Type / Dénomination (pas de déduction hasardeuse ici — on lit
    directement les valeurs de Type telles que Marion les choisit : Recettes, Aides & Allocation
    Recettes, Cotisations sociales, Honoraires versés, Frais généraux, Achats, Frais quotidiens).

    Le relevé bancaire (bank_df) sert uniquement à détecter le virement vers le compte commun,
    que le Sheet ne suit pas.
    """
    monthly: dict[str, dict] = {}

    def month_bucket(month: str) -> dict:
        return monthly.setdefault(month, {
            "mois": month, "recettes": 0.0, "charges": 0.0,
            "carmf": 0.0, "urssaf": 0.0, "retro": 0.0,
            "loyer": 0.0, "fraisFixes": 0.0, "achats": 0.0,
            "verse": None, "reel": True, "_has_recette_row": False,
        })

    for tab_name, records in sheet_data.items():
        if not tab_name.startswith("ML-"):
            continue
        for raw_row in records:
            row = {normalize_key(k): v for k, v in raw_row.items()}
            month = parse_date_fr(row.get("Date (jj-mm-aaaa)"))
            if not month:
                continue
            denom = str(row.get("Dénomination (libre)", "")).upper()
            type_ = str(row.get("Type", "")).strip()
            montant = parse_amount_fr(row.get("Montant Final (ne pas toucher)") or row.get("Montant (€)"))
            m = month_bucket(month)

            if type_ in ("Recettes", "Aides & Allocation Recettes"):
                m["recettes"] += montant
                if type_ == "Recettes" and montant != 0:
                    m["_has_recette_row"] = True
            elif type_ == "Cotisations sociales":
                if "CARMF" in denom:
                    m["carmf"] += montant
                elif "URSSAF" in denom:
                    m["urssaf"] += montant
                m["charges"] += montant
            elif type_ == "Honoraires versés":
                m["retro"] += montant
                m["charges"] += montant
            elif type_ == "Frais généraux":
                if "LOYER" in denom:
                    m["loyer"] += montant
                else:
                    m["fraisFixes"] += montant
                m["charges"] += montant
            elif type_ in ("Achats", "Frais quotidiens"):
                m["achats"] += montant
                m["charges"] += montant

    # Le virement vers le compte commun n'est pas dans le Sheet : on le lit sur le relevé bancaire.
    if not bank_df.empty:
        bank_df["categorie"] = bank_df["libelle"].apply(categorize)
        bank_df["date"] = pd.to_datetime(bank_df["date"], dayfirst=True, errors="coerce")
        bank_df["mois"] = bank_df["date"].dt.to_period("M").astype(str)
        versements = bank_df[bank_df["categorie"] == "virement_compte_commun"]
        for month, sub in versements.groupby("mois"):
            if month in monthly:
                monthly[month]["verse"] = round(-sub["montant"].sum(), 2)

    result = []
    for month in sorted(monthly.keys()):
        d = monthly[month]
        # Un mois dont les recettes n'ont pas encore été saisies par Marion (pas de ligne
        # "Recettes Carcans"/"Recettes Ste Hélène" avec un montant) reste marqué non-réel :
        # mieux vaut garder l'ancienne estimation côté dashboard que d'écraser avec un 0 trompeur.
        d["reel"] = d.pop("_has_recette_row")
        year, mon = month.split("-")
        d["label"] = f"{FR_MONTH_LABELS[int(mon)]} {year[2:]}"
        for k in ("recettes", "charges", "carmf", "urssaf", "retro", "loyer", "fraisFixes", "achats"):
            d[k] = round(d[k], 2)
        result.append(d)
    return result


# ---------- 3bis. État persistant : ne jamais écraser un mois réel par une donnée incomplète ----------

def load_state() -> dict:
    """dict {mois: {...}} — ce qui a été calculé lors des refresh précédents."""
    if STATE_PATH.exists():
        rows = json.loads(STATE_PATH.read_text(encoding="utf-8"))
        return {r["mois"]: r for r in rows}
    return {}


def merge_with_state(new_rows: list[dict], old_state: dict) -> list[dict]:
    """
    Règle : un mois fraîchement calculé comme reel=True remplace toujours l'ancien.
    Un mois calculé comme reel=False (recettes pas encore saisies par Marion) ne doit
    JAMAIS écraser une version antérieure qui existait déjà (surtout pas avec des 0) —
    dans ce cas on garde l'ancienne version telle quelle, réelle ou estimée.
    """
    merged = dict(old_state)
    for row in new_rows:
        month = row["mois"]
        if row["reel"] or month not in merged:
            merged[month] = row
        # sinon : on garde merged[month] tel qu'il était déjà (ne rien faire)
    return [merged[m] for m in sorted(merged.keys())]


def save_state(rows: list[dict]):
    STATE_PATH.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")


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

    print("Étape 3bis/4 — fusion avec l'état précédent (ne jamais écraser un mois réel par du vide)...")
    old_state = load_state()
    merged = merge_with_state(monthly, old_state)
    save_state(merged)

    print("Étape 4/4 — régénération du dashboard...")
    inject_into_template(merged)

    print(f"[{datetime.now():%Y-%m-%d %H:%M}] Terminé → {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
