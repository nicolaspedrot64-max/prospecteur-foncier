import gzip
import io
import json
import math
import re
import time
import unicodedata
import zipfile
from datetime import date
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import pydeck as pdk
import requests
try:
    import duckdb
except ImportError:
    duckdb = None
import streamlit as st
from docx import Document
from pypdf import PdfReader
from pyproj import Geod
from shapely.geometry import shape, mapping
from shapely.strtree import STRtree

APP_DIR = Path(__file__).parent
LETTER_TEMPLATE = APP_DIR / "lettre_sagec_modele.docx"

GEO_API = "https://geo.api.gouv.fr"
CADASTRE_BASE = "https://cadastre.data.gouv.fr/data/etalab-cadastre/latest/geojson/communes"
CADASTRE_BASE_FALLBACK = "https://files.data.gouv.fr/cadastre/etalab-cadastre/latest/geojson/communes"
GPU_API = "https://apicarto.ign.fr/api/gpu"
GEOCODAGE_API = "https://data.geopf.fr/geocodage"

# Fichier open data DGFiP des parcelles détenues par des personnes morales.
# La ressource data.gouv.fr redirige vers le parquet courant.
FPMU_PARCELLES_RESOURCE = (
    "https://www.data.gouv.fr/api/1/datasets/r/"
    "913b0f65-3a22-4049-beb0-c33e5084df79"
)

GEOD = Geod(ellps="WGS84")

REGIONS = {
    "Nouvelle-Aquitaine": {
        "16": "Charente",
        "17": "Charente-Maritime",
        "19": "Corrèze",
        "23": "Creuse",
        "24": "Dordogne",
        "33": "Gironde",
        "40": "Landes",
        "47": "Lot-et-Garonne",
        "64": "Pyrénées-Atlantiques",
        "79": "Deux-Sèvres",
        "86": "Vienne",
        "87": "Haute-Vienne",
    },
    "Occitanie": {
        "09": "Ariège",
        "11": "Aude",
        "12": "Aveyron",
        "30": "Gard",
        "31": "Haute-Garonne",
        "32": "Gers",
        "34": "Hérault",
        "46": "Lot",
        "48": "Lozère",
        "65": "Hautes-Pyrénées",
        "66": "Pyrénées-Orientales",
        "81": "Tarn",
        "82": "Tarn-et-Garonne",
    },
}

st.set_page_config(
    page_title="Prospecteur Foncier V3.4",
    page_icon="🏗️",
    layout="wide",
)

# --------------------------
# Helpers réseau
# --------------------------
BDNB_API = "https://api.bdnb.io/v1/bdnb/donnees/batiment_groupe_complet"

HEADERS = {
    "User-Agent": "ProspecteurFoncier/3.4 (Streamlit; donnees publiques)",
    "Accept": "application/json, application/geo+json, */*",
}


def http_get_json(url, params=None, timeout=60):
    r = requests.get(url, params=params, headers=HEADERS, timeout=timeout)
    r.raise_for_status()
    return r.json()


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def fetch_communes(dept_code):
    url = f"{GEO_API}/departements/{dept_code}/communes"
    data = http_get_json(
        url,
        params={
            "fields": "nom,code,codeEpci,codesPostaux,population",
            "format": "json",
        },
        timeout=30,
    )
    return sorted(data, key=lambda x: x.get("nom", ""))


def _dept_from_insee(insee):
    insee = str(insee)
    if insee.startswith("97"):
        return insee[:3]
    return insee[:2]


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def fetch_cadastre_layer(insee, layer):
    dep = _dept_from_insee(insee)
    filename = f"cadastre-{insee}-{layer}.json.gz"
    urls = [
        f"{CADASTRE_BASE}/{dep}/{insee}/{filename}",
        f"{CADASTRE_BASE_FALLBACK}/{dep}/{insee}/{filename}",
    ]
    last_error = None
    for url in urls:
        try:
            r = requests.get(url, headers=HEADERS, timeout=120)
            r.raise_for_status()
            return json.loads(gzip.decompress(r.content).decode("utf-8"))
        except Exception as exc:
            last_error = exc
    raise RuntimeError(
        f"Impossible de télécharger la couche cadastrale '{layer}' pour {insee}: {last_error}"
    )


def _normalise_parcelle_id(value):
    if value is None:
        return ""
    return re.sub(r"[^0-9A-Za-z]", "", str(value)).upper()


def _parse_parcelle_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return [_normalise_parcelle_id(v) for v in value if v]
    if isinstance(value, tuple):
        return [_normalise_parcelle_id(v) for v in value if v]

    txt = str(value).strip()
    if not txt:
        return []

    try:
        parsed = json.loads(txt)
        if isinstance(parsed, list):
            return [_normalise_parcelle_id(v) for v in parsed if v]
    except Exception:
        pass

    txt = txt.strip("{}[]")
    parts = re.split(r"[,;]", txt)
    return [_normalise_parcelle_id(p.strip().strip('"').strip("'")) for p in parts if p.strip()]


@st.cache_data(ttl=7 * 24 * 3600, show_spinner=False)
def fetch_bdnb_commune(insee):
    """
    Charge les usages principaux des bâtiments BDNB de la commune.
    Cette donnée sert uniquement à repérer les parcelles déjà occupées
    par du logement collectif.
    """
    fields = ",".join([
        "batiment_groupe_id",
        "code_commune_insee",
        "l_parcelle_id",
        "usage_principal_bdnb_open",
        "nb_log",
        "libelle_adr_principale_ban",
    ])

    rows = []
    limit = 1000
    offset = 0

    for _ in range(100):
        params = {
            "code_commune_insee": f"eq.{insee}",
            "select": fields,
            "limit": limit,
            "offset": offset,
            "order": "batiment_groupe_id.asc",
        }
        r = requests.get(BDNB_API, params=params, headers=HEADERS, timeout=120)
        r.raise_for_status()
        page = r.json()
        if not isinstance(page, list):
            raise RuntimeError("Réponse BDNB inattendue.")
        rows.extend(page)
        if len(page) < limit:
            break
        offset += limit

    return rows


def build_bdnb_parcel_index(rows):
    index = {}
    for row in rows or []:
        for pid in _parse_parcelle_list(row.get("l_parcelle_id")):
            if pid:
                index.setdefault(pid, []).append(row)
    return index


def detect_collective_housing(parcel_bdnb_rows):
    """
    Retourne True uniquement si la BDNB identifie du résidentiel collectif
    sur la parcelle. Tous les autres cas restent éligibles.
    """
    usages = []
    adresse = ""
    for row in parcel_bdnb_rows or []:
        usage = str(row.get("usage_principal_bdnb_open") or "").strip()
        if usage:
            usages.append(usage)
        if not adresse:
            adresse = str(row.get("libelle_adr_principale_ban") or "").strip()

    usages_norm = [normalize_text(u) for u in usages]
    collectif = any("residentiel collectif" in u for u in usages_norm)

    return {
        "collectif_existant": collectif,
        "usage_bdnb": " / ".join(sorted(set(usages))),
        "adresse_bdnb": adresse,
    }

def _sql_ident(name):
    return '"' + str(name).replace('"', '""') + '"'


def _pick_column(columns, exact=(), contains_all=(), contains_any=()):
    """
    Trouve une colonne dans un schéma qui peut évoluer légèrement entre millésimes.
    """
    lower = {str(c).lower(): str(c) for c in columns}
    for e in exact:
        if e.lower() in lower:
            return lower[e.lower()]

    for c in columns:
        lc = str(c).lower()
        if contains_all and all(x.lower() in lc for x in contains_all):
            return str(c)

    for c in columns:
        lc = str(c).lower()
        if contains_any and any(x.lower() in lc for x in contains_any):
            return str(c)

    return None


def _clean_owner_text(value):
    if value is None:
        return ""
    txt = str(value).strip()
    if txt.lower() in {"", "nan", "none", "null", "<na>"}:
        return ""
    return re.sub(r"\s+", " ", txt)


def _canonical_parcel_id(insee, prefixe, section, numero):
    insee = str(insee or "").strip()
    prefixe = re.sub(r"\D", "", str(prefixe or "")).zfill(3)[-3:]
    section = re.sub(r"[^0-9A-Za-z]", "", str(section or "")).upper().zfill(2)[-2:]
    numero = re.sub(r"\D", "", str(numero or "")).zfill(4)[-4:]
    return _normalise_parcelle_id(f"{insee}{prefixe}{section}{numero}")


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def fetch_personnes_morales_commune(insee):
    """
    Interroge le parquet open data des parcelles appartenant à des personnes morales.
    Le fichier complet fait plusieurs centaines de Mo ; DuckDB ne télécharge que les
    blocs nécessaires grâce aux requêtes HTTP Range.
    """
    if duckdb is None:
        raise RuntimeError(
            "Le module duckdb n'est pas installé. Ajoute duckdb dans requirements.txt."
        )

    con = duckdb.connect(database=":memory:")
    try:
        # Les versions récentes de DuckDB chargent httpfs automatiquement ; on tente
        # explicitement le chargement puis on laisse l'auto-load prendre le relais.
        try:
            con.execute("LOAD httpfs")
        except Exception:
            try:
                con.execute("INSTALL httpfs")
                con.execute("LOAD httpfs")
            except Exception:
                pass

        parquet_sql = F"read_parquet('{FPMU_PARCELLES_RESOURCE}')"
        schema_df = con.execute(f"DESCRIBE SELECT * FROM {parquet_sql}").fetchdf()
        columns = schema_df["column_name"].astype(str).tolist()

        col_insee = _pick_column(
            columns,
            exact=("code_insee", "code_commune_insee", "codcomm"),
            contains_all=("insee",),
        )
        col_parcelle = _pick_column(
            columns,
            exact=("parcelle_id", "id_parcelle", "parcelle", "idpar"),
            contains_all=("parcelle", "id"),
        )
        col_prefixe = _pick_column(
            columns,
            exact=("prefixe", "prefixe_section", "ccopre"),
            contains_any=("prefix", "ccopre"),
        )
        col_section = _pick_column(
            columns,
            exact=("section", "ccosec"),
            contains_any=("section", "ccosec"),
        )
        col_numero = _pick_column(
            columns,
            exact=("numero", "numero_parcelle", "dnupla"),
            contains_any=("numero_parcelle", "dnupla"),
        )
        col_denom = _pick_column(
            columns,
            exact=("denomination", "denomination_proprietaire", "dforme"),
            contains_any=("denomination", "raison_sociale", "proprietaire_nom"),
        )
        col_forme = _pick_column(
            columns,
            exact=("forme_juridique", "forme_juridique_abregee", "groupe_personne"),
            contains_any=("forme_juridique", "forme juridique", "groupe_personne"),
        )
        col_siren = _pick_column(
            columns,
            exact=("siren", "numero_siren"),
            contains_any=("siren",),
        )

        if not col_denom:
            raise RuntimeError(
                "Le fichier personnes morales est accessible mais la colonne de dénomination "
                "n'a pas pu être identifiée automatiquement."
            )

        wanted = []
        aliases = {}
        for alias, col in [
            ("code_insee", col_insee),
            ("parcelle_id", col_parcelle),
            ("prefixe", col_prefixe),
            ("section", col_section),
            ("numero", col_numero),
            ("denomination", col_denom),
            ("forme_juridique", col_forme),
            ("siren", col_siren),
        ]:
            if col:
                wanted.append(f"{_sql_ident(col)} AS {_sql_ident(alias)}")
                aliases[alias] = col

        if not wanted:
            return []

        where = ""
        params = []

        if col_insee:
            where = f"WHERE CAST({_sql_ident(col_insee)} AS VARCHAR) = ?"
            params = [str(insee)]
        elif col_parcelle:
            # Le début de l'identifiant cadastral contient le code commune INSEE
            # dans la représentation utilisée par les jeux cadastraux ouverts.
            where = f"WHERE CAST({_sql_ident(col_parcelle)} AS VARCHAR) LIKE ?"
            params = [f"{insee}%"]

        query = f"SELECT {', '.join(wanted)} FROM {parquet_sql} {where}"
        df = con.execute(query, params).fetchdf()
        return df.to_dict("records")
    finally:
        con.close()


def build_personnes_morales_index(insee, rows):
    """
    Indexe les propriétaires personnes morales par identifiant cadastral normalisé.
    Plusieurs propriétaires sont concaténés lorsqu'ils existent.
    """
    temp = {}

    for r in rows or []:
        pid = _normalise_parcelle_id(r.get("parcelle_id"))
        if not pid:
            pid = _canonical_parcel_id(
                insee,
                r.get("prefixe"),
                r.get("section"),
                r.get("numero"),
            )
        if not pid:
            continue

        denom = _clean_owner_text(r.get("denomination"))
        forme = _clean_owner_text(r.get("forme_juridique"))
        siren = _clean_owner_text(r.get("siren"))

        if not denom:
            continue

        key = (denom, forme, siren)
        temp.setdefault(pid, set()).add(key)

    index = {}
    for pid, owners in temp.items():
        owners = sorted(owners)
        names = [o[0] for o in owners if o[0]]
        formes = [o[1] for o in owners if o[1]]
        sirens = [o[2] for o in owners if o[2]]

        joined_name = " / ".join(dict.fromkeys(names))
        joined_forme = " / ".join(dict.fromkeys(formes))
        joined_siren = " / ".join(dict.fromkeys(sirens))

        upper = normalize_text(joined_name + " " + joined_forme)
        is_commune = any(
            token in upper
            for token in [
                "commune de ",
                "commune d ",
                "mairie",
                "ville de ",
                "commune",
            ]
        )

        index[pid] = {
            "proprietaire_personne_morale": joined_name,
            "forme_juridique_proprietaire": joined_forme,
            "siren_proprietaire": joined_siren,
            "proprietaire_commune": is_commune,
            "proprietaire_type": "Commune / collectivité" if is_commune else "Société / personne morale",
        }

    return index


def enrich_with_personnes_morales(df, insee, owner_index):
    if df is None:
        return df

    df = df.copy()
    defaults = {
        "proprietaire_personne_morale": "",
        "forme_juridique_proprietaire": "",
        "siren_proprietaire": "",
        "proprietaire_type": "",
        "proprietaire_commune": False,
    }
    for c, default in defaults.items():
        if c not in df.columns:
            df[c] = default

    if df.empty:
        return df

    for idx, row in df.iterrows():
        pid = _normalise_parcelle_id(row.get("reference"))
        info = owner_index.get(pid)
        if not info:
            # Fallback : reconstruire un identifiant à partir des champs cadastraux.
            pid2 = _canonical_parcel_id(
                insee,
                "",
                row.get("section"),
                row.get("numero"),
            )
            info = owner_index.get(pid2)

        if info:
            for k, v in info.items():
                df.at[idx, k] = v

    return df

def gpu_get(layer, params, timeout=90):
    """
    Appel robuste à l'API Carto GPU.
    Les erreurs 500/502/503/504 sont parfois temporaires : on réessaie avant abandon.
    """
    url = f"{GPU_API}/{layer}"
    last_error = None

    # GET est la voie principale documentée.
    for attempt in range(3):
        try:
            r = requests.get(url, params=params, headers=HEADERS, timeout=timeout)
            if r.ok:
                return r.json()

            last_error = requests.HTTPError(
                f"{r.status_code} {r.reason} pour {r.url}",
                response=r,
            )
            if r.status_code not in {429, 500, 502, 503, 504}:
                break
        except Exception as exc:
            last_error = exc

        time.sleep(0.7 * (attempt + 1))

    # Certaines requêtes géométriques peuvent fonctionner en POST lorsque GET échoue.
    try:
        r2 = requests.post(url, json=params, headers=HEADERS, timeout=timeout)
        if r2.ok:
            return r2.json()
        last_error = requests.HTTPError(
            f"{r2.status_code} {r2.reason} pour {r2.url}",
            response=r2,
        )
    except Exception as exc:
        last_error = exc

    if isinstance(last_error, Exception):
        raise last_error
    raise RuntimeError(f"Échec de l'appel GPU {layer}")


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def fetch_gpu_document(insee, commune_geojson, code_epci=None):
    """
    Résout le document d'urbanisme en vigueur.

    1. Essaie l'intersection ponctuelle /gpu/document?geom=...
    2. Si ce service renvoie une erreur (ex. HTTP 500), teste directement les
       partitions normalisées GPU :
       - DU_<code EPCI> pour un PLUi ;
       - DU_<code INSEE> pour un PLU communal.

    Le second chemin évite qu'une panne du seul endpoint /document bloque toute l'application.
    """
    errors = []

    # 1) Recherche géométrique normale.
    features = commune_geojson.get("features", [])
    if features:
        try:
            commune_geom = shape(features[0]["geometry"])
            point = commune_geom.representative_point()
            geom = json.dumps(mapping(point), separators=(",", ":"))

            data = gpu_get("document", {"geom": geom})
            docs = data.get("features", []) or []

            local_docs = []
            for f in docs:
                p = f.get("properties", {}) or {}
                typedoc = str(p.get("typedoc") or p.get("type_doc") or "").upper()
                if "SCOT" in typedoc:
                    continue
                local_docs.append(f)

            candidates = local_docs or docs
            if candidates:
                def doc_sort_key(f):
                    p = f.get("properties", {}) or {}
                    d = (
                        p.get("datvalid")
                        or p.get("datappro")
                        or p.get("date_appro")
                        or p.get("datefin")
                        or ""
                    )
                    return str(d)

                candidates = sorted(candidates, key=doc_sort_key, reverse=True)
                props = candidates[0].get("properties", {}) or {}
                partition = props.get("partition")
                if partition:
                    return partition, candidates
        except Exception as exc:
            errors.append(f"recherche géométrique: {exc}")

    # 2) Fallback par partitions. Pour les PLUi, le GPU utilise DU_<SIREN EPCI>.
    partition_candidates = []
    if code_epci:
        partition_candidates.append(f"DU_{code_epci}")
    partition_candidates.append(f"DU_{insee}")

    seen = set()
    for partition in partition_candidates:
        if not partition or partition in seen:
            continue
        seen.add(partition)

        try:
            # On filtre déjà sur la commune pour éviter de charger tout le PLUi.
            probe = gpu_get(
                "zone-urba",
                {"partition": partition, "insee": str(insee)},
                timeout=90,
            )
            feats = probe.get("features", []) or []
            if feats:
                return partition, []
        except Exception as exc:
            errors.append(f"{partition}: {exc}")

    details = " | ".join(errors[-4:]) if errors else "aucun zonage trouvé"
    raise RuntimeError(
        "Impossible de résoudre le document PLU/PLUi. "
        f"Partitions testées : {', '.join(partition_candidates)}. Détail : {details}"
    )


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def fetch_gpu_zones(partition, insee=None):
    # API GPU : limite haute de plusieurs milliers d'objets.
    all_features = []
    start = 0
    while True:
        params = {"partition": partition}
        if insee:
            params["insee"] = str(insee)
        if start:
            params["_start"] = start

        data = gpu_get("zone-urba", params)
        feats = data.get("features", []) or []
        all_features.extend(feats)

        returned = data.get("numberReturned", len(feats))
        total = data.get("totalFeatures", data.get("numberMatched"))

        if not feats or returned == 0:
            break
        if total is not None and len(all_features) >= int(total):
            break
        if len(feats) < 1000:
            break

        start += len(feats)
        if start >= 5000:
            break

    return {"type": "FeatureCollection", "features": all_features}


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def fetch_gpu_municipality(insee):
    try:
        return gpu_get("municipality", {"insee": insee})
    except Exception:
        return {"type": "FeatureCollection", "features": []}


@st.cache_data(ttl=30 * 24 * 3600, show_spinner=False)
def reverse_geocode(lat, lon, citycode=None):
    params = {
        "lat": lat,
        "lon": lon,
        "index": "address",
        "limit": 1,
    }
    if citycode:
        params["citycode"] = citycode
    data = http_get_json(f"{GEOCODAGE_API}/reverse", params=params, timeout=30)
    feats = data.get("features", []) or []
    if not feats:
        return ""
    p = feats[0].get("properties", {}) or {}
    return (
        p.get("label")
        or p.get("name")
        or " ".join(
            str(x)
            for x in [p.get("housenumber"), p.get("street"), p.get("postcode"), p.get("city")]
            if x
        )
    )


# --------------------------
# Helpers géographiques
# --------------------------
def _norm_rule_text(value):
    txt = str(value or "").replace("\xa0", " ")
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    txt = txt.lower()
    txt = txt.replace("’", "'").replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", txt).strip()


def _resolve_pdf_bytes(url):
    if not url:
        raise RuntimeError("URL de règlement absente.")

    r = requests.get(url, headers=HEADERS, timeout=120)
    r.raise_for_status()
    ctype = (r.headers.get("content-type") or "").lower()

    if r.content[:4] == b"%PDF" or "application/pdf" in ctype:
        return r.content, r.url

    html = r.text
    links = re.findall(
        r"href\s*=\s*[\"']([^\"']+\.pdf(?:\?[^\"']*)?)[\"']",
        html,
        flags=re.I,
    )
    if not links:
        raise RuntimeError("URLFIC ne renvoie pas de PDF exploitable.")

    pdf_url = urljoin(r.url, links[0])
    r2 = requests.get(pdf_url, headers=HEADERS, timeout=120)
    r2.raise_for_status()
    if r2.content[:4] != b"%PDF" and "application/pdf" not in (
        r2.headers.get("content-type") or ""
    ).lower():
        raise RuntimeError("Le règlement trouvé n'est pas un PDF exploitable.")
    return r2.content, r2.url


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def fetch_plu_regulation_pages(url):
    pdf_bytes, final_url = _resolve_pdf_bytes(url)
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pages = []
    for i, page in enumerate(reader.pages):
        try:
            txt = page.extract_text() or ""
        except Exception:
            txt = ""
        pages.append({"page": i + 1, "text": txt, "norm": _norm_rule_text(txt)})

    if not any(p["norm"] for p in pages):
        raise RuntimeError("Le PDF ne contient pas de texte extractible automatiquement.")
    return pages, final_url


def _zone_page_candidates(pages, zone_label, zone_long, commune_name):
    label = _norm_rule_text(zone_label)
    long_label = _norm_rule_text(zone_long)
    commune = _norm_rule_text(commune_name)

    hits = set()
    for i, p in enumerate(pages):
        txt = p["norm"]
        score = 0

        if long_label and len(long_label) >= 6 and long_label in txt:
            score += 5

        if label and len(label) >= 2:
            if re.search(rf"\bzone\s+{re.escape(label)}\b", txt):
                score += 5
            elif re.search(rf"\bsecteur\s+{re.escape(label)}\b", txt):
                score += 4
            elif re.search(rf"\b{re.escape(label)}\b", txt):
                score += 2

        if commune and commune in txt:
            score += 1

        if ("emprise au sol" in txt or "coefficient d'emprise" in txt) and (
            "hauteur" in txt or "niveau" in txt
        ):
            score += 1

        if score >= 3:
            hits.add(i)

    # Si URLFIC pointe déjà vers un petit règlement de zone, utiliser tout le document.
    if not hits and len(pages) <= 40:
        hits = set(range(len(pages)))

    expanded = set()
    for i in hits:
        for j in range(max(0, i - 2), min(len(pages), i + 4)):
            expanded.add(j)

    return [pages[i] for i in sorted(expanded)]


def _candidate_windows(text, keyword_regex, before=220, after=520):
    norm = _norm_rule_text(text)
    windows = []
    for m in re.finditer(keyword_regex, norm, flags=re.I):
        windows.append(norm[max(0, m.start()-before):min(len(norm), m.end()+after)])
    return windows


def _extract_emprise_rule(context_text, commune_name="", zone_label=""):
    norm = _norm_rule_text(context_text)
    commune = _norm_rule_text(commune_name)
    zone = _norm_rule_text(zone_label)

    windows = _candidate_windows(
        norm,
        r"(?:emprise\s+au\s+sol|coefficient\s+d[' ]?emprise\s+au\s+sol|\bces\b)",
        before=260,
        after=700,
    )
    candidates = []

    for w in windows:
        if any(x in w for x in ["non reglementee", "non reglemente", "sans objet"]):
            continue

        base_score = 0
        if any(x in w for x in [
            "maximum", "maximale", "ne peut exceder", "ne doit pas exceder",
            "limitee a", "limite a", "est fixee a", "au maximum"
        ]):
            base_score += 5
        if commune and commune in w:
            base_score += 2
        if zone and re.search(rf"\b{re.escape(zone)}\b", w):
            base_score += 2

        for m in re.finditer(r"(?<!\d)(\d{1,3}(?:[.,]\d+)?)\s*%", w):
            val = float(m.group(1).replace(",", "."))
            if 1 <= val <= 100:
                around = w[max(0, m.start()-120):m.end()+120]
                score = base_score + (4 if "emprise" in around else 0)
                candidates.append((score, val, around))

        for m in re.finditer(
            r"(?:ces|coefficient\s+d[' ]?emprise\s+au\s+sol)"
            r".{0,120}?(?:=|:|fixe(?:e)?\s+a)?\s*(0[.,]\d+)",
            w,
        ):
            val = float(m.group(1).replace(",", ".")) * 100
            if 1 <= val <= 100:
                candidates.append((base_score + 7, val, w[max(0,m.start()-80):m.end()+120]))

    if not candidates:
        return None, 0, ""

    candidates.sort(key=lambda x: (-x[0], x[1]))
    score, val, excerpt = candidates[0]
    confidence = 92 if score >= 9 else 78 if score >= 6 else 60
    return round(val, 2), confidence, excerpt[:650]


def _extract_levels_rule(context_text, floor_height_m, commune_name="", zone_label=""):
    norm = _norm_rule_text(context_text)
    commune = _norm_rule_text(commune_name)
    zone = _norm_rule_text(zone_label)

    windows = _candidate_windows(
        norm,
        r"(?:hauteur|nombre\s+de\s+niveaux|niveaux?|r\s*\+\s*\d+)",
        before=260,
        after=760,
    )
    direct = []
    heights = []

    for w in windows:
        if "cloture" in w and not any(
            x in w for x in ["construction", "batiment", "facade", "egout", "acrot"]
        ):
            continue

        base = 0
        if any(x in w for x in ["maximum", "maximale", "ne peut exceder", "limitee"]):
            base += 4
        if commune and commune in w:
            base += 2
        if zone and re.search(rf"\b{re.escape(zone)}\b", w):
            base += 2

        for m in re.finditer(r"\br\s*\+\s*(\d{1,2})\b", w):
            n = int(m.group(1)) + 1
            if 1 <= n <= 20:
                direct.append((base + 8, n, "R+N explicite", w[max(0,m.start()-140):m.end()+180]))

        for pat in [
            r"(\d{1,2})\s+niveaux?(?:\s+maximum)?",
            r"(?:maximum|maximale?)\s+(?:de\s+)?(\d{1,2})\s+niveaux?",
            r"rez[- ]de[- ]chaussee.{0,80}?(\d{1,2})\s+etages?",
        ]:
            for m in re.finditer(pat, w):
                n = int(m.group(1))
                if "rez" in m.group(0):
                    n += 1
                if 1 <= n <= 20:
                    direct.append((base + 7, n, "nombre de niveaux explicite", w[max(0,m.start()-140):m.end()+180]))

        for m in re.finditer(
            r"(?:hauteur.{0,180}?(?:maximum|maximale|ne\s+peut\s+exceder|limitee\s+a)?"
            r".{0,100}?)(\d{1,2}(?:[.,]\d+)?)\s*m(?:etre)?s?\b",
            w,
        ):
            h = float(m.group(1).replace(",", "."))
            if 2.5 <= h <= 60:
                around = w[max(0,m.start()-180):m.end()+180]
                score = base + 3
                if any(x in around for x in ["egout", "acrot", "facade"]):
                    score += 4
                if "faitage" in around and not any(x in around for x in ["egout", "acrot"]):
                    score -= 2
                heights.append((score, h, around))

    if direct:
        direct.sort(key=lambda x: (-x[0], -x[1]))
        score, n, method, excerpt = direct[0]
        confidence = 95 if score >= 12 else 85 if score >= 8 else 72
        return n, None, confidence, method, excerpt[:650]

    if heights:
        heights.sort(key=lambda x: (-x[0], x[1]))
        score, h, excerpt = heights[0]
        n = max(1, int(math.floor(h / max(float(floor_height_m), 0.1))))
        confidence = 65 if score >= 7 else 50
        return n, round(h, 2), confidence, "estimé depuis hauteur PLU", excerpt[:650]

    return None, None, 0, "", ""


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def extract_plu_zone_rule(reglement_url, zone_label, zone_long, commune_name, floor_height_m):
    if not reglement_url:
        return {
            "emprise_plu_pct": None, "niveaux_plu": None, "hauteur_plu_m": None,
            "regle_plu_confiance": 0, "regle_plu_methode": "URLFIC absente",
            "regle_plu_extrait": "", "reglement_final_url": "",
        }

    try:
        pages, final_url = fetch_plu_regulation_pages(reglement_url)
        candidates = _zone_page_candidates(pages, zone_label, zone_long, commune_name)
        if not candidates:
            return {
                "emprise_plu_pct": None, "niveaux_plu": None, "hauteur_plu_m": None,
                "regle_plu_confiance": 0,
                "regle_plu_methode": "zone non localisée dans le règlement",
                "regle_plu_extrait": "", "reglement_final_url": final_url,
            }

        context = "\n".join(p["text"] for p in candidates)
        emprise, c1, ex1 = _extract_emprise_rule(context, commune_name, zone_label)
        levels, height, c2, method2, ex2 = _extract_levels_rule(
            context, floor_height_m, commune_name, zone_label
        )

        if emprise is not None and levels is not None:
            status = f"emprise + niveaux ({method2})"
            confidence = min(c1, c2)
        elif emprise is not None:
            status = "emprise trouvée, niveaux non trouvés"
            confidence = min(c1, 45)
        elif levels is not None:
            status = f"niveaux trouvés ({method2}), emprise non trouvée"
            confidence = min(c2, 45)
        else:
            status = "règles emprise/niveaux non extraites"
            confidence = 0

        return {
            "emprise_plu_pct": emprise,
            "niveaux_plu": levels,
            "hauteur_plu_m": height,
            "regle_plu_confiance": int(confidence),
            "regle_plu_methode": status,
            "regle_plu_extrait": " | ".join(x for x in [ex1, ex2] if x)[:1100],
            "reglement_final_url": final_url,
        }
    except Exception as exc:
        return {
            "emprise_plu_pct": None, "niveaux_plu": None, "hauteur_plu_m": None,
            "regle_plu_confiance": 0,
            "regle_plu_methode": f"erreur lecture règlement: {exc}",
            "regle_plu_extrait": "", "reglement_final_url": reglement_url,
        }


def enrich_with_plu_capacity_rules(
    df, commune_name, floor_height_m, ratio_sdp_pct, ratio_shab_pct, shab_par_logement
):
    if df is None:
        return df, pd.DataFrame()
    df = df.copy()

    for col, default in {
        "emprise_plu_pct": None,
        "emprise_constructible_m2": None,
        "niveaux_plu": None,
        "hauteur_plu_m": None,
        "surface_brute_m2": None,
        "sdp_estimee_m2": None,
        "shab_estimee_m2": None,
        "logements_estimes": None,
        "regle_plu_confiance": 0,
        "regle_plu_methode": "",
        "regle_plu_extrait": "",
        "reglement_final_url": "",
    }.items():
        if col not in df.columns:
            df[col] = default

    if df.empty:
        return df, pd.DataFrame()

    zone_cols = ["zone_plu", "zone_description", "reglement_url"]
    unique_zones = df[zone_cols].fillna("").drop_duplicates().to_dict("records")

    rules = []
    for z in unique_zones:
        rule = extract_plu_zone_rule(
            z["reglement_url"], z["zone_plu"], z["zone_description"],
            commune_name, float(floor_height_m)
        )
        rules.append({**z, **rule})

    rule_df = pd.DataFrame(rules)
    rule_index = {
        (str(r.get("zone_plu") or ""), str(r.get("zone_description") or ""), str(r.get("reglement_url") or "")): r
        for r in rules
    }

    for idx, row in df.iterrows():
        key = (
            str(row.get("zone_plu") or ""),
            str(row.get("zone_description") or ""),
            str(row.get("reglement_url") or ""),
        )
        rule = rule_index.get(key, {})
        emprise = rule.get("emprise_plu_pct")
        levels = rule.get("niveaux_plu")

        for c in [
            "emprise_plu_pct", "niveaux_plu", "hauteur_plu_m",
            "regle_plu_confiance", "regle_plu_methode",
            "regle_plu_extrait", "reglement_final_url"
        ]:
            df.at[idx, c] = rule.get(c)

        if emprise is None or levels is None:
            for c in [
                "emprise_constructible_m2", "surface_brute_m2",
                "sdp_estimee_m2", "shab_estimee_m2", "logements_estimes"
            ]:
                df.at[idx, c] = None
            continue

        terrain = float(row.get("surface_m2") or 0)
        footprint = terrain * float(emprise) / 100.0
        gross = footprint * float(levels)
        sdp = gross * float(ratio_sdp_pct) / 100.0
        shab = sdp * float(ratio_shab_pct) / 100.0
        logements = math.floor(shab / max(float(shab_par_logement), 1.0))

        df.at[idx, "emprise_constructible_m2"] = round(footprint)
        df.at[idx, "surface_brute_m2"] = round(gross)
        df.at[idx, "sdp_estimee_m2"] = round(sdp)
        df.at[idx, "shab_estimee_m2"] = round(shab)
        df.at[idx, "logements_estimes"] = int(max(0, logements))

    return df, rule_df

def geodesic_area_m2(geom):
    try:
        area, _ = GEOD.geometry_area_perimeter(geom)
        return abs(float(area))
    except Exception:
        return 0.0


def valid_shape(feature):
    try:
        g = shape(feature["geometry"])
        if g.is_empty:
            return None
        if not g.is_valid:
            g = g.buffer(0)
        return g if not g.is_empty else None
    except Exception:
        return None


def normalize_text(s):
    s = str(s or "").lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip()


NON_RESIDENTIAL_KEYWORDS = [
    "activite",
    "activites",
    "industri",
    "artisan",
    "econom",
    "commercial",
    "commerce",
    "logistique",
    "portuaire",
    "aeroport",
    "ferroviaire",
    "carriere",
    "dechet",
    "camping",
    "equipement",
    "sport",
    "tourisme",
    "loisir",
    "agricol",
    "naturel",
]

HABITAT_KEYWORDS = [
    "habitat",
    "habitation",
    "residentiel",
    "residentielle",
    "resident",
    "logement",
    "logements",
    "mixte habitat",
    "mixte resident",
]


ECONOMIC_VOCATION_KEYWORDS = [
    "zone d activite",
    "zone d activites",
    "secteur d activite",
    "secteur d activites",
    "activite economique",
    "activites economiques",
    "vocation economique",
    "zone economique",
    "parc d activite",
    "parc d activites",
    "zone industrielle",
    "secteur industriel",
    "zone artisanale",
    "secteur artisanal",
    "zone commerciale",
    "secteur commercial",
    "zone logistique",
    "secteur logistique",
    "plateforme logistique",
    "zone tertiaire",
    "secteur tertiaire",
    "pole economique",
    "pole d activite",
    "pole d activites",
]


def detect_economic_vocation(props):
    """
    Exclusion forte des secteurs dont la vocation dominante est l'activité économique.

    Sources structurées :
    - ancien standard CNIG : DESTDOMI=02 = activité ;
    - standard CNIG récent : FORMDOMI 0200 à 0203 = activité /
      industrielle-logistique-commerciale / commerces / bureaux.

    En complément, on analyse le libellé et le nom long du zonage.
    Les secteurs mixtes habitat/activité (DESTDOMI=03 / FORMDOMI=0300)
    ne sont pas considérés comme purement économiques par cette fonction.
    """
    destdomi = str(props.get("destdomi") or "").strip()
    formdomi = str(props.get("formdomi") or "").strip()
    libelle = str(props.get("libelle") or "")
    libelong = str(props.get("libelong") or "")
    text = normalize_text(" ".join([libelle, libelong]))

    dd = re.sub(r"\D", "", destdomi)
    fd = re.sub(r"\D", "", formdomi)

    if dd in {"2", "02"}:
        return True, "DESTDOMI=02 : vocation dominante activité"

    if fd.startswith("02") and len(fd) >= 2:
        labels = {
            "0200": "activité",
            "0201": "activité industrielle / logistique / commerciale",
            "0202": "activité commerces",
            "0203": "activité bureaux",
        }
        return True, f"FORMDOMI={formdomi} : {labels.get(fd, 'activité économique')}"

    for keyword in ECONOMIC_VOCATION_KEYWORDS:
        if keyword in text:
            return True, f"Libellé PLU : {libelle or libelong}"

    return False, ""

# Standard CNIG 2024/2025 :
# 20 = destination Habitation ; 21 = sous-destination Logement ; 22 = Hébergement.
DESTINATION_CODES = [
    "10", "11", "12",
    "20", "21", "22",
    "30", "31", "32", "33", "34", "35", "36", "37",
    "40", "41", "42", "43", "44", "45", "46", "47",
    "50", "51", "52", "53", "54", "55",
    "99",
]


def parse_destination_codes(value):
    """
    Extrait les codes CNIG de destination/sous-destination.
    Fonctionne avec des listes séparées ou des chaînes compactées.
    """
    if value is None:
        return set()

    if isinstance(value, (list, tuple, set)):
        txt = " ".join(str(x) for x in value)
    else:
        txt = str(value)

    found = set()
    i = 0
    while i < len(txt):
        matched = False
        for c in DESTINATION_CODES:
            if txt.startswith(c, i):
                found.add(c)
                i += len(c)
                matched = True
                break
        if not matched:
            i += 1
    return found


def _normalise_typezone(typezone):
    t = str(typezone or "").strip()
    if not t:
        return ""
    # Conserver la casse utile pour AUc/AUs
    tu = t.upper()
    if tu == "U":
        return "U"
    if tu == "AUC":
        return "AUc"
    if tu == "AUS":
        return "AUs"
    if tu == "AU":
        # Données plus anciennes : AU générique, traité comme zone à urbaniser conditionnelle.
        return "AU"
    if tu == "A":
        return "A"
    if tu == "N":
        return "N"
    return t


def analyse_habitat_zone(props, include_au=True, include_conditionnel=True):
    """
    Moteur Habitat à haute précision.

    Ordre de preuve :
    1) données structurées CNIG récentes DESTOUI / DESTCDT / DESTNON ;
    2) ancien attribut DESTDOMI (01 habitat, 03 mixte habitat/activité) ;
    3) libellé / nom long de la zone ;
    4) si aucune preuve explicite : zone classée "indéterminée", donc rejetée en mode strict.
    """
    raw_typezone = props.get("typezone")
    typezone = _normalise_typezone(raw_typezone)

    libelle = str(props.get("libelle") or "")
    libelong = str(props.get("libelong") or "")
    destdomi = str(props.get("destdomi") or "").strip()
    destoui_raw = props.get("destoui")
    destcdt_raw = props.get("destcdt")
    destnon_raw = props.get("destnon")

    text = normalize_text(" ".join([libelle, libelong]))

    destoui = parse_destination_codes(destoui_raw)
    destcdt = parse_destination_codes(destcdt_raw)
    destnon = parse_destination_codes(destnon_raw)

    # Priorité absolue : éliminer les secteurs dont la vocation dominante
    # est l'activité économique, même si certaines destinations y sont ponctuellement admises.
    economic_vocation, economic_reason = detect_economic_vocation(props)
    if economic_vocation:
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — secteur à vocation d'activités économiques",
            "habitat_preuve": economic_reason,
            "habitat_confiance": 100,
            "typezone_normalise": typezone,
            "economic_vocation": True,
        }

    # 0. Type de zone : on exclut les zones non immédiatement destinées à bâtir du logement.
    if typezone in {"A", "N"}:
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — zone agricole/naturelle",
            "habitat_preuve": f"TYPEZONE={typezone}",
            "habitat_confiance": 100,
            "typezone_normalise": typezone,
            "economic_vocation": False,
        }

    if typezone == "AUs":
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — zone AU bloquée",
            "habitat_preuve": "TYPEZONE=AUs : ouverture subordonnée à modification/révision du PLU",
            "habitat_confiance": 100,
            "typezone_normalise": typezone,
            "economic_vocation": False,
        }

    if typezone in {"AUc", "AU"} and not include_au:
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — zone AU désactivée dans les critères",
            "habitat_preuve": f"TYPEZONE={typezone}",
            "habitat_confiance": 100,
            "typezone_normalise": typezone,
            "economic_vocation": False,
        }

    # 1. Standards CNIG récents : destination Habitation / Logement.
    housing_codes = {"20", "21"}

    # Interdiction explicite prioritaire lorsqu'aucune autorisation explicite ne la contredit.
    if (destnon & housing_codes) and not (destoui & housing_codes) and not (destcdt & housing_codes):
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — logement/habitation interdit",
            "habitat_preuve": f"DESTNON={','.join(sorted(destnon & housing_codes))}",
            "habitat_confiance": 100,
            "typezone_normalise": typezone,
            "economic_vocation": False,
        }

    if destoui & housing_codes:
        return {
            "habitat_eligible": True,
            "habitat_statut": "Habitat autorisé",
            "habitat_preuve": f"DESTOUI={','.join(sorted(destoui & housing_codes))}",
            "habitat_confiance": 100,
            "typezone_normalise": typezone,
            "economic_vocation": False,
        }

    if destcdt & housing_codes:
        if include_conditionnel:
            return {
                "habitat_eligible": True,
                "habitat_statut": "Habitat autorisé sous conditions",
                "habitat_preuve": f"DESTCDT={','.join(sorted(destcdt & housing_codes))}",
                "habitat_confiance": 90,
                "typezone_normalise": typezone,
                "economic_vocation": False,
            }
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — habitat seulement sous conditions",
            "habitat_preuve": f"DESTCDT={','.join(sorted(destcdt & housing_codes))}",
            "habitat_confiance": 90,
            "typezone_normalise": typezone,
            "economic_vocation": False,
        }

    # Si les champs DEST* sont renseignés mais ne contiennent ni Habitation ni Logement,
    # on considère qu'il n'y a pas de preuve de destination logement.
    if destoui or destcdt or destnon:
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — destination logement non identifiée",
            "habitat_preuve": (
                f"DESTOUI={','.join(sorted(destoui)) or '—'} ; "
                f"DESTCDT={','.join(sorted(destcdt)) or '—'} ; "
                f"DESTNON={','.join(sorted(destnon)) or '—'}"
            ),
            "habitat_confiance": 95,
            "typezone_normalise": typezone,
            "economic_vocation": False,
        }

    # 2. Ancien standard CNIG : vocation dominante.
    # 01 = habitat ; 03 = mixte habitat/activité.
    dd = re.sub(r"\D", "", destdomi)
    if dd in {"1", "01"}:
        return {
            "habitat_eligible": True,
            "habitat_statut": "Habitat — vocation dominante",
            "habitat_preuve": "DESTDOMI=01 (habitat)",
            "habitat_confiance": 95,
            "typezone_normalise": typezone,
            "economic_vocation": False,
        }
    if dd in {"3", "03"}:
        return {
            "habitat_eligible": True,
            "habitat_statut": "Habitat/mixte — vocation dominante",
            "habitat_preuve": "DESTDOMI=03 (mixte habitat/activité)",
            "habitat_confiance": 90,
            "typezone_normalise": typezone,
            "economic_vocation": False,
        }
    if dd in {"2", "02", "4", "04", "5", "05", "7", "07", "8", "08", "9", "09", "10"}:
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — vocation dominante non habitat",
            "habitat_preuve": f"DESTDOMI={destdomi}",
            "habitat_confiance": 95,
            "typezone_normalise": typezone,
            "economic_vocation": False,
        }

    # 3. Analyse sémantique du libellé et du nom long.
    positive = any(k in text for k in HABITAT_KEYWORDS)
    negative = any(k in text for k in NON_RESIDENTIAL_KEYWORDS)

    if positive and not negative:
        return {
            "habitat_eligible": True,
            "habitat_statut": "Habitat probable — libellé PLU explicite",
            "habitat_preuve": f"{libelle} — {libelong}".strip(" —"),
            "habitat_confiance": 80,
            "typezone_normalise": typezone,
            "economic_vocation": False,
        }

    if negative and not positive:
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — libellé PLU non résidentiel",
            "habitat_preuve": f"{libelle} — {libelong}".strip(" —"),
            "habitat_confiance": 80,
            "typezone_normalise": typezone,
            "economic_vocation": False,
        }

    # 4. Haute précision : en mode Habitat strict, une zone sans preuve explicite
    # n'est pas retenue. Cela réduit les faux positifs.
    return {
        "habitat_eligible": False,
        "habitat_statut": "Exclue — destination habitat non prouvée",
        "habitat_preuve": f"{libelle} — {libelong}".strip(" —") or "Métadonnées PLU insuffisantes",
        "habitat_confiance": 60,
        "typezone_normalise": typezone,
    }


def classify_zone(props):
    """Compatibilité avec quelques affichages historiques de l'application."""
    a = analyse_habitat_zone(props, include_au=True, include_conditionnel=True)
    return a["habitat_statut"], a["habitat_eligible"]


def parcel_properties(feature):
    p = feature.get("properties", {}) or {}
    section = str(p.get("section") or "").strip()
    numero = str(p.get("numero") or "").strip()
    prefixe = str(p.get("prefixe") or "").strip()
    parcel_id = str(feature.get("id") or p.get("id") or "").strip()
    contenance = p.get("contenance")
    try:
        contenance = float(contenance)
    except Exception:
        contenance = None
    return {
        "section": section,
        "numero": numero,
        "prefixe": prefixe,
        "id_parcelle": parcel_id,
        "contenance": contenance,
    }


def analyse_commune(
    commune_name,
    insee,
    parcelles_geojson,
    batiments_geojson,
    zones_geojson,
    bdnb_parcel_index,
    include_au,
    habitat_only,
    include_conditionnel,
    ratio_sdp_pct,
    ratio_shab_pct,
    shab_par_logement,
):
    zone_rows = []
    zone_geoms = []
    for zf in zones_geojson.get("features", []):
        zg = valid_shape(zf)
        if zg is None:
            continue
        props = zf.get("properties", {}) or {}
        habitat = analyse_habitat_zone(
            props,
            include_au=include_au,
            include_conditionnel=include_conditionnel,
        )
        typezone = habitat["typezone_normalise"]

        # Sans filtre Habitat, on conserve U et AUc/AU, mais jamais A, N ou AUs.
        base_constructible = (
            typezone == "U"
            or (include_au and typezone in {"AUc", "AU"})
        )
        allowed = base_constructible
        if habitat_only:
            allowed = allowed and habitat["habitat_eligible"]

        zone_rows.append(
            {
                "geom": zg,
                "allowed": allowed,
                "typezone": typezone,
                "libelle": props.get("libelle") or "",
                "libelong": props.get("libelong") or "",
                "destdomi": props.get("destdomi") or "",
                "formdomi": props.get("formdomi") or "",
                "economic_vocation": bool(habitat.get("economic_vocation", False)),
                "destoui": props.get("destoui") or "",
                "destcdt": props.get("destcdt") or "",
                "destnon": props.get("destnon") or "",
                "zone_class": habitat["habitat_statut"],
                "habitat_eligible": habitat["habitat_eligible"],
                "habitat_statut": habitat["habitat_statut"],
                "habitat_preuve": habitat["habitat_preuve"],
                "habitat_confiance": habitat["habitat_confiance"],
                "url_reglement": props.get("urlfic") or "",
                "datvalid": props.get("datvalid") or props.get("datappro") or "",
            }
        )
        zone_geoms.append(zg)

    if not zone_geoms:
        return pd.DataFrame(), []

    zone_tree = STRtree(zone_geoms)

    # On indexe les centroïdes des bâtiments : si leur centroïde tombe dans la parcelle,
    # on considère la parcelle comme bâtie.
    building_geoms = []
    building_centroids = []
    for bf in batiments_geojson.get("features", []):
        bg = valid_shape(bf)
        if bg is None:
            continue
        building_geoms.append(bg)
        building_centroids.append(bg.representative_point())

    building_tree = STRtree(building_centroids) if building_centroids else None

    rows = []
    feature_map = {}

    for pf in parcelles_geojson.get("features", []):
        pg = valid_shape(pf)
        if pg is None:
            continue

        rp = pg.representative_point()
        zone_idx = zone_tree.query(rp, predicate="intersects")
        if len(zone_idx) == 0:
            continue

        # En cas de recouvrement, retenir la première zone candidate.
        zi = int(zone_idx[0])
        zr = zone_rows[zi]
        if not zr["allowed"]:
            continue

        pp = parcel_properties(pf)
        surface = pp["contenance"]
        if surface is None or surface <= 0:
            surface = geodesic_area_m2(pg)

        built_count = 0
        built_footprint = 0.0
        if building_tree is not None:
            bidx = building_tree.query(pg, predicate="contains")
            built_count = len(bidx)
            if built_count:
                for bi in bidx:
                    built_footprint += geodesic_area_m2(building_geoms[int(bi)])

        terrain_bati = built_count > 0

        raw_ref = pp["id_parcelle"] or f"{insee}-{pp['section']}-{pp['numero']}"
        ref_norm = _normalise_parcelle_id(raw_ref)
        bdnb_info = detect_collective_housing(bdnb_parcel_index.get(ref_norm, []))

        # La capacité sera calculée après lecture du règlement PLU.
        surface_brute = None
        sdp_estimee = None
        shab_estimee = None
        logements = None

        # Score de présélection, pas un score juridique.
        score = 50
        if zr["zone_class"].startswith("U habitat"):
            score += 20
        elif zr["typezone"] == "U":
            score += 10
        if surface >= 1500:
            score += 10
        if surface >= 3000:
            score += 5
        if not terrain_bati:
            score += 5
        if zr["typezone"] == "AU":
            score -= 10
        score = max(0, min(100, score))

        ref = raw_ref
        rows.append(
            {
                "selection": False,
                "ville": commune_name,
                "code_insee": insee,
                "reference": ref,
                "section": pp["section"],
                "numero": pp["numero"],
                "surface_m2": round(surface),
                "emprise_plu_pct": None,
                "emprise_constructible_m2": None,
                "niveaux_plu": None,
                "hauteur_plu_m": None,
                "surface_brute_m2": None,
                "sdp_estimee_m2": None,
                "shab_estimee_m2": None,
                "regle_plu_confiance": 0,
                "regle_plu_methode": "",
                "regle_plu_extrait": "",
                "reglement_final_url": "",
                "ratio_sdp_pct": float(ratio_sdp_pct),
                "ratio_shab_pct": float(ratio_shab_pct),
                "shab_par_logement": float(shab_par_logement),
                "terrain_bati": terrain_bati,
                "nb_batiments": built_count,
                "emprise_batie_m2": round(built_footprint),
                "collectif_existant": bool(bdnb_info["collectif_existant"]),
                "usage_bdnb": bdnb_info["usage_bdnb"],
                "zone_type": zr["typezone"],
                "zone_plu": zr["libelle"],
                "zone_description": zr["libelong"],
                "classe_zone": zr["zone_class"],
                "habitat_eligible": bool(zr["habitat_eligible"]),
                "habitat_statut": zr["habitat_statut"],
                "habitat_preuve": zr["habitat_preuve"],
                "habitat_confiance": int(zr["habitat_confiance"]),
                "destoui": zr["destoui"],
                "destcdt": zr["destcdt"],
                "destnon": zr["destnon"],
                "destdomi": zr["destdomi"],
                "formdomi": zr["formdomi"],
                "economic_vocation": bool(zr["economic_vocation"]),
                "reglement_url": zr["url_reglement"],
                "date_zone": zr["datvalid"],
                "logements_estimes": logements,
                "score": score,
                "latitude": rp.y,
                "longitude": rp.x,
                "adresse": bdnb_info["adresse_bdnb"],
            }
        )
        feature_map[ref] = {
            "type": "Feature",
            "properties": {
                "reference": ref,
                "section": pp["section"],
                "numero": pp["numero"],
                "surface_m2": round(surface),
                "sdp_m2": None,
                "shab_m2": None,
                "emprise_plu_pct": None,
                "niveaux_plu": None,
                "surface_brute_m2": None,
                "zone": zr["libelle"],
                "typezone": zr["typezone"],
                "habitat": zr["habitat_statut"],
                "vocation_economique": "Oui" if zr["economic_vocation"] else "Non",
                "confiance_habitat": zr["habitat_confiance"],
                "collectif": "Oui" if bdnb_info["collectif_existant"] else "Non",
                "logements": logements,
            },
            "geometry": mapping(pg),
        }

    expected_columns = [
        "selection",
        "ville",
        "code_insee",
        "reference",
        "section",
        "numero",
        "surface_m2",
        "emprise_plu_pct",
        "emprise_constructible_m2",
        "niveaux_plu",
        "hauteur_plu_m",
        "surface_brute_m2",
        "sdp_estimee_m2",
        "shab_estimee_m2",
        "ratio_sdp_pct",
        "ratio_shab_pct",
        "shab_par_logement",
        "terrain_bati",
        "nb_batiments",
        "emprise_batie_m2",
        "collectif_existant",
        "usage_bdnb",
        "zone_type",
        "zone_plu",
        "zone_description",
        "classe_zone",
        "habitat_eligible",
        "habitat_statut",
        "habitat_preuve",
        "habitat_confiance",
        "destoui",
        "destcdt",
        "destnon",
        "destdomi",
        "formdomi",
        "economic_vocation",
        "reglement_url",
        "reglement_final_url",
        "regle_plu_confiance",
        "regle_plu_methode",
        "regle_plu_extrait",
        "date_zone",
        "logements_estimes",
        "score",
        "latitude",
        "longitude",
        "adresse",
        "proprietaire_personne_morale",
        "forme_juridique_proprietaire",
        "siren_proprietaire",
        "proprietaire_type",
        "proprietaire_commune",
    ]
    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(columns=expected_columns)
    else:
        for col in expected_columns:
            if col not in df.columns:
                df[col] = None
        df = df[expected_columns]

    return df, feature_map


# --------------------------
# Courriers
# --------------------------
def replace_text_in_runs(paragraph, replacements):
    full = "".join(run.text for run in paragraph.runs)
    if not full:
        return
    new = full
    for old, repl in replacements.items():
        if old in new:
            new = new.replace(old, repl)
    if new != full:
        if paragraph.runs:
            paragraph.runs[0].text = new
            for r in paragraph.runs[1:]:
                r.text = ""
        else:
            paragraph.text = new


def generate_letter(row, signataire, fonction, email, ville_signature):
    if not LETTER_TEMPLATE.exists():
        raise FileNotFoundError(
            "Le fichier lettre_sagec_modele.docx doit être présent à côté de app.py."
        )

    doc = Document(LETTER_TEMPLATE)

    adresse = str(row.get("adresse") or "").strip()
    if not adresse:
        adresse = f"Parcelle section {row['section']} n° {row['numero']}, {row['ville']}"

    ville = str(row["ville"]).strip()
    section = str(row["section"]).strip()
    numero = str(row["numero"]).strip()

    replacements = {
        "Adresse….": adresse,
        "Adresse...": adresse,
        "Ville……": ville,
        "Ville......": ville,
        "Anglet, le 19/05/2026.": f"{ville_signature}, le {date.today().strftime('%d/%m/%Y')}.",
        "Objet : Votre propriété à …………….": f"Objet : Votre propriété à {adresse}",
        "votre propriété située à ………": f"votre propriété située à {adresse}",
        "cadastrée section………………………,": f"cadastrée section {section} n° {numero},",
        "Nicolas PEDROT": signataire,
        "Directeur": fonction,
        "Nicolas.pedrot@sagec.fr": email,
    }

    for p in doc.paragraphs:
        replace_text_in_runs(p, replacements)
    for table in doc.tables:
        for row_ in table.rows:
            for cell in row_.cells:
                for p in cell.paragraphs:
                    replace_text_in_runs(p, replacements)

    for p in doc.paragraphs:
        t = p.text
        if "Objet : Votre propriété à" in t and adresse not in t:
            p.text = f"Objet : Votre propriété à {adresse}"
        if "Dans le cadre de nos recherches foncières" in t:
            p.text = (
                f"Dans le cadre de nos recherches foncières, nous avons identifié que votre propriété "
                f"située à {adresse}, cadastrée section {section} n° {numero}, comme pouvant présenter "
                f"un fort potentiel pour le développement d'une opération immobilière."
            )

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


# --------------------------
# Interface
# --------------------------
st.title("🏗️ Prospecteur Foncier — V3.4 — Gabarit PLU")
st.caption(
    "Cadastre réel + PLU/PLUi + extraction emprise/niveaux + propriétaires personnes morales + calcul de capacité."
)

st.info(
    "Cette V2 ne travaille plus sur Bayonne/Anglet en dur. Les communes sont chargées "
    "depuis l'API administrative officielle. Les parcelles et bâtiments viennent du cadastre réel."
)

with st.sidebar:
    st.header("Territoire")
    region_choice = st.selectbox(
        "Périmètre",
        ["Nouvelle-Aquitaine", "Occitanie", "Tout le Sud-Ouest"],
    )

    if region_choice == "Tout le Sud-Ouest":
        deps = {}
        for d in REGIONS.values():
            deps.update(d)
    else:
        deps = REGIONS[region_choice]

    dep_label_to_code = {f"{code} — {name}": code for code, name in deps.items()}
    dep_label = st.selectbox("Département", list(dep_label_to_code.keys()))
    dep_code = dep_label_to_code[dep_label]

    try:
        communes = fetch_communes(dep_code)
    except Exception as exc:
        st.error(f"Impossible de charger les communes : {exc}")
        st.stop()

    commune_labels = {
        f"{c['nom']} ({c['code']})": c for c in communes
    }
    commune_label = st.selectbox(
        "Commune",
        list(commune_labels.keys()),
        index=0,
    )
    commune = commune_labels[commune_label]
    commune_name = commune["nom"]
    insee = commune["code"]
    code_epci = commune.get("codeEpci")

st.subheader("1. Critères de recherche")

c1, c2 = st.columns(2)
with c1:
    min_log = st.number_input("Logements minimum", min_value=0, value=20, step=1)
with c2:
    max_log = st.number_input("Logements maximum", min_value=0, value=50, step=1)

st.markdown("#### Ratios de calcul de capacité")
r1, r2, r3, r4 = st.columns(4)
with r1:
    ratio_sdp_pct = st.number_input(
        "SDP / surface brute (%)",
        min_value=1.0,
        max_value=100.0,
        value=80.0,
        step=1.0,
        help="SDP estimée = surface brute construite × ce ratio.",
    )
with r2:
    ratio_shab_pct = st.number_input(
        "SHAB / SDP (%)",
        min_value=1.0,
        max_value=100.0,
        value=80.0,
        step=1.0,
        help="SHAB estimée = SDP × ce ratio.",
    )
with r3:
    shab_par_logement = st.number_input(
        "SHAB moyenne / logement (m²)",
        min_value=10.0,
        max_value=250.0,
        value=55.0,
        step=1.0,
        help="Par défaut : 55 m² SHAB par logement.",
    )
with r4:
    floor_height_m = st.number_input(
        "Hauteur théorique / niveau (m)",
        min_value=2.4,
        max_value=5.0,
        value=3.0,
        step=0.1,
        help=(
            "Utilisée seulement si le règlement donne une hauteur maximale en mètres "
            "sans nombre de niveaux ou R+N explicite."
        ),
    )

st.caption(
    "Calcul : emprise constructible = surface terrain × emprise PLU ; "
    "surface brute = emprise constructible × niveaux PLU ; "
    f"SDP = surface brute × {ratio_sdp_pct:.0f} % ; "
    f"SHAB = SDP × {ratio_shab_pct:.0f} % ; "
    f"logements = SHAB ÷ {shab_par_logement:.0f} m²."
)

c5, c6 = st.columns(2)
with c5:
    zone_mode = st.selectbox(
        "Zones à analyser",
        ["Zones U uniquement", "Zones U + AUc constructibles"],
        help=(
            "Les zones A et N sont éliminées. Les zones AUs (à urbaniser bloquées) "
            "sont également éliminées automatiquement."
        ),
    )
with c6:
    terrain_mode = st.selectbox(
        "Terrain",
        ["Indifférent", "Terrain nu", "Terrain bâti"],
    )

# Le filtre Habitat strict est supprimé à la demande.
# Les secteurs à vocation économique restent en revanche systématiquement exclus.
habitat_only = False
include_conditionnel = True

st.info(
    "Le logiciel lit maintenant le règlement écrit du PLU/PLUi pour chercher l'emprise maximale "
    "et le gabarit vertical. Si l'une des deux règles n'est pas identifiable, il n'invente pas "
    "de capacité pour la parcelle."
)

st.caption(
    "Filtre économique actif : les zones dont la vocation dominante est l'activité économique "
    "(activité, industrie, logistique, commerce, artisanat, bureaux, parc d'activités, etc.) "
    "sont éliminées avant même le filtre Habitat."
)

st.caption(
    "Filtre bâti : une parcelle déjà construite reste éligible, sauf si la BDNB identifie "
    "un bâtiment en « Résidentiel collectif » sur cette parcelle."
)

st.warning(
    "Le moteur élimine d'abord les secteurs à vocation d'activités économiques (DESTDOMI=02, FORMDOMI 0200–0203 et libellés explicites). Le filtre Habitat réduit ensuite les faux positifs : il exploite d'abord les champs structurés "
    "DESTOUI / DESTCDT / DESTNON du standard CNIG, puis DESTDOMI pour les anciens PLU et enfin "
    "les libellés explicites de zone. Les règles de gabarit (hauteur, retraits, emprise, stationnement, "
    "OAP, prescriptions et servitudes) restent à vérifier pour la faisabilité détaillée."
)

analyse_button = st.button(
    f"🔎 Analyser le cadastre réel de {commune_name}",
    type="primary",
)

if analyse_button:
    try:
        with st.spinner("Téléchargement des parcelles cadastrales réelles…"):
            parcelles_geojson = fetch_cadastre_layer(insee, "parcelles")
        with st.spinner("Téléchargement des bâtiments cadastraux…"):
            batiments_geojson = fetch_cadastre_layer(insee, "batiments")
        with st.spinner("Recherche des logements collectifs existants via la BDNB…"):
            try:
                bdnb_rows = fetch_bdnb_commune(insee)
                bdnb_parcel_index = build_bdnb_parcel_index(bdnb_rows)
                st.session_state["bdnb_error"] = ""
            except Exception as bdnb_exc:
                bdnb_parcel_index = {}
                st.session_state["bdnb_error"] = str(bdnb_exc)
        with st.spinner("Recherche du document d'urbanisme en vigueur…"):
            commune_geojson = fetch_cadastre_layer(insee, "communes")
            partition, documents = fetch_gpu_document(
                insee,
                commune_geojson,
                code_epci=code_epci,
            )
        with st.spinner(f"Chargement du zonage PLU/PLUi ({partition})…"):
            zones_geojson = fetch_gpu_zones(partition, insee=insee)

        if not zones_geojson.get("features"):
            municipality = fetch_gpu_municipality(insee)
            st.error(
                "Aucun zonage PLU/PLUi exploitable n'a été renvoyé pour cette commune. "
                f"Partition résolue : {partition}. Code EPCI : {code_epci or 'non disponible'}."
            )
            st.json(municipality)
            st.stop()

        with st.spinner("Croisement spatial cadastre ↔ PLU ↔ bâtiments…"):
            results, feature_map = analyse_commune(
                commune_name=commune_name,
                insee=insee,
                parcelles_geojson=parcelles_geojson,
                batiments_geojson=batiments_geojson,
                zones_geojson=zones_geojson,
                bdnb_parcel_index=bdnb_parcel_index,
                include_au=(zone_mode == "Zones U + AUc constructibles"),
                habitat_only=habitat_only,
                include_conditionnel=include_conditionnel,
                ratio_sdp_pct=ratio_sdp_pct,
                ratio_shab_pct=ratio_shab_pct,
                shab_par_logement=shab_par_logement,
            )

        # Écarter avant lecture du règlement les cibles déjà exclues.
        if not results.empty:
            if "economic_vocation" in results.columns:
                results = results[results["economic_vocation"] == False].copy()
            if "collectif_existant" in results.columns:
                results = results[results["collectif_existant"] == False].copy()

        with st.spinner("Lecture du règlement PLU : emprise au sol et niveaux…"):
            results, plu_rules = enrich_with_plu_capacity_rules(
                results,
                commune_name=commune_name,
                floor_height_m=floor_height_m,
                ratio_sdp_pct=ratio_sdp_pct,
                ratio_shab_pct=ratio_shab_pct,
                shab_par_logement=shab_par_logement,
            )
            st.session_state["plu_rules"] = plu_rules

            if feature_map and not results.empty:
                cap_by_ref = results.set_index("reference").to_dict("index")
                for ref, feat in feature_map.items():
                    row = cap_by_ref.get(ref)
                    if not row:
                        continue
                    props = feat.setdefault("properties", {})
                    props["emprise_plu_pct"] = row.get("emprise_plu_pct")
                    props["niveaux_plu"] = row.get("niveaux_plu")
                    props["surface_brute_m2"] = row.get("surface_brute_m2")
                    props["sdp_m2"] = row.get("sdp_estimee_m2")
                    props["shab_m2"] = row.get("shab_estimee_m2")
                    props["logements"] = row.get("logements_estimes")
                    props["confiance_plu"] = row.get("regle_plu_confiance")

        with st.spinner("Recherche des propriétaires personnes morales (sociétés / communes)…"):
            try:
                pm_rows = fetch_personnes_morales_commune(insee)
                pm_index = build_personnes_morales_index(insee, pm_rows)
                results = enrich_with_personnes_morales(results, insee, pm_index)

                if feature_map and not results.empty:
                    owner_by_ref = results.set_index("reference")[
                        ["proprietaire_personne_morale", "proprietaire_type"]
                    ].to_dict("index")
                    for ref, feat in feature_map.items():
                        owner = owner_by_ref.get(ref, {})
                        feat.setdefault("properties", {})["proprietaire"] = (
                            owner.get("proprietaire_personne_morale") or "Non identifié (personne morale)"
                        )
                        feat["properties"]["type_proprietaire"] = owner.get("proprietaire_type") or ""

                st.session_state["owner_error"] = ""
            except Exception as owner_exc:
                results = enrich_with_personnes_morales(results, insee, {})
                if feature_map:
                    for _, feat in feature_map.items():
                        feat.setdefault("properties", {})["proprietaire"] = "Non disponible"
                        feat["properties"]["type_proprietaire"] = ""
                st.session_state["owner_error"] = str(owner_exc)

        st.session_state["analysis_results"] = results
        st.session_state["feature_map"] = feature_map
        st.session_state["analysis_insee"] = insee
        st.session_state["analysis_commune"] = commune_name
        st.session_state["analysis_partition"] = partition

    except Exception as exc:
        st.error(
            "L'analyse n'a pas pu être terminée. Les API publiques peuvent parfois être temporairement "
            "indisponibles ou une commune peut ne pas avoir de zonage GPU exploitable."
        )
        st.exception(exc)

results = st.session_state.get("analysis_results")
if results is not None and st.session_state.get("analysis_insee") == insee:
    if results.empty or "logements_estimes" not in results.columns:
        st.subheader("2. Résultats sur le cadastre réel")
        st.warning(
            "Aucune parcelle n'a été retenue pour cette commune avec les règles actuelles "
            "(zones U/AUc, exclusion des secteurs économiques et du collectif existant)."
        )
        st.stop()

    results = results.copy()
    results["logements_estimes"] = pd.to_numeric(results["logements_estimes"], errors="coerce")
    results["emprise_plu_pct"] = pd.to_numeric(results["emprise_plu_pct"], errors="coerce")
    results["niveaux_plu"] = pd.to_numeric(results["niveaux_plu"], errors="coerce")

    calculable_mask = (
        results["emprise_plu_pct"].notna()
        & results["niveaux_plu"].notna()
        & results["logements_estimes"].notna()
    )

    filtered = results[
        calculable_mask
        & (results["logements_estimes"] >= min_log)
        & (results["logements_estimes"] <= max_log)
    ].copy()

    # Sécurité supplémentaire : les secteurs à vocation économique sont toujours exclus.
    if "economic_vocation" in filtered.columns:
        filtered = filtered[filtered["economic_vocation"] == False].copy()

    # Règle métier : on élimine les parcelles sur lesquelles
    # la BDNB identifie déjà du logement collectif.
    filtered = filtered[filtered["collectif_existant"] == False].copy()

    if terrain_mode == "Terrain nu":
        filtered = filtered[filtered["terrain_bati"] == False]
    elif terrain_mode == "Terrain bâti":
        filtered = filtered[filtered["terrain_bati"] == True]

    filtered = filtered.sort_values(
        ["score", "surface_m2"],
        ascending=[False, False],
    ).reset_index(drop=True)

    st.subheader("2. Résultats sur le cadastre réel")

    if filtered.empty:
        st.warning(
            "Le cadastre et le PLU ont bien été analysés, mais aucune parcelle ne correspond "
            "à l'ensemble des critères actuels (habitat, nombre de logements, type de terrain, collectif existant, etc.)."
        )
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Parcelles retenues", len(filtered))
    m2.metric("Parcelles avec gabarit PLU", int(calculable_mask.sum()))
    m3.metric("Parcelles avant filtre logements", len(results))
    m4.metric("Document urbanisme", st.session_state.get("analysis_partition", "—"))

    if st.session_state.get("analysis_partition", "").startswith("DU_") and code_epci:
        if st.session_state.get("analysis_partition") == f"DU_{code_epci}":
            st.caption(
                "Document détecté comme PLUi intercommunal via le code EPCI "
                f"{code_epci}. Le zonage affiché est filtré sur la commune {insee}."
            )

    if st.session_state.get("bdnb_error"):
        st.warning(
            "La BDNB n'a pas répondu pendant cette analyse. Dans ce cas, le logiciel ne peut pas "
            "garantir l'exclusion des résidences collectives. Détail : "
            + st.session_state["bdnb_error"]
        )

    if st.session_state.get("owner_error"):
        st.warning(
            "La recherche des propriétaires personnes morales n'a pas pu être chargée. "
            "Les parcelles restent analysées normalement, mais le nom du propriétaire société/commune "
            "peut être vide. Détail : " + st.session_state["owner_error"]
        )

    plu_rules = st.session_state.get("plu_rules")
    if isinstance(plu_rules, pd.DataFrame) and not plu_rules.empty:
        failed_rules = plu_rules[
            plu_rules["emprise_plu_pct"].isna() | plu_rules["niveaux_plu"].isna()
        ]
        if not failed_rules.empty:
            st.warning(
                f"{len(failed_rules)} zone(s) PLU n'ont pas fourni à la fois l'emprise et les niveaux. "
                "Leurs parcelles sont exclues du calcul automatique."
            )
        with st.expander("Audit des règles PLU extraites"):
            audit_cols = [
                "zone_plu", "zone_description", "emprise_plu_pct", "niveaux_plu",
                "hauteur_plu_m", "regle_plu_confiance", "regle_plu_methode",
                "regle_plu_extrait", "reglement_final_url",
            ]
            st.dataframe(
                plu_rules[[c for c in audit_cols if c in plu_rules.columns]],
                use_container_width=True,
                hide_index=True,
            )

    if filtered.empty:
        st.warning("Aucune parcelle ne correspond aux critères actuels.")
    else:
        # Carte réelle des parcelles.
        fmap = st.session_state.get("feature_map", {})
        map_features = []
        for ref in filtered["reference"].head(2000):
            feat = fmap.get(ref)
            if feat:
                map_features.append(feat)

        if map_features:
            fc = {"type": "FeatureCollection", "features": map_features}
            center_lat = float(filtered["latitude"].mean())
            center_lon = float(filtered["longitude"].mean())

            layer = pdk.Layer(
                "GeoJsonLayer",
                fc,
                pickable=True,
                stroked=True,
                filled=True,
                opacity=0.25,
                get_line_width=1,
            )
            deck = pdk.Deck(
                layers=[layer],
                initial_view_state=pdk.ViewState(
                    latitude=center_lat,
                    longitude=center_lon,
                    zoom=12,
                    pitch=0,
                ),
                tooltip={
                    "html": (
                        "<b>{reference}</b><br/>"
                        "Surface terrain : {surface_m2} m²<br/>"
                        "Emprise PLU : {emprise_plu_pct} %<br/>"
                        "Niveaux PLU : {niveaux_plu}<br/>"
                        "Surface brute : {surface_brute_m2} m²<br/>"
                        "SDP : {sdp_m2} m²<br/>"
                        "SHAB : {shab_m2} m²<br/>"
                        "Habitat : {habitat}<br/>"
                        "Confiance habitat : {confiance_habitat} %<br/>"
                        "Collectif existant : {collectif}<br/>"
                        "Propriétaire PM : {proprietaire}<br/>"
                        "Type propriétaire : {type_proprietaire}<br/>"
                        "Zone : {typezone} {zone}<br/>"
                        "Potentiel : {logements} logements<br/>"
                        "Confiance PLU : {confiance_plu} %"
                    )
                },
            )
            st.pydeck_chart(deck, use_container_width=True)

        st.caption(
            "La carte affiche les vraies géométries cadastrales. Pour les communes très importantes, "
            "l'affichage cartographique est limité aux 2 000 premiers résultats afin de garder l'application fluide."
        )

        st.subheader("3. Sélection des parcelles")
        table = filtered[
            [
                "selection",
                "reference",
                "section",
                "numero",
                "proprietaire_personne_morale",
                "proprietaire_type",
                "forme_juridique_proprietaire",
                "siren_proprietaire",
                "surface_m2",
                "emprise_plu_pct",
                "emprise_constructible_m2",
                "niveaux_plu",
                "hauteur_plu_m",
                "surface_brute_m2",
                "sdp_estimee_m2",
                "shab_estimee_m2",
                "terrain_bati",
                "collectif_existant",
                "usage_bdnb",
                "nb_batiments",
                "zone_type",
                "zone_plu",
                "habitat_statut",
                "habitat_preuve",
                "habitat_confiance",
                "formdomi",
                "economic_vocation",
                "classe_zone",
                "logements_estimes",
                "score",
                "regle_plu_confiance",
                "regle_plu_methode",
                "regle_plu_extrait",
                "reglement_url",
                "latitude",
                "longitude",
                "adresse",
            ]
        ].copy()

        edited = st.data_editor(
            table,
            hide_index=True,
            use_container_width=True,
            column_config={
                "selection": st.column_config.CheckboxColumn("Sélectionner"),
                "reference": st.column_config.TextColumn("Référence"),
                "section": st.column_config.TextColumn("Section"),
                "numero": st.column_config.TextColumn("N°"),
                "proprietaire_personne_morale": st.column_config.TextColumn("Propriétaire société / commune"),
                "proprietaire_type": st.column_config.TextColumn("Type propriétaire"),
                "forme_juridique_proprietaire": st.column_config.TextColumn("Forme juridique"),
                "siren_proprietaire": st.column_config.TextColumn("SIREN"),
                "surface_m2": st.column_config.NumberColumn("Surface terrain m²"),
                "emprise_plu_pct": st.column_config.NumberColumn("Emprise PLU %", format="%.1f %%"),
                "emprise_constructible_m2": st.column_config.NumberColumn("Emprise constructible m²"),
                "niveaux_plu": st.column_config.NumberColumn("Niveaux PLU"),
                "hauteur_plu_m": st.column_config.NumberColumn("Hauteur PLU m", format="%.1f"),
                "surface_brute_m2": st.column_config.NumberColumn("Surface brute m²"),
                "sdp_estimee_m2": st.column_config.NumberColumn("SDP estimée m²"),
                "shab_estimee_m2": st.column_config.NumberColumn("SHAB estimée m²"),
                "terrain_bati": st.column_config.CheckboxColumn("Bâti"),
                "collectif_existant": st.column_config.CheckboxColumn("Collectif existant"),
                "usage_bdnb": st.column_config.TextColumn("Usage BDNB"),
                "nb_batiments": st.column_config.NumberColumn("Nb bâtiments"),
                "zone_type": st.column_config.TextColumn("Type zone"),
                "zone_plu": st.column_config.TextColumn("Zone PLU"),
                "habitat_statut": st.column_config.TextColumn("Statut habitat"),
                "habitat_preuve": st.column_config.TextColumn("Preuve habitat"),
                "habitat_confiance": st.column_config.ProgressColumn(
                    "Confiance habitat", min_value=0, max_value=100
                ),
                "formdomi": st.column_config.TextColumn("Forme dominante CNIG"),
                "economic_vocation": st.column_config.CheckboxColumn("Vocation économique"),
                "classe_zone": st.column_config.TextColumn("Qualification"),
                "logements_estimes": st.column_config.NumberColumn("Logements estimés"),
                "score": st.column_config.ProgressColumn("Score", min_value=0, max_value=100),
                "regle_plu_confiance": st.column_config.ProgressColumn(
                    "Confiance PLU", min_value=0, max_value=100
                ),
                "regle_plu_methode": st.column_config.TextColumn("Méthode règle PLU"),
                "regle_plu_extrait": st.column_config.TextColumn("Extrait règlement"),
                "reglement_url": st.column_config.LinkColumn("Règlement"),
                "latitude": None,
                "longitude": None,
                "adresse": st.column_config.TextColumn("Adresse"),
            },
            disabled=[
                "reference",
                "section",
                "numero",
                "proprietaire_personne_morale",
                "proprietaire_type",
                "forme_juridique_proprietaire",
                "siren_proprietaire",
                "surface_m2",
                "emprise_plu_pct",
                "emprise_constructible_m2",
                "niveaux_plu",
                "hauteur_plu_m",
                "surface_brute_m2",
                "sdp_estimee_m2",
                "shab_estimee_m2",
                "terrain_bati",
                "collectif_existant",
                "usage_bdnb",
                "nb_batiments",
                "zone_type",
                "zone_plu",
                "habitat_statut",
                "habitat_preuve",
                "habitat_confiance",
                "formdomi",
                "economic_vocation",
                "classe_zone",
                "logements_estimes",
                "score",
                "regle_plu_confiance",
                "regle_plu_methode",
                "regle_plu_extrait",
                "reglement_url",
                "latitude",
                "longitude",
                "adresse",
            ],
            key="parcel_editor",
        )

        selected = edited[edited["selection"] == True].copy()
        st.write(f"**{len(selected)} parcelle(s) sélectionnée(s)**")

        if not selected.empty:
            if st.button("📍 Rechercher l'adresse des parcelles sélectionnées"):
                progress = st.progress(0)
                addresses = {}
                for i, (_, row) in enumerate(selected.iterrows(), start=1):
                    existing_addr = str(row.get("adresse") or "").strip()
                    if existing_addr:
                        addr = existing_addr
                    else:
                        try:
                            addr = reverse_geocode(
                                float(row["latitude"]),
                                float(row["longitude"]),
                                insee,
                            )
                        except Exception:
                            addr = ""
                    addresses[row["reference"]] = addr
                    progress.progress(i / len(selected))
                    time.sleep(0.03)

                # Conserver les adresses dans les résultats de session.
                base_results = st.session_state["analysis_results"].copy()
                for ref, addr in addresses.items():
                    base_results.loc[
                        base_results["reference"] == ref, "adresse"
                    ] = addr
                st.session_state["analysis_results"] = base_results
                st.success(
                    "Adresses recherchées. La page va réutiliser ces adresses pour les courriers."
                )
                st.rerun()

            st.subheader("4. Courriers de prospection")
            l1, l2 = st.columns(2)
            with l1:
                signataire = st.text_input("Signataire", "Nicolas PEDROT")
                fonction = st.text_input("Fonction", "Directeur")
            with l2:
                email = st.text_input("E-mail", "Nicolas.pedrot@sagec.fr")
                ville_signature = st.text_input("Ville de signature", "Anglet")

            if st.button("✉️ Générer les lettres sélectionnées"):
                # Récupérer les adresses éventuellement trouvées dans le DataFrame principal.
                latest = st.session_state["analysis_results"].set_index("reference")
                zip_buffer = io.BytesIO()
                generated = 0
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
                    for _, row in selected.iterrows():
                        ref = row["reference"]
                        if ref in latest.index:
                            current = latest.loc[ref].to_dict()
                        else:
                            current = row.to_dict()

                        letter = generate_letter(
                            current,
                            signataire,
                            fonction,
                            email,
                            ville_signature,
                        )
                        safe_ref = re.sub(r"[^A-Za-z0-9_-]+", "_", str(ref))
                        z.writestr(
                            f"Lettre_{commune_name}_{safe_ref}.docx",
                            letter.getvalue(),
                        )
                        generated += 1

                zip_buffer.seek(0)
                st.download_button(
                    f"Télécharger {generated} lettre(s)",
                    data=zip_buffer.getvalue(),
                    file_name=f"courriers_{commune_name}_{date.today().isoformat()}.zip",
                    mime="application/zip",
                )

        st.download_button(
            "⬇️ Exporter les résultats en CSV",
            data=filtered.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"prospection_{commune_name}_{insee}.csv",
            mime="text/csv",
        )

with st.expander("Ce que la V2 analyse réellement"):
    st.markdown(
        """
- **Communes** : chargées dynamiquement pour la Nouvelle-Aquitaine et l'Occitanie.
- **Cadastre** : parcelles et bâtiments réels du dernier millésime Etalab.
- **PLU / PLUi** : zonages du Géoportail de l'Urbanisme via l'API Carto IGN.
- **Exclusion économique prioritaire** : `DESTDOMI=02`, `FORMDOMI=0200/0201/0202/0203` et libellés explicites d'activités économiques sont éliminés.
- **Filtre “uniquement habitat” supprimé** : les autres zones U/AUc restent analysées dès lors qu'elles ne sont pas identifiées comme secteurs économiques.
- **Propriétaires personnes morales** : tentative de rapprochement avec le fichier open data DGFiP des parcelles détenues par des personnes morales ; affichage du nom, de la forme juridique et du SIREN lorsqu'ils sont disponibles.
- **Compatibilité anciens PLU** : lecture de `DESTDOMI` (01 habitat ; 03 mixte habitat/activité) puis analyse des libellés explicites.
- **Exclusions de zonage** : A, N et AUs sont écartées ; U et AUc sont analysées selon la destination logement.
- **Mode haute précision** : une zone ambiguë sans preuve explicite d'habitat est exclue.
- **Terrain nu / bâti** : déterminé par croisement spatial avec les bâtiments cadastraux.
- **Collectif existant** : les parcelles identifiées par la BDNB comme déjà occupées par du « Résidentiel collectif » sont exclues.
- **Autres parcelles bâties** : elles restent éligibles, quelle que soit la taille ou l'emprise du bâtiment.
- **Emprise PLU** : extraction automatique du règlement PDF lié à la zone via `URLFIC`.
- **Niveaux PLU** : priorité à `R+N` / nombre de niveaux ; à défaut conversion d'une hauteur maximale.
- **Surface brute** : surface terrain × emprise maximale PLU × nombre de niveaux.
- **Capacité logements** : SDP = surface brute × ratio SDP ; SHAB = SDP × ratio SHAB ; logements = SHAB ÷ ratio SHAB/logement.
- **Traçabilité** : méthode, confiance et extrait du règlement affichés par zone.
- **Adresse** : recherchée pour les parcelles sélectionnées via le géocodage inverse de la Géoplateforme.
- **Courrier** : généré à partir du modèle SAGEC fourni.

### Limite actuelle
Le filtre Habitat exploite désormais les destinations structurées du PLU lorsqu'elles sont disponibles.
Il ne remplace toutefois pas une faisabilité complète : le nombre de logements reste une estimation de
**présélection**. Une prochaine version pourra extraire automatiquement du règlement : emprise au sol,
hauteur, retraits, pleine terre, stationnement, OAP, prescriptions et servitudes.
        """
    )
