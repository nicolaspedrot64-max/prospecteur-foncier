import gzip
import io
import json
import math
import re
import time
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path
from urllib.parse import urlparse

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
try:
    import shapefile
except ImportError:
    shapefile = None
from pyproj import Geod, CRS, Transformer
from shapely.geometry import shape, mapping
from shapely.ops import transform as shapely_transform
from shapely.strtree import STRtree

APP_DIR = Path(__file__).parent
LETTER_TEMPLATE = APP_DIR / "lettre_sagec_modele.docx"

GEO_API = "https://geo.api.gouv.fr"
CADASTRE_BASE = "https://cadastre.data.gouv.fr/data/etalab-cadastre/latest/geojson/communes"
CADASTRE_BASE_FALLBACK = "https://files.data.gouv.fr/cadastre/etalab-cadastre/latest/geojson/communes"
GPU_API = "https://apicarto.ign.fr/api/gpu"
GEOCODAGE_API = "https://data.geopf.fr/geocodage"

GPU_ARCHIVE_API = (
    "https://www.geoportail-urbanisme.gouv.fr/api/document/"
    "download-by-partition/{partition}"
)

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
    page_title="Prospecteur Foncier V3.7",
    page_icon="🏗️",
    layout="wide",
)

# --------------------------
# Helpers réseau
# --------------------------
BDNB_API = "https://api.bdnb.io/v1/bdnb/donnees/batiment_groupe_complet"

HEADERS = {
    "User-Agent": "ProspecteurFoncier/3.7 (Streamlit; donnees publiques)",
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
def fetch_gpu_layer(layer, partition, insee=None):
    """
    Charge une couche GPU par partition et commune.
    Utilisé notamment pour prescription-surf / lin / pct.
    """
    all_features = []
    start = 0
    while True:
        params = {"partition": partition}
        if insee:
            params["insee"] = str(insee)
        if start:
            params["_start"] = start

        data = gpu_get(layer, params)
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
        if start >= 10000:
            break

    return {"type": "FeatureCollection", "features": all_features}


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def fetch_gpu_graphic_prescriptions(partition, insee=None):
    """
    Les codes CNIG recherchés sont :
    - 38 / 02 : emprise au sol maximale
    - 39 / 02 : hauteur maximale
    Les occurrences 39-50 / 39-51 sont aussi conservées lorsqu'elles existent
    dans certains standards/documents (façade / faîtage ou construction).
    """
    result = {}
    for layer in ["prescription-surf", "prescription-lin", "prescription-pct"]:
        try:
            result[layer] = fetch_gpu_layer(layer, partition, insee)
        except Exception:
            result[layer] = {"type": "FeatureCollection", "features": []}
    return result


def _zip_name_map(zf):
    return {n.lower(): n for n in zf.namelist()}


def _shape_to_wgs84(shp_obj, prj_text=""):
    geom = shape(shp_obj.__geo_interface__)
    if geom.is_empty:
        return geom

    src_crs = None
    if prj_text:
        try:
            src_crs = CRS.from_wkt(prj_text)
        except Exception:
            src_crs = None

    # Les lots CNIG métropolitains sont le plus souvent en Lambert-93.
    # Si le .prj est absent, on le déduit seulement lorsque les coordonnées
    # sont manifestement projetées.
    if src_crs is None:
        minx, miny, maxx, maxy = geom.bounds
        if abs(minx) > 180 or abs(miny) > 90:
            try:
                src_crs = CRS.from_epsg(2154)
            except Exception:
                src_crs = None

    if src_crs is not None and not src_crs.is_geographic:
        transformer = Transformer.from_crs(src_crs, CRS.from_epsg(4326), always_xy=True)
        geom = shapely_transform(transformer.transform, geom)

    return geom


def _read_shapefile_from_zip(zf, shp_name):
    if shapefile is None:
        raise RuntimeError(
            "La bibliothèque pyshp est nécessaire pour lire l'archive CNIG complète."
        )

    names = _zip_name_map(zf)
    base = shp_name[:-4]
    shp_key = shp_name.lower()
    shx_key = (base + ".shx").lower()
    dbf_key = (base + ".dbf").lower()
    prj_key = (base + ".prj").lower()

    if shp_key not in names or dbf_key not in names:
        return []

    shp_io = io.BytesIO(zf.read(names[shp_key]))
    dbf_io = io.BytesIO(zf.read(names[dbf_key]))
    shx_io = io.BytesIO(zf.read(names[shx_key])) if shx_key in names else None

    kwargs = {
        "shp": shp_io,
        "dbf": dbf_io,
        "encoding": "utf-8",
        "encodingErrors": "replace",
    }
    if shx_io is not None:
        kwargs["shx"] = shx_io

    try:
        reader = shapefile.Reader(**kwargs)
    except Exception:
        # Beaucoup d'anciens lots CNIG sont en latin-1.
        kwargs["encoding"] = "latin-1"
        shp_io.seek(0)
        dbf_io.seek(0)
        if shx_io is not None:
            shx_io.seek(0)
        reader = shapefile.Reader(**kwargs)

    prj_text = ""
    if prj_key in names:
        try:
            prj_text = zf.read(names[prj_key]).decode("utf-8", errors="replace")
        except Exception:
            prj_text = ""

    field_names = [f[0] for f in reader.fields[1:]]
    features = []
    for sr in reader.iterShapeRecords():
        props = dict(zip(field_names, list(sr.record)))
        try:
            geom = _shape_to_wgs84(sr.shape, prj_text)
            if geom.is_empty:
                continue
            features.append(
                {
                    "type": "Feature",
                    "properties": props,
                    "geometry": mapping(geom),
                }
            )
        except Exception:
            continue
    return features


def _archive_layer_name(filename):
    up = filename.upper()
    if "PRESCRIPTION_SURF" in up and filename.lower().endswith(".shp"):
        return "prescription-surf"
    if "PRESCRIPTION_LIN" in up and filename.lower().endswith(".shp"):
        return "prescription-lin"
    if "PRESCRIPTION_PCT" in up and filename.lower().endswith(".shp"):
        return "prescription-pct"
    if "INFO_SURF" in up and filename.lower().endswith(".shp"):
        return "info-surf"
    if "INFO_LIN" in up and filename.lower().endswith(".shp"):
        return "info-lin"
    if "INFO_PCT" in up and filename.lower().endswith(".shp"):
        return "info-pct"
    return None


def _xml_text_candidates(zf):
    """
    Repère la présence d'un règlement structuré XML/SRU.
    On garde le texte compact comme source secondaire pour les règles de zone.
    """
    results = []
    for name in zf.namelist():
        low = name.lower()
        if not low.endswith(".xml"):
            continue
        if "reglement" not in low:
            continue
        try:
            raw = zf.read(name)
            root = ET.fromstring(raw)
            texts = []
            for elem in root.iter():
                if elem.text and elem.text.strip():
                    texts.append(elem.text.strip())
                for k, v in elem.attrib.items():
                    if v:
                        texts.append(f"{k}={v}")
            compact = " ".join(texts)
            if compact:
                results.append({"filename": name, "text": compact})
        except Exception:
            continue
    return results


@st.cache_data(ttl=7 * 24 * 3600, show_spinner=False)
def fetch_cnig_archive_layers(partition):
    """
    Télécharge l'archive officielle GPU du document et récupère les couches CNIG
    avec TOUS leurs attributs, y compris les attributs supplémentaires LIB_ATTR/LIB_VAL
    lorsqu'ils ont été publiés par la collectivité.

    Le résultat est mis en cache une semaine et ne conserve pas l'archive ZIP brute.
    """
    url = GPU_ARCHIVE_API.format(partition=partition)
    r = requests.get(url, headers=HEADERS, timeout=240)
    r.raise_for_status()

    if r.content[:2] != b"PK":
        raise RuntimeError("Le service de téléchargement GPU n'a pas renvoyé une archive ZIP.")

    layers = {
        "prescription-surf": [],
        "prescription-lin": [],
        "prescription-pct": [],
        "info-surf": [],
        "info-lin": [],
        "info-pct": [],
    }

    with zipfile.ZipFile(io.BytesIO(r.content)) as zf:
        for name in zf.namelist():
            layer = _archive_layer_name(name)
            if not layer:
                continue
            try:
                feats = _read_shapefile_from_zip(zf, name)
                layers[layer].extend(feats)
            except Exception:
                continue

        xml_rules = _xml_text_candidates(zf)
        file_list = zf.namelist()

    return {
        "layers": layers,
        "xml_rules": xml_rules,
        "files": file_list,
        "archive_url": url,
        "archive_size_mb": round(len(r.content) / 1024 / 1024, 1),
    }


def merge_graphic_sources(api_layers, archive_bundle):
    """
    Priorité à l'archive CNIG, car elle contient potentiellement les attributs
    supplémentaires absents de l'API Carto. Si une couche archive est vide,
    on conserve la couche API.
    """
    merged = {}
    archive_layers = (archive_bundle or {}).get("layers", {})
    for layer in ["prescription-surf", "prescription-lin", "prescription-pct"]:
        af = archive_layers.get(layer, []) or []
        if af:
            merged[layer] = {"type": "FeatureCollection", "features": af}
        else:
            merged[layer] = api_layers.get(
                layer, {"type": "FeatureCollection", "features": []}
            )
    return merged


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
def _rule_norm(value):
    txt = str(value or "").replace("\xa0", " ")
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    txt = txt.lower().replace("’", "'").replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", txt).strip()


def _url_page_hint(url):
    """
    URLFIC peut contenir #page=N. C'est une information très utile :
    elle cible souvent directement l'article de la zone ou de la prescription.
    """
    try:
        fragment = urlparse(str(url or "")).fragment or ""
        m = re.search(r"(?:^|&)page=(\d+)", fragment, flags=re.I)
        return int(m.group(1)) if m else None
    except Exception:
        return None


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def _fetch_pdf_pages(url):
    if not url:
        return [], None

    page_hint = _url_page_hint(url)
    r = requests.get(url, headers=HEADERS, timeout=120)
    r.raise_for_status()

    if r.content[:4] != b"%PDF" and "application/pdf" not in (
        r.headers.get("content-type") or ""
    ).lower():
        raise RuntimeError("Le lien du règlement ne renvoie pas un PDF.")

    reader = PdfReader(io.BytesIO(r.content))
    pages = []
    for i, page in enumerate(reader.pages):
        try:
            txt = page.extract_text() or ""
        except Exception:
            txt = ""
        pages.append(
            {
                "page": i + 1,
                "text": txt,
                "norm": _rule_norm(txt),
            }
        )
    return pages, page_hint


def _page_subset_for_rule(pages, page_hint=None, zone_label="", rule_kind=""):
    if not pages:
        return []

    indexes = set()
    if page_hint:
        idx = max(0, int(page_hint) - 1)
        # Les chapitres d'une zone peuvent s'étendre sur plusieurs pages.
        for j in range(max(0, idx - 2), min(len(pages), idx + 14)):
            indexes.add(j)

    z = _rule_norm(zone_label)
    if z:
        scored = []
        for i, p in enumerate(pages):
            txt = p["norm"]
            score = 0
            if re.search(rf"\bzone\s+{re.escape(z)}\b", txt):
                score += 6
            elif re.search(rf"\bsecteur\s+{re.escape(z)}\b", txt):
                score += 5
            elif len(z) >= 2 and re.search(rf"\b{re.escape(z)}\b", txt):
                score += 2

            if rule_kind == "emprise" and (
                "emprise au sol" in txt or "coefficient d'emprise" in txt
            ):
                score += 2
            if rule_kind == "height" and (
                "hauteur" in txt or "r+" in txt or "niveaux" in txt
            ):
                score += 2

            if score >= 4:
                scored.append((score, i))

        if scored:
            best = max(s for s, _ in scored)
            for score, i in scored:
                if score >= best - 1:
                    for j in range(max(0, i - 2), min(len(pages), i + 12)):
                        indexes.add(j)

    if not indexes:
        # Petit document dédié à une zone/prescription.
        if len(pages) <= 50:
            indexes.update(range(len(pages)))
        else:
            keyword = "emprise au sol" if rule_kind == "emprise" else "hauteur"
            for i, p in enumerate(pages):
                if keyword in p["norm"]:
                    for j in range(max(0, i - 1), min(len(pages), i + 4)):
                        indexes.add(j)
                    if len(indexes) > 35:
                        break

    return [pages[i] for i in sorted(indexes)]


def _score_context(context, wanted=""):
    txt = _rule_norm(context)
    score = 0

    positive = [
        "maximum",
        "maximale",
        "ne peut exceder",
        "ne doit pas exceder",
        "est limitee a",
        "est limite a",
        "au maximum",
        "fixee a",
        "fixe a",
        "habitation",
        "logement",
    ]
    negative = [
        "annexe",
        "abri",
        "piscine",
        "local technique",
        "equipement public",
        "equipements publics",
        "service public",
        "commerce",
        "activite artisanale",
        "exploitation agricole",
    ]

    score += sum(2 for x in positive if x in txt)
    score -= sum(2 for x in negative if x in txt)

    if wanted and wanted in txt:
        score += 4
    return score


def _extract_emprise_from_text(text):
    txt = _rule_norm(text)
    candidates = []

    # Fenêtres centrées sur l'emprise/CES.
    for m in re.finditer(
        r"(?:emprise\s+au\s+sol|coefficient\s+d[' ]?emprise|\bces\b)",
        txt,
    ):
        w = txt[max(0, m.start()-260):min(len(txt), m.end()+700)]
        if "non reglement" in w or "sans objet" in w:
            continue

        for p in re.finditer(r"(?<!\d)(\d{1,3}(?:[.,]\d+)?)\s*%", w):
            val = float(p.group(1).replace(",", "."))
            if 0 < val <= 100:
                around = w[max(0, p.start()-180):p.end()+180]
                score = _score_context(around, "emprise")
                # Exiger une proximité sémantique avec l'emprise.
                if "emprise" in around or "ces" in around:
                    score += 5
                candidates.append((score, val, around))

        # CES 0,40
        for p in re.finditer(
            r"(?:ces|coefficient\s+d[' ]?emprise).{0,100}?"
            r"(?:=|:|fixe(?:e)?\s+a)?\s*(0[.,]\d+)",
            w,
        ):
            val = float(p.group(1).replace(",", ".")) * 100.0
            if 0 < val <= 100:
                around = w[max(0, p.start()-160):p.end()+180]
                candidates.append((_score_context(around, "ces") + 6, val, around))

    if not candidates:
        return None, 0, ""

    # Meilleur contexte ; à score égal, valeur la plus restrictive pour limiter les faux positifs.
    candidates.sort(key=lambda x: (-x[0], x[1]))
    score, val, excerpt = candidates[0]
    conf = 92 if score >= 10 else 82 if score >= 7 else 68
    return round(val, 2), conf, excerpt[:700]


def _extract_height_from_text(text, floor_height_m):
    txt = _rule_norm(text)
    direct = []
    heights = []

    # Fenêtres centrées sur hauteur/niveaux.
    for m in re.finditer(r"(?:hauteur|niveaux?|r\s*\+\s*\d+)", txt):
        w = txt[max(0, m.start()-280):min(len(txt), m.end()+850)]

        if "cloture" in w and not any(
            x in w for x in ["construction", "batiment", "facade", "egout", "acrot"]
        ):
            continue

        for p in re.finditer(r"\br\s*\+\s*(\d{1,2})\b", w):
            levels = int(p.group(1)) + 1
            if 1 <= levels <= 20:
                around = w[max(0, p.start()-180):p.end()+200]
                direct.append(
                    (_score_context(around, "hauteur") + 8, levels, None, around, "R+N")
                )

        patterns = [
            r"(?<!\d)(\d{1,2})\s+niveaux?\s*(?:maximum|maxi)?",
            r"(?:maximum|maximale?)\s+(?:de\s+)?(\d{1,2})\s+niveaux?",
        ]
        for pat in patterns:
            for p in re.finditer(pat, w):
                levels = int(p.group(1))
                if 1 <= levels <= 20:
                    around = w[max(0, p.start()-180):p.end()+200]
                    direct.append(
                        (_score_context(around, "niveaux") + 6, levels, None, around, "niveaux")
                    )

        # Hauteur en mètres. Écarter les petites cotes typiques de clôtures.
        for p in re.finditer(
            r"(?:hauteur.{0,180}?)(\d{1,2}(?:[.,]\d+)?)\s*m(?:etre)?s?\b",
            w,
        ):
            h = float(p.group(1).replace(",", "."))
            if 2.5 <= h <= 60:
                around = w[max(0, p.start()-180):p.end()+220]
                score = _score_context(around, "hauteur") + 3
                if any(x in around for x in ["egout", "acrot", "facade"]):
                    score += 4
                if "faitage" in around and not any(x in around for x in ["egout", "acrot"]):
                    score -= 2
                heights.append((score, h, around))

    if direct:
        direct.sort(key=lambda x: (-x[0], -x[1]))
        score, levels, _, excerpt, method = direct[0]
        conf = 95 if score >= 12 else 86 if score >= 8 else 72
        return levels, None, conf, method, excerpt[:700]

    if heights:
        heights.sort(key=lambda x: (-x[0], x[1]))
        score, h, excerpt = heights[0]
        levels = max(1, int(math.floor(h / max(float(floor_height_m), 0.1))))
        conf = 78 if score >= 10 else 66 if score >= 7 else 52
        return levels, round(h, 2), conf, "hauteur convertie en niveaux", excerpt[:700]

    return None, None, 0, "", ""


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def _rule_from_document(url, zone_label, kind, floor_height_m):
    """
    Fallback écrit ciblé :
    - priorité au #page=N de URLFIC ;
    - puis recherche du chapitre de zone ;
    - aucune valeur n'est inventée si le texte reste ambigu.
    """
    if not url:
        return {
            "value": None,
            "levels": None,
            "height_m": None,
            "confidence": 0,
            "method": "aucun règlement lié",
            "excerpt": "",
        }

    try:
        pages, page_hint = _fetch_pdf_pages(url)
        subset = _page_subset_for_rule(
            pages,
            page_hint=page_hint,
            zone_label=zone_label,
            rule_kind=kind,
        )
        text = "\n".join(p["text"] for p in subset)

        if kind == "emprise":
            val, conf, excerpt = _extract_emprise_from_text(text)
            return {
                "value": val,
                "levels": None,
                "height_m": None,
                "confidence": conf,
                "method": "règlement écrit ciblé" if val is not None else "emprise non extraite",
                "excerpt": excerpt,
            }

        levels, h, conf, method, excerpt = _extract_height_from_text(
            text,
            floor_height_m=floor_height_m,
        )
        return {
            "value": None,
            "levels": levels,
            "height_m": h,
            "confidence": conf,
            "method": f"règlement écrit ciblé — {method}" if levels is not None else "hauteur non extraite",
            "excerpt": excerpt,
        }
    except Exception as exc:
        return {
            "value": None,
            "levels": None,
            "height_m": None,
            "confidence": 0,
            "method": f"lecture règlement impossible: {exc}",
            "excerpt": "",
        }


def _complete_graphic_rule(rule, floor_height_m):
    """
    Une prescription graphique donne parfaitement le périmètre d'application.
    Si API Carto ne restitue pas LIB_ATTR/LIB_VAL, on lit uniquement le document
    pointé par URLFIC pour récupérer la valeur chiffrée.
    """
    if not rule:
        return rule

    rule = dict(rule)
    zone_hint = rule.get("libelle") or rule.get("txt") or ""

    if rule.get("typepsc") == "38" and rule.get("emprise_pct") is None:
        doc = _rule_from_document(
            rule.get("urlfic"),
            zone_hint,
            "emprise",
            floor_height_m,
        )
        if doc["value"] is not None:
            rule["emprise_pct"] = doc["value"]
            rule["confidence"] = max(rule.get("confidence", 0), doc["confidence"])
            rule["method"] = "Prescription graphique + " + doc["method"]
            rule["excerpt"] = doc["excerpt"]

    if rule.get("typepsc") == "39" and rule.get("levels") is None and rule.get("height_m") is None:
        doc = _rule_from_document(
            rule.get("urlfic"),
            zone_hint,
            "height",
            floor_height_m,
        )
        if doc["levels"] is not None:
            rule["levels"] = doc["levels"]
            rule["height_m"] = doc["height_m"]
            rule["confidence"] = max(rule.get("confidence", 0), doc["confidence"])
            rule["method"] = "Prescription graphique + " + doc["method"]
            rule["excerpt"] = doc["excerpt"]

    return rule


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def _zone_written_rule(reglement_url, zone_plu, floor_height_m):
    er = _rule_from_document(
        reglement_url,
        zone_plu,
        "emprise",
        floor_height_m,
    )
    hr = _rule_from_document(
        reglement_url,
        zone_plu,
        "height",
        floor_height_m,
    )
    return {
        "emprise_pct": er["value"],
        "emprise_conf": er["confidence"],
        "emprise_method": er["method"],
        "emprise_excerpt": er["excerpt"],
        "levels": hr["levels"],
        "height_m": hr["height_m"],
        "height_conf": hr["confidence"],
        "height_method": hr["method"],
        "height_excerpt": hr["excerpt"],
    }


def build_zone_written_rule_index(df, floor_height_m):
    """
    Une lecture par zone, jamais une lecture par parcelle.
    Le résultat est ensuite appliqué à toutes les parcelles de la même zone,
    sauf lorsqu'une prescription graphique locale vient le remplacer.
    """
    index = {}
    if df is None or df.empty:
        return index

    cols = ["zone_plu", "reglement_url"]
    for rec in df[cols].fillna("").drop_duplicates().to_dict("records"):
        key = (rec["zone_plu"], rec["reglement_url"])
        index[key] = _zone_written_rule(
            rec["reglement_url"],
            rec["zone_plu"],
            float(floor_height_m),
        )
    return index


def _overlay_piece(pieces, rule_geom, field, value, source, confidence):
    """
    Découpe réellement la parcelle : une prescription locale ne s'applique
    qu'à sa zone d'intersection, pas à toute la parcelle.
    """
    out = []
    for piece in pieces:
        geom = piece["geom"]
        try:
            inter = geom.intersection(rule_geom)
            diff = geom.difference(rule_geom)
        except Exception:
            out.append(piece)
            continue

        if not inter.is_empty and geodesic_area_m2(inter) > 0.05:
            p = dict(piece)
            p["geom"] = inter
            p[field] = value
            p[field + "_source"] = source
            p[field + "_conf"] = confidence
            out.append(p)

        if not diff.is_empty and geodesic_area_m2(diff) > 0.05:
            p = dict(piece)
            p["geom"] = diff
            out.append(p)
    return out

def _parse_distance_m(text):
    txt = _rule_norm(text)
    values = []
    for m in re.finditer(r"(?<!\d)(\d{1,3}(?:[.,]\d+)?)\s*m(?:etre)?s?\b", txt):
        val = float(m.group(1).replace(",", "."))
        if 0 <= val <= 100:
            values.append(val)
    return min(values) if values else None


def _extract_setback_from_text(text, kind):
    txt = _rule_norm(text)
    if kind == "voie":
        keywords = [
            "voies et emprises publiques",
            "voies publiques",
            "par rapport aux voies",
            "alignement",
        ]
    elif kind == "lateral":
        keywords = ["limites separatives laterales", "limite separative laterale"]
    else:
        keywords = ["fonds de parcelles", "fond de parcelle"]

    candidates = []
    for kw in keywords:
        for m in re.finditer(re.escape(kw), txt):
            w = txt[max(0, m.start()-250):min(len(txt), m.end()+700)]
            if "alignement" in w and not re.search(r"\d+\s*m", w):
                candidates.append((7, 0.0, w))
            d = _parse_distance_m(w)
            if d is not None:
                score = 5
                if any(x in w for x in ["minimum", "au moins", "recul", "distance"]):
                    score += 3
                if any(x in w for x in ["annexe", "piscine", "abri"]):
                    score -= 3
                candidates.append((score, d, w))

    if not candidates:
        return None, 0, ""
    candidates.sort(key=lambda x: (-x[0], -x[1]))
    score, value, excerpt = candidates[0]
    return value, 88 if score >= 8 else 70, excerpt[:700]


def _extract_green_ratio(text):
    txt = _rule_norm(text)
    candidates = []
    for kw in ["pleine terre", "espace vert", "espaces verts", "surface permeable"]:
        for m in re.finditer(re.escape(kw), txt):
            w = txt[max(0, m.start()-220):min(len(txt), m.end()+600)]
            for p in re.finditer(r"(?<!\d)(\d{1,3}(?:[.,]\d+)?)\s*%", w):
                val = float(p.group(1).replace(",", "."))
                if 0 <= val <= 100:
                    score = 5
                    if any(x in w for x in ["minimum", "au moins", "doit representer"]):
                        score += 3
                    candidates.append((score, val, w))
    if not candidates:
        return None, 0, ""
    candidates.sort(key=lambda x: (-x[0], -x[1]))
    score, val, excerpt = candidates[0]
    return val, 85 if score >= 8 else 65, excerpt[:650]


def _extract_parking_ratio(text):
    txt = _rule_norm(text)
    candidates = []
    # Exemples : "1 place par logement", "2 places de stationnement par logement"
    for m in re.finditer(
        r"(?<!\d)(\d+(?:[.,]\d+)?)\s+places?(?:\s+de\s+stationnement)?\s+par\s+logement",
        txt,
    ):
        val = float(m.group(1).replace(",", "."))
        if 0 <= val <= 10:
            w = txt[max(0, m.start()-180):min(len(txt), m.end()+260)]
            candidates.append((8, val, w))
    if not candidates:
        return None, 0, ""
    candidates.sort(key=lambda x: (-x[0], -x[1]))
    score, val, excerpt = candidates[0]
    return val, 82, excerpt[:600]


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def _zone_extra_written_constraints(reglement_url, zone_plu):
    if not reglement_url:
        return {}

    try:
        pages, page_hint = _fetch_pdf_pages(reglement_url)
        subset = _page_subset_for_rule(
            pages,
            page_hint=page_hint,
            zone_label=zone_plu,
            rule_kind="other",
        )
        text = "\n".join(p["text"] for p in subset)

        road, c_road, ex_road = _extract_setback_from_text(text, "voie")
        lat, c_lat, ex_lat = _extract_setback_from_text(text, "lateral")
        back, c_back, ex_back = _extract_setback_from_text(text, "back")
        green, c_green, ex_green = _extract_green_ratio(text)
        parking, c_parking, ex_parking = _extract_parking_ratio(text)

        return {
            "recul_voie_m": road,
            "recul_voie_conf": c_road,
            "recul_limite_m": lat,
            "recul_limite_conf": c_lat,
            "recul_fond_m": back,
            "recul_fond_conf": c_back,
            "pleine_terre_pct": green,
            "pleine_terre_conf": c_green,
            "stationnement_par_logement": parking,
            "stationnement_conf": c_parking,
            "contraintes_extrait": " | ".join(
                x for x in [ex_road, ex_lat, ex_back, ex_green, ex_parking] if x
            )[:1500],
        }
    except Exception:
        return {}


def enrich_zone_rule_index_with_constraints(zone_rule_index, df):
    if df is None or df.empty:
        return zone_rule_index

    for rec in df[["zone_plu", "reglement_url"]].fillna("").drop_duplicates().to_dict("records"):
        key = (rec["zone_plu"], rec["reglement_url"])
        base = dict(zone_rule_index.get(key, {}))
        extra = _zone_extra_written_constraints(rec["reglement_url"], rec["zone_plu"])
        base.update(extra)
        zone_rule_index[key] = base
    return zone_rule_index


def _prescription_extra_value(rule):
    """
    Extrait les attributs métier des prescriptions CNIG supplémentaires :
    15 = implantation/recul
    18 = OAP
    42 = coefficient de biotope
    02 = limitation/interdiction de constructibilité
    05 = emplacement réservé
    """
    props = rule.get("props", {}) or {}
    optionals = _parse_optional_psc_attrs(props)
    raw = " | ".join(
        [rule.get("libelle", ""), rule.get("txt", "")]
        + [f"{a}={v}" for a, v in optionals]
    )
    typepsc = rule.get("typepsc")
    stypepsc = rule.get("stypepsc")

    result = {"raw": raw}
    if typepsc == "15":
        dist = None
        for a, v in optionals:
            na = _rule_norm(a)
            if "valeur de recul" in na or "recul" in na:
                dist = _parse_distance_m(v)
                if dist is not None:
                    break
        if dist is None:
            dist = _parse_distance_m(raw)
        result["recul_m"] = dist
        result["recul_kind"] = {
            "01": "voie",
            "02": "lateral",
            "03": "fond",
        }.get(stypepsc, "autre")

    if typepsc == "42":
        pct = None
        for a, v in optionals:
            if "biotope" in _rule_norm(a) and "min" in _rule_norm(a):
                pct = _parse_percent_value(v)
                if pct is not None:
                    break
        if pct is None:
            pct = _parse_percent_value(raw)
        result["biotope_min_pct"] = pct

    if typepsc == "18":
        result["oap"] = True
    if typepsc == "02":
        result["constructibilite_limitee"] = True
        result["interdiction"] = stypepsc == "01"
    if typepsc == "05":
        result["emplacement_reserve"] = True
    return result


def build_extended_constraint_index(prescriptions):
    rules = []
    for layer_name, fc in (prescriptions or {}).items():
        for feat in fc.get("features", []) or []:
            props = feat.get("properties", {}) or {}
            typepsc = _clean_psc_code(props.get("typepsc"))
            stypepsc = _clean_psc_code(props.get("stypepsc"))
            if typepsc not in {"02", "05", "15", "18", "42"}:
                continue
            geom = valid_shape(feat)
            if geom is None:
                continue
            base = {
                "geom": geom,
                "layer": layer_name,
                "typepsc": typepsc,
                "stypepsc": stypepsc,
                "libelle": str(props.get("libelle") or ""),
                "txt": str(props.get("txt") or ""),
                "urlfic": str(props.get("urlfic") or ""),
                "props": props,
            }
            base.update(_prescription_extra_value(base))
            rules.append(base)

    geoms = [r["geom"] for r in rules]
    return {
        "rules": rules,
        "tree": STRtree(geoms) if geoms else None,
    }


def _local_projected_area_and_buffer(geom, setback_m):
    """
    Buffer métrique local en UTM. Utilisé uniquement pour produire une enveloppe
    prudente lorsque les reculs ont été extraits.
    """
    if setback_m is None or setback_m <= 0:
        return geodesic_area_m2(geom), geom

    c = geom.representative_point()
    zone = int((c.x + 180) // 6) + 1
    epsg = 32600 + zone if c.y >= 0 else 32700 + zone
    to_local = Transformer.from_crs(4326, epsg, always_xy=True)
    to_wgs = Transformer.from_crs(epsg, 4326, always_xy=True)

    local = shapely_transform(to_local.transform, geom)
    inner = local.buffer(-float(setback_m))
    if inner.is_empty:
        return 0.0, None
    return float(inner.area), shapely_transform(to_wgs.transform, inner)


def apply_extended_constraints_to_results(df, feature_map, constraint_index, zone_rule_index, apply_setbacks=True):
    df = df.copy()

    defaults = {
        "recul_voie_m": None,
        "recul_limite_m": None,
        "recul_fond_m": None,
        "pleine_terre_pct": None,
        "biotope_min_pct": None,
        "stationnement_par_logement": None,
        "oap": False,
        "emplacement_reserve": False,
        "constructibilite_limitee": False,
        "interdiction_constructibilite": False,
        "surface_enveloppe_prudente_m2": None,
        "emprise_effective_pct": None,
        "surface_brute_corrigee_m2": None,
        "sdp_corrigee_m2": None,
        "shab_corrigee_m2": None,
        "logements_corriges": None,
        "contraintes_plu": "",
    }
    for c, v in defaults.items():
        if c not in df.columns:
            df[c] = v

    tree = constraint_index.get("tree")
    rules = constraint_index.get("rules", [])

    for idx, row in df.iterrows():
        zone_key = (str(row.get("zone_plu") or ""), str(row.get("reglement_url") or ""))
        zr = zone_rule_index.get(zone_key, {})

        # Règles écrites de base.
        df.at[idx, "recul_voie_m"] = zr.get("recul_voie_m")
        df.at[idx, "recul_limite_m"] = zr.get("recul_limite_m")
        df.at[idx, "recul_fond_m"] = zr.get("recul_fond_m")
        df.at[idx, "pleine_terre_pct"] = zr.get("pleine_terre_pct")
        df.at[idx, "stationnement_par_logement"] = zr.get("stationnement_par_logement")

        feat = feature_map.get(row["reference"])
        pg = valid_shape(feat) if feat else None
        if pg is None:
            continue

        hits = []
        if tree is not None:
            for ridx in tree.query(pg, predicate="intersects"):
                r = rules[int(ridx)]
                try:
                    inter = pg.intersection(r["geom"])
                    ratio = geodesic_area_m2(inter) / max(geodesic_area_m2(pg), 0.01)
                except Exception:
                    ratio = 0.0
                if ratio <= 0:
                    continue
                hits.append((ratio, r))

        # Prescriptions graphiques locales : priorité sur le texte de zone.
        notes = []
        biotope_vals = []
        for ratio, r in hits:
            t, st = r["typepsc"], r["stypepsc"]

            if t == "15" and r.get("recul_m") is not None:
                kind = r.get("recul_kind")
                if kind == "voie":
                    df.at[idx, "recul_voie_m"] = r["recul_m"]
                elif kind == "lateral":
                    df.at[idx, "recul_limite_m"] = r["recul_m"]
                elif kind == "fond":
                    df.at[idx, "recul_fond_m"] = r["recul_m"]
                notes.append(f"recul graphique {kind} {r['recul_m']} m")

            if t == "42" and r.get("biotope_min_pct") is not None:
                biotope_vals.append(float(r["biotope_min_pct"]))
                notes.append(f"biotope {r['biotope_min_pct']}%")

            if t == "18":
                df.at[idx, "oap"] = True
                notes.append("OAP")

            if t == "05":
                df.at[idx, "emplacement_reserve"] = True
                notes.append("emplacement réservé")

            if t == "02":
                df.at[idx, "constructibilite_limitee"] = True
                if r.get("interdiction") and ratio >= 0.50:
                    df.at[idx, "interdiction_constructibilite"] = True
                notes.append(f"constructibilité limitée {round(ratio*100)}%")

        if biotope_vals:
            df.at[idx, "biotope_min_pct"] = max(biotope_vals)

        # Enveloppe prudente : faute de distinguer précisément la façade sur rue
        # avec le seul cadastre, on applique le plus grand recul à tout le contour.
        # C'est volontairement conservateur et clairement identifié comme tel.
        setbacks = [
            x for x in [
                df.at[idx, "recul_voie_m"],
                df.at[idx, "recul_limite_m"],
                df.at[idx, "recul_fond_m"],
            ]
            if x not in [None, ""]
        ]
        envelope_area = float(row["surface_m2"])
        if apply_setbacks and setbacks:
            envelope_area, _ = _local_projected_area_and_buffer(pg, max(float(x) for x in setbacks))
        df.at[idx, "surface_enveloppe_prudente_m2"] = round(envelope_area)

        base_emprise = row.get("emprise_plu_pct")
        if base_emprise not in [None, ""]:
            effective = float(base_emprise)
            green_caps = []
            if df.at[idx, "pleine_terre_pct"] not in [None, ""]:
                green_caps.append(100.0 - float(df.at[idx, "pleine_terre_pct"]))
            if df.at[idx, "biotope_min_pct"] not in [None, ""]:
                green_caps.append(100.0 - float(df.at[idx, "biotope_min_pct"]))
            if green_caps:
                effective = min(effective, min(green_caps))
            df.at[idx, "emprise_effective_pct"] = round(max(0.0, effective), 2)

            levels = row.get("niveaux_plu")
            if levels not in [None, ""] and envelope_area > 0:
                footprint_by_ratio = float(row["surface_m2"]) * effective / 100.0
                footprint = min(footprint_by_ratio, envelope_area)
                gross = footprint * float(levels)
                sdp_ratio = float(row.get("ratio_sdp_pct") or 80.0) / 100.0
                shab_ratio = float(row.get("ratio_shab_pct") or 80.0) / 100.0
                shab_per_dwelling = float(row.get("shab_par_logement") or 55.0)

                sdp = gross * sdp_ratio
                shab = sdp * shab_ratio
                logements = math.floor(shab / max(shab_per_dwelling, 1.0))

                df.at[idx, "surface_brute_corrigee_m2"] = round(gross)
                df.at[idx, "sdp_corrigee_m2"] = round(sdp)
                df.at[idx, "shab_corrigee_m2"] = round(shab)
                df.at[idx, "logements_corriges"] = int(max(0, logements))

        if df.at[idx, "interdiction_constructibilite"]:
            df.at[idx, "logements_corriges"] = 0

        df.at[idx, "contraintes_plu"] = " | ".join(dict.fromkeys(notes))

    return df


def build_preinterpreted_rule_base(df):
    if df is None or df.empty:
        return pd.DataFrame()

    cols = [
        "zone_plu",
        "zone_description",
        "emprise_plu_pct",
        "niveaux_plu",
        "hauteur_plu_m",
        "recul_voie_m",
        "recul_limite_m",
        "recul_fond_m",
        "pleine_terre_pct",
        "stationnement_par_logement",
        "regle_zone_emprise",
        "regle_zone_hauteur",
        "reglement_url",
    ]
    existing = [c for c in cols if c in df.columns]
    base = df[existing].drop_duplicates().copy()
    return base.sort_values(["zone_plu"]).reset_index(drop=True)

def _clean_psc_code(value):
    txt = re.sub(r"\D", "", str(value or ""))
    if not txt:
        return ""
    return txt.zfill(2)[-2:]


def _parse_optional_psc_attrs(props):
    """
    Les attributs supplémentaires CNIG sont stockés par paires
    LIB_ATTRn / LIB_VALn lorsqu'ils sont exposés par la source.
    """
    lower = {str(k).lower(): v for k, v in (props or {}).items()}
    pairs = []
    for i in range(1, 21):
        a = lower.get(f"lib_attr{i}")
        v = lower.get(f"lib_val{i}")
        if a not in [None, ""] or v not in [None, ""]:
            pairs.append((str(a or ""), str(v or "")))
    return pairs


def _parse_percent_value(text):
    txt = normalize_text(text)
    m = re.search(r"(?<!\d)(\d{1,3}(?:[.,]\d+)?)\s*%", txt)
    if m:
        val = float(m.group(1).replace(",", "."))
        if 0 < val <= 100:
            return val

    # coefficient décimal : 0,4 => 40 %
    m = re.search(r"(?:ces|coef(?:ficient)?[^0-9]{0,20})(0[.,]\d+)", txt)
    if m:
        val = float(m.group(1).replace(",", ".")) * 100.0
        if 0 < val <= 100:
            return val
    return None


def _parse_rplus_levels(text):
    txt = normalize_text(text)
    m = re.search(r"\br\s*\+\s*(\d{1,2})\b", txt)
    if m:
        return int(m.group(1)) + 1
    if re.search(r"\brdc\b|\brez[- ]de[- ]chaussee\b", txt):
        # Ne conclure à R seul que si le texte parle explicitement de hauteur/niveaux.
        if "niveau" in txt or "hauteur" in txt:
            return 1
    m = re.search(r"(?<!\d)(\d{1,2})\s+niveaux?\b", txt)
    if m:
        n = int(m.group(1))
        if 1 <= n <= 20:
            return n
    return None


def _parse_height_m(text):
    txt = normalize_text(text)
    # Chercher une valeur liée à la hauteur, pas n'importe quelle cote du libellé.
    patterns = [
        r"hauteur.{0,80}?(\d{1,2}(?:[.,]\d+)?)\s*m\b",
        r"(\d{1,2}(?:[.,]\d+)?)\s*m\b.{0,60}?hauteur",
        r"h\s*[=:]\s*(\d{1,2}(?:[.,]\d+)?)\s*m\b",
    ]
    for pat in patterns:
        m = re.search(pat, txt)
        if m:
            val = float(m.group(1).replace(",", "."))
            if 2 <= val <= 80:
                return val
    return None


def _rule_from_prescription(feature, layer_name):
    props = feature.get("properties", {}) or {}
    typepsc = _clean_psc_code(props.get("typepsc"))
    stypepsc = _clean_psc_code(props.get("stypepsc"))

    if typepsc not in {"38", "39"}:
        return None

    geom = valid_shape(feature)
    if geom is None:
        return None

    libelle = str(props.get("libelle") or "")
    txt = str(props.get("txt") or "")
    nature = str(props.get("nature") or "")
    urlfic = str(props.get("urlfic") or "")
    base_text = " | ".join(x for x in [libelle, txt, nature] if x)

    optionals = _parse_optional_psc_attrs(props)
    optional_text = " | ".join(f"{a}={v}" for a, v in optionals)
    all_text = " | ".join(x for x in [base_text, optional_text] if x)

    rule = {
        "geom": geom,
        "layer": layer_name,
        "typepsc": typepsc,
        "stypepsc": stypepsc,
        "libelle": libelle,
        "txt": txt,
        "urlfic": urlfic,
        "raw": all_text,
        "emprise_pct": None,
        "height_m": None,
        "levels": None,
        "confidence": 0,
        "method": "",
        "excerpt": "",
        "props": props,
    }

    if typepsc == "38":
        # 38-02 = emprise maximale. 38-00 générique accepté si valeur explicite.
        if stypepsc not in {"00", "02"}:
            return None

        for attr, val in optionals:
            na = normalize_text(attr).replace(" ", "_")
            if "empris" in na and ("max" in na or stypepsc == "02"):
                pct = _parse_percent_value(val)
                if pct is not None:
                    rule["emprise_pct"] = pct
                    rule["confidence"] = 100
                    rule["method"] = f"Attribut graphique {attr}"
                    return rule

        pct = _parse_percent_value(all_text)
        if pct is not None:
            rule["emprise_pct"] = pct
            rule["confidence"] = 88 if stypepsc == "02" else 75
            rule["method"] = "Étiquette/libellé prescription graphique"
            return rule
        return rule

    # type 39 : hauteur.
    # On privilégie 39-02 (max). 50/51 sont tolérés pour documents qui les publient.
    if stypepsc not in {"00", "02", "50", "51"}:
        return None

    for attr, val in optionals:
        na = normalize_text(attr).replace(" ", "_")
        if "hauteur" in na and ("rpl" in na or "etage" in na):
            levels = _parse_rplus_levels(val)
            if levels is not None:
                rule["levels"] = levels
                rule["confidence"] = 100
                rule["method"] = f"Attribut graphique {attr}"
        if "hauteur" in na and ("metre" in na or "meter" in na or "max" in na):
            h = _parse_height_m(f"hauteur {val}")
            if h is not None:
                rule["height_m"] = h
                rule["confidence"] = max(rule["confidence"], 100)
                rule["method"] = f"Attribut graphique {attr}"

    if rule["levels"] is None:
        levels = _parse_rplus_levels(all_text)
        if levels is not None:
            rule["levels"] = levels
            rule["confidence"] = max(rule["confidence"], 90 if stypepsc == "02" else 80)
            rule["method"] = "Étiquette/libellé prescription graphique"

    if rule["height_m"] is None:
        h = _parse_height_m(all_text)
        if h is not None:
            rule["height_m"] = h
            rule["confidence"] = max(rule["confidence"], 88 if stypepsc == "02" else 75)
            if not rule["method"]:
                rule["method"] = "Étiquette/libellé prescription graphique"

    return rule


def build_graphic_rule_indexes(prescriptions, floor_height_m):
    """
    Prépare un index spatial. Les règles surfaciques sont la source principale
    pour un calcul par parcelle. Les règles linéaires/ponctuelles de hauteur
    sont conservées comme indications supplémentaires.
    """
    surf_rules = []
    other_height_rules = []

    for layer_name, fc in (prescriptions or {}).items():
        for feat in fc.get("features", []) or []:
            rule = _rule_from_prescription(feat, layer_name)
            if not rule:
                continue
            rule = _complete_graphic_rule(rule, float(floor_height_m))
            if layer_name == "prescription-surf":
                surf_rules.append(rule)
            elif rule["typepsc"] == "39":
                other_height_rules.append(rule)

    surf_geoms = [r["geom"] for r in surf_rules]
    other_geoms = [r["geom"] for r in other_height_rules]

    return {
        "surf_rules": surf_rules,
        "surf_tree": STRtree(surf_geoms) if surf_geoms else None,
        "other_height_rules": other_height_rules,
        "other_tree": STRtree(other_geoms) if other_geoms else None,
    }


def _select_surface_rule(parcel_geom, tree, rules, kind):
    if tree is None:
        return None

    idxs = tree.query(parcel_geom, predicate="intersects")
    candidates = []
    parcel_area = max(geodesic_area_m2(parcel_geom), 0.01)

    for idx in idxs:
        rule = rules[int(idx)]
        if kind == "emprise" and rule.get("emprise_pct") is None:
            continue
        if kind == "height" and rule.get("levels") is None and rule.get("height_m") is None:
            continue

        try:
            inter = parcel_geom.intersection(rule["geom"])
            overlap_area = geodesic_area_m2(inter)
        except Exception:
            overlap_area = 0.0

        if overlap_area <= 0:
            continue

        ratio = overlap_area / parcel_area
        candidates.append((ratio, rule))

    if not candidates:
        return None

    # Priorité à la prescription couvrant la plus grande part de la parcelle.
    candidates.sort(key=lambda x: (-x[0], -x[1].get("confidence", 0)))
    ratio, best = candidates[0]

    # Une prescription graphique sur une toute petite portion ne doit pas être
    # appliquée à toute la parcelle. On la signale seulement si < 50 %.
    result = dict(best)
    result["coverage_ratio"] = ratio
    result["partial"] = ratio < 0.50
    return result


def _select_line_point_height(parcel_geom, tree, rules):
    if tree is None:
        return None
    idxs = tree.query(parcel_geom, predicate="intersects")
    candidates = []
    for idx in idxs:
        rule = rules[int(idx)]
        if rule.get("levels") is None and rule.get("height_m") is None:
            continue
        candidates.append(rule)
    if not candidates:
        return None

    # En cas de plusieurs plafonds, prendre le plus restrictif.
    def key(r):
        if r.get("levels") is not None:
            return (0, r["levels"])
        return (1, r.get("height_m") or 999)
    return sorted(candidates, key=key)[0]


def enrich_with_graphic_plu_rules(
    df,
    feature_map,
    rule_indexes,
    zone_rule_index,
    ratio_sdp_pct,
    ratio_shab_pct,
    shab_par_logement,
    floor_height_m,
):
    """
    Moteur hybride :
    1. règle de base extraite du règlement écrit de la zone ;
    2. prescriptions graphiques 38/39 qui remplacent localement la règle de base ;
    3. découpage géométrique réel de la parcelle en sous-surfaces homogènes ;
    4. somme des surfaces brutes de chaque sous-surface.

    On ne généralise donc plus une prescription locale à toute la parcelle.
    """
    df = df.copy()

    cols = {
        "emprise_plu_pct": None,
        "emprise_graphique_couverture_pct": None,
        "hauteur_plu_m": None,
        "niveaux_plu": None,
        "surface_brute_plu_m2": None,
        "sdp_plu_m2": None,
        "shab_plu_m2": None,
        "logements_plu": None,
        "gabarit_plu_statut": "À vérifier",
        "gabarit_plu_source": "",
        "gabarit_plu_confiance": 0,
        "gabarit_couverture_pct": 0.0,
        "prescription_emprise": "",
        "prescription_hauteur": "",
        "prescription_url": "",
        "regle_zone_emprise": "",
        "regle_zone_hauteur": "",
        "regle_zone_extrait": "",
    }
    for c, default in cols.items():
        if c not in df.columns:
            df[c] = default

    if df.empty:
        return df

    surf_tree = rule_indexes.get("surf_tree")
    surf_rules = rule_indexes.get("surf_rules", [])

    # Les prescriptions linéaires/ponctuelles restent informatives :
    # elles ne définissent pas seules une surface homogène de calcul.
    other_tree = rule_indexes.get("other_tree")
    other_rules = rule_indexes.get("other_height_rules", [])

    for idx, row in df.iterrows():
        feat = feature_map.get(row["reference"])
        if not feat:
            continue
        pg = valid_shape(feat)
        if pg is None:
            continue

        zone_key = (
            str(row.get("zone_plu") or ""),
            str(row.get("reglement_url") or ""),
        )
        base = zone_rule_index.get(zone_key, {})

        base_e = base.get("emprise_pct")
        base_l = base.get("levels")
        base_h = base.get("height_m")

        if base_l is None and base_h is not None:
            base_l = max(
                1,
                int(math.floor(float(base_h) / max(float(floor_height_m), 0.1))),
            )

        pieces = [
            {
                "geom": pg,
                "emprise": base_e,
                "levels": base_l,
                "emprise_source": "Règlement écrit de zone" if base_e is not None else "",
                "levels_source": "Règlement écrit de zone" if base_l is not None else "",
                "emprise_conf": base.get("emprise_conf", 0),
                "levels_conf": base.get("height_conf", 0),
            }
        ]

        applied_emprise = []
        applied_height = []

        # Prescriptions surfaciques qui intersectent la parcelle.
        if surf_tree is not None:
            idxs = surf_tree.query(pg, predicate="intersects")
            # Appliquer d'abord les prescriptions les plus couvrantes.
            rule_hits = []
            for ridx in idxs:
                rule = surf_rules[int(ridx)]
                try:
                    inter = pg.intersection(rule["geom"])
                    ar = geodesic_area_m2(inter)
                except Exception:
                    ar = 0
                if ar > 0.05:
                    rule_hits.append((ar, rule))
            rule_hits.sort(key=lambda x: -x[0])

            for _, rule in rule_hits:
                if rule.get("emprise_pct") is not None:
                    pieces = _overlay_piece(
                        pieces,
                        rule["geom"],
                        "emprise",
                        float(rule["emprise_pct"]),
                        f"Graphique {rule.get('typepsc')}-{rule.get('stypepsc')}",
                        rule.get("confidence", 0),
                    )
                    applied_emprise.append(rule)

                levels = rule.get("levels")
                height = rule.get("height_m")
                if levels is None and height is not None:
                    levels = max(
                        1,
                        int(math.floor(float(height) / max(float(floor_height_m), 0.1))),
                    )
                if levels is not None and rule.get("typepsc") == "39":
                    pieces = _overlay_piece(
                        pieces,
                        rule["geom"],
                        "levels",
                        int(levels),
                        f"Graphique {rule.get('typepsc')}-{rule.get('stypepsc')}",
                        rule.get("confidence", 0),
                    )
                    applied_height.append(rule)

        parcel_area = max(geodesic_area_m2(pg), 0.01)
        known_area = 0.0
        gross = 0.0
        weighted_e = 0.0
        weighted_l = 0.0
        confs = []

        for piece in pieces:
            area = geodesic_area_m2(piece["geom"])
            e = piece.get("emprise")
            levels = piece.get("levels")

            if e is not None:
                weighted_e += area * float(e)
            if levels is not None:
                weighted_l += area * float(levels)

            if e is not None and levels is not None:
                known_area += area
                gross += area * (float(e) / 100.0) * float(levels)
                if piece.get("emprise_conf"):
                    confs.append(piece["emprise_conf"])
                if piece.get("levels_conf"):
                    confs.append(piece["levels_conf"])

        coverage = max(0.0, min(100.0, 100.0 * known_area / parcel_area))
        df.at[idx, "gabarit_couverture_pct"] = round(coverage, 1)

        if weighted_e > 0:
            df.at[idx, "emprise_plu_pct"] = round(weighted_e / parcel_area, 2)
        if weighted_l > 0:
            df.at[idx, "niveaux_plu"] = round(weighted_l / parcel_area, 2)
        if base_h is not None:
            df.at[idx, "hauteur_plu_m"] = base_h

        df.at[idx, "regle_zone_emprise"] = base.get("emprise_method", "")
        df.at[idx, "regle_zone_hauteur"] = base.get("height_method", "")
        df.at[idx, "regle_zone_extrait"] = " | ".join(
            x
            for x in [
                base.get("emprise_excerpt", ""),
                base.get("height_excerpt", ""),
            ]
            if x
        )[:1200]

        if applied_emprise:
            df.at[idx, "prescription_emprise"] = " / ".join(
                dict.fromkeys(
                    f"{r.get('typepsc')}-{r.get('stypepsc')} {r.get('libelle') or r.get('txt') or ''}".strip()
                    for r in applied_emprise
                )
            )[:700]
            url = next((r.get("urlfic") for r in applied_emprise if r.get("urlfic")), "")
            if url:
                df.at[idx, "prescription_url"] = url

        if applied_height:
            df.at[idx, "prescription_hauteur"] = " / ".join(
                dict.fromkeys(
                    f"{r.get('typepsc')}-{r.get('stypepsc')} {r.get('libelle') or r.get('txt') or ''}".strip()
                    for r in applied_height
                )
            )[:700]
            if not df.at[idx, "prescription_url"]:
                url = next((r.get("urlfic") for r in applied_height if r.get("urlfic")), "")
                if url:
                    df.at[idx, "prescription_url"] = url

        # Calcul fiable seulement lorsque pratiquement toute la parcelle dispose
        # simultanément d'une règle d'emprise et de hauteur.
        if coverage >= 98.0 and gross > 0:
            sdp = gross * float(ratio_sdp_pct) / 100.0
            shab = sdp * float(ratio_shab_pct) / 100.0
            logements = math.floor(shab / max(float(shab_par_logement), 1.0))

            df.at[idx, "surface_brute_plu_m2"] = round(gross)
            df.at[idx, "sdp_plu_m2"] = round(sdp)
            df.at[idx, "shab_plu_m2"] = round(shab)
            df.at[idx, "logements_plu"] = int(max(0, logements))

            if applied_emprise or applied_height:
                source = "Règlement écrit + prescriptions graphiques"
            else:
                source = "Règlement écrit de zone"

            df.at[idx, "gabarit_plu_source"] = source
            df.at[idx, "gabarit_plu_statut"] = "Calculé — couverture complète"
            df.at[idx, "gabarit_plu_confiance"] = int(min(confs)) if confs else 60
        else:
            missing = []
            if coverage < 98:
                missing.append(f"règles complètes sur {coverage:.1f}%")
            if base_e is None and not applied_emprise:
                missing.append("emprise non trouvée")
            if base_l is None and not applied_height:
                missing.append("hauteur/niveaux non trouvés")

            df.at[idx, "gabarit_plu_source"] = (
                "Règlement + graphique incomplet"
                if (base_e is not None or base_l is not None or applied_emprise or applied_height)
                else "Aucune règle chiffrée fiable"
            )
            df.at[idx, "gabarit_plu_statut"] = "À vérifier — " + ", ".join(missing)

        # Les prescriptions linéaires/ponctuelles de hauteur restent visibles comme alerte.
        if other_tree is not None:
            hits = other_tree.query(pg, predicate="intersects")
            hints = []
            for hid in hits:
                r = other_rules[int(hid)]
                if r.get("levels") is not None or r.get("height_m") is not None:
                    hints.append(
                        f"{r.get('typepsc')}-{r.get('stypepsc')} {r.get('libelle') or r.get('txt') or ''}".strip()
                    )
            if hints and not df.at[idx, "prescription_hauteur"]:
                df.at[idx, "prescription_hauteur"] = "Indication linéaire/ponctuelle : " + " / ".join(
                    dict.fromkeys(hints)
                )[:600]

    return df


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
st.title("🏗️ Prospecteur Foncier — V3.7 — Préfaisabilité PLU")
st.caption(
    "Cadastre + archive CNIG complète + règles pré-interprétées + contraintes par parcelle."
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
    )
with r2:
    ratio_shab_pct = st.number_input(
        "SHAB / SDP (%)",
        min_value=1.0,
        max_value=100.0,
        value=80.0,
        step=1.0,
    )
with r3:
    shab_par_logement = st.number_input(
        "SHAB moyenne / logement (m²)",
        min_value=10.0,
        max_value=250.0,
        value=55.0,
        step=1.0,
    )
with r4:
    floor_height_m = st.number_input(
        "Hauteur / niveau si PLU en mètres",
        min_value=2.4,
        max_value=5.0,
        value=3.0,
        step=0.1,
        help="Utilisé uniquement si la prescription graphique indique une hauteur en mètres sans R+N.",
    )

st.caption(
    "Calcul graphique : surface brute = surface cadastrale × emprise maximale graphique × niveaux graphiques ; "
    f"SDP × {ratio_sdp_pct:.0f} % ; SHAB × {ratio_shab_pct:.0f} % ; "
    f"{shab_par_logement:.0f} m² SHAB/logement."
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
    "Méthode V3.7 : le logiciel télécharge aussi l'archive CNIG complète du PLU/PLUi afin de lire "
    "les attributs supplémentaires publiés par la collectivité (LIB_ATTR/LIB_VAL). "
    "Il combine ensuite zonage, emprise, hauteur, reculs, biotope/pleine terre, OAP et autres prescriptions."
)

x1, x2 = st.columns(2)
with x1:
    analyse_cnig_complete = st.checkbox(
        "Analyse renforcée de l'archive CNIG complète",
        value=True,
        help=(
            "Permet notamment de récupérer les attributs complémentaires des prescriptions "
            "qui ne remontent pas toujours dans l'API Carto."
        ),
    )
with x2:
    appliquer_reculs_prudents = st.checkbox(
        "Appliquer les reculs au calcul (mode prudent)",
        value=True,
        help=(
            "En l'absence d'identification parfaite de chaque façade de parcelle, "
            "le plus grand recul trouvé est appliqué à tout le contour. "
            "Le résultat est donc volontairement conservateur."
        ),
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
    "La V3.7 intègre désormais plusieurs contraintes supplémentaires, mais le résultat reste une préfaisabilité. "
    "Les règles atypiques, les risques et certaines servitudes peuvent nécessiter une vérification humaine."
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

        with st.spinner("Chargement des prescriptions graphiques PLU…"):
            api_graphic_prescriptions = fetch_gpu_graphic_prescriptions(partition, insee=insee)

            archive_bundle = None
            if analyse_cnig_complete:
                try:
                    with st.spinner("Téléchargement et lecture de l'archive CNIG complète…"):
                        archive_bundle = fetch_cnig_archive_layers(partition)
                    st.session_state["archive_error"] = ""
                except Exception as archive_exc:
                    archive_bundle = None
                    st.session_state["archive_error"] = str(archive_exc)
            else:
                st.session_state["archive_error"] = ""

            graphic_prescriptions = merge_graphic_sources(
                api_graphic_prescriptions,
                archive_bundle,
            )
            graphic_rule_indexes = build_graphic_rule_indexes(
                graphic_prescriptions,
                floor_height_m=floor_height_m,
            )
            extended_constraint_index = build_extended_constraint_index(
                graphic_prescriptions
            )

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

        with st.spinner("Lecture ciblée des règles de zone + croisement graphique 38/39…"):
            zone_rule_index = build_zone_written_rule_index(
                results,
                floor_height_m=floor_height_m,
            )
            zone_rule_index = enrich_zone_rule_index_with_constraints(
                zone_rule_index,
                results,
            )
            results = enrich_with_graphic_plu_rules(
                results,
                feature_map=feature_map,
                rule_indexes=graphic_rule_indexes,
                zone_rule_index=zone_rule_index,
                ratio_sdp_pct=ratio_sdp_pct,
                ratio_shab_pct=ratio_shab_pct,
                shab_par_logement=shab_par_logement,
                floor_height_m=floor_height_m,
            )

            results = apply_extended_constraints_to_results(
                results,
                feature_map=feature_map,
                constraint_index=extended_constraint_index,
                zone_rule_index=zone_rule_index,
                apply_setbacks=appliquer_reculs_prudents,
            )
            st.session_state["rule_base"] = build_preinterpreted_rule_base(results)
            st.session_state["archive_info"] = {
                "archive_size_mb": (archive_bundle or {}).get("archive_size_mb"),
                "archive_url": (archive_bundle or {}).get("archive_url"),
                "xml_rules_count": len((archive_bundle or {}).get("xml_rules", [])),
            }

            if feature_map and not results.empty:
                rule_by_ref = results.set_index("reference").to_dict("index")
                for ref, feat in feature_map.items():
                    row = rule_by_ref.get(ref)
                    if not row:
                        continue
                    props = feat.setdefault("properties", {})
                    props["emprise_plu"] = row.get("emprise_plu_pct")
                    props["niveaux_plu"] = row.get("niveaux_plu")
                    props["hauteur_plu"] = row.get("hauteur_plu_m")
                    props["surface_brute_plu"] = row.get("surface_brute_plu_m2")
                    props["sdp_plu"] = row.get("sdp_plu_m2")
                    props["shab_plu"] = row.get("shab_plu_m2")
                    props["logements_plu"] = row.get("logements_plu")
                    props["statut_gabarit"] = row.get("gabarit_plu_statut")
                    props["source_gabarit"] = row.get("gabarit_plu_source")
                    props["couverture_gabarit"] = row.get("gabarit_couverture_pct")
                    props["recul_voie"] = row.get("recul_voie_m")
                    props["recul_limite"] = row.get("recul_limite_m")
                    props["pleine_terre"] = row.get("pleine_terre_pct")
                    props["logements_corriges"] = row.get("logements_corriges")
                    props["oap"] = "Oui" if row.get("oap") else "Non"

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
    results["logements_plu"] = pd.to_numeric(results["logements_plu"], errors="coerce")
    results["logements_corriges"] = pd.to_numeric(
        results.get("logements_corriges"),
        errors="coerce",
    )

    # Priorité au calcul corrigé par contraintes ; fallback au gabarit emprise/hauteur.
    results["logements_retenus"] = results["logements_corriges"].combine_first(
        results["logements_plu"]
    )

    calculable = results["logements_retenus"].notna()

    filtered = results[
        calculable
        & (results["logements_retenus"] >= min_log)
        & (results["logements_retenus"] <= max_log)
        & (results["interdiction_constructibilite"] != True)
    ].copy()

    unresolved = results[~calculable].copy()

    cleaned = []
    for frame in [filtered, unresolved]:
        if "economic_vocation" in frame.columns:
            frame = frame[frame["economic_vocation"] == False].copy()
        if "collectif_existant" in frame.columns:
            frame = frame[frame["collectif_existant"] == False].copy()

        if terrain_mode == "Terrain nu":
            frame = frame[frame["terrain_bati"] == False]
        elif terrain_mode == "Terrain bâti":
            frame = frame[frame["terrain_bati"] == True]

        frame = frame.sort_values(
            ["score", "surface_m2"],
            ascending=[False, False],
        ).reset_index(drop=True)
        cleaned.append(frame)

    filtered, unresolved = cleaned

    st.subheader("2. Résultats sur le cadastre réel")

    if filtered.empty and not unresolved.empty:
        st.warning(
            "Des parcelles candidates ont été trouvées, mais le moteur hybride n'a pas obtenu "
            "une couverture suffisante en emprise + hauteur pour calculer leur capacité. "
            "Elles restent visibles dans « Gabarit PLU à vérifier »."
        )
    elif filtered.empty and unresolved.empty:
        st.warning("Aucune parcelle ne correspond aux critères actuels.")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Parcelles dans la cible logements corrigée", len(filtered))
    m2.metric("Gabarit à vérifier", len(unresolved))
    m3.metric("Parcelles candidates analysées", len(results))
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

    if st.session_state.get("archive_error"):
        st.warning(
            "L'archive CNIG complète n'a pas pu être exploitée. Le logiciel continue avec l'API Carto "
            "et le règlement écrit. Détail : " + st.session_state["archive_error"]
        )

    archive_info = st.session_state.get("archive_info", {})
    if archive_info.get("archive_size_mb"):
        st.caption(
            f"Archive CNIG analysée : {archive_info['archive_size_mb']} Mo — "
            f"{archive_info.get('xml_rules_count', 0)} règlement(s) XML structuré(s) détecté(s)."
        )

    rule_base = st.session_state.get("rule_base")
    if isinstance(rule_base, pd.DataFrame) and not rule_base.empty:
        with st.expander("Base PLU pré-interprétée par zone"):
            st.dataframe(rule_base, use_container_width=True, hide_index=True)
            st.download_button(
                "Télécharger la base de règles en CSV",
                data=rule_base.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"base_regles_PLU_{commune_name}_{insee}.csv",
                mime="text/csv",
            )

    map_source = filtered if not filtered.empty else unresolved

    if map_source.empty:
        st.warning("Aucune parcelle ne correspond aux critères actuels.")
    else:
        fmap = st.session_state.get("feature_map", {})
        map_features = []
        for ref in map_source["reference"].head(2000):
            feat = fmap.get(ref)
            if feat:
                map_features.append(feat)

        if map_features:
            fc = {"type": "FeatureCollection", "features": map_features}
            center_lat = float(map_source["latitude"].mean())
            center_lon = float(map_source["longitude"].mean())

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
                        "Emprise PLU : {emprise_plu} %<br/>"
                        "Niveaux PLU : {niveaux_plu}<br/>"
                        "Hauteur PLU : {hauteur_plu} m<br/>"
                        "Surface brute PLU : {surface_brute_plu} m²<br/>"
                        "SDP PLU : {sdp_plu} m²<br/>"
                        "SHAB PLU : {shab_plu} m²<br/>"
                        "Habitat : {habitat}<br/>"
                        "Confiance habitat : {confiance_habitat} %<br/>"
                        "Collectif existant : {collectif}<br/>"
                        "Propriétaire PM : {proprietaire}<br/>"
                        "Type propriétaire : {type_proprietaire}<br/>"
                        "Zone : {typezone} {zone}<br/>"
                        "Potentiel PLU : {logements_plu} logements<br/>"
                        "Gabarit : {statut_gabarit}<br/>"
                        "Source : {source_gabarit}<br/>"
                        "Couverture règles : {couverture_gabarit} %<br/>"
                        "Recul voie : {recul_voie} m<br/>"
                        "Recul limites : {recul_limite} m<br/>"
                        "Pleine terre : {pleine_terre} %<br/>"
                        "OAP : {oap}<br/>"
                        "Logements corrigés : {logements_corriges}"
                    )
                },
            )
            st.pydeck_chart(deck, use_container_width=True)

        st.caption(
            "La carte affiche les parcelles cadastrales croisées avec les prescriptions graphiques du PLU."
        )

        if not unresolved.empty:
            with st.expander(
                f"Gabarit PLU à vérifier ({len(unresolved)} parcelles)",
                expanded=filtered.empty,
            ):
                cols = [
                    "reference",
                    "section",
                    "numero",
                    "surface_m2",
                    "zone_plu",
                    "emprise_plu_pct",
                    "niveaux_plu",
                    "hauteur_plu_m",
                    "gabarit_plu_statut",
                    "gabarit_plu_source",
                    "recul_voie_m",
                    "recul_limite_m",
                    "pleine_terre_pct",
                    "oap",
                    "emplacement_reserve",
                    "gabarit_couverture_pct",
                    "regle_zone_emprise",
                    "regle_zone_hauteur",
                    "regle_zone_extrait",
                    "prescription_emprise",
                    "prescription_hauteur",
                    "prescription_url",
                ]
                st.dataframe(
                    unresolved[[c for c in cols if c in unresolved.columns]],
                    use_container_width=True,
                    hide_index=True,
                )
                st.caption(
                    "Ces parcelles ne sont pas supprimées : le moteur n'a pas obtenu une règle chiffrée "
                    "d'emprise et de hauteur sur au moins 98 % de leur surface."
                )

        st.subheader("3. Sélection des parcelles")
        if filtered.empty:
            st.info(
                "Aucune parcelle n'est encore sélectionnable par nombre de logements calculé. "
                "Consulte la liste « Gabarit PLU à vérifier »."
            )
            st.stop()

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
                "emprise_graphique_couverture_pct",
                "niveaux_plu",
                "hauteur_plu_m",
                "surface_brute_plu_m2",
                "surface_enveloppe_prudente_m2",
                "emprise_effective_pct",
                "surface_brute_corrigee_m2",
                "sdp_corrigee_m2",
                "shab_corrigee_m2",
                "recul_voie_m",
                "recul_limite_m",
                "recul_fond_m",
                "pleine_terre_pct",
                "biotope_min_pct",
                "stationnement_par_logement",
                "oap",
                "emplacement_reserve",
                "constructibilite_limitee",
                "sdp_plu_m2",
                "shab_plu_m2",
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
                "logements_plu",
                "logements_corriges",
                "logements_retenus",
                "gabarit_plu_statut",
                "gabarit_plu_source",
                "gabarit_couverture_pct",
                "gabarit_plu_confiance",
                "regle_zone_emprise",
                "regle_zone_hauteur",
                "prescription_emprise",
                "prescription_hauteur",
                "prescription_url",
                "score",
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
                "emprise_graphique_couverture_pct": st.column_config.NumberColumn("Couverture emprise %", format="%.1f %%"),
                "niveaux_plu": st.column_config.NumberColumn("Niveaux PLU"),
                "hauteur_plu_m": st.column_config.NumberColumn("Hauteur PLU m", format="%.1f"),
                "surface_brute_plu_m2": st.column_config.NumberColumn("Surface brute gabarit m²"),
                "surface_enveloppe_prudente_m2": st.column_config.NumberColumn("Enveloppe prudente m²"),
                "emprise_effective_pct": st.column_config.NumberColumn("Emprise effective %", format="%.1f %%"),
                "surface_brute_corrigee_m2": st.column_config.NumberColumn("Surface brute corrigée m²"),
                "sdp_corrigee_m2": st.column_config.NumberColumn("SDP corrigée m²"),
                "shab_corrigee_m2": st.column_config.NumberColumn("SHAB corrigée m²"),
                "recul_voie_m": st.column_config.NumberColumn("Recul voie m"),
                "recul_limite_m": st.column_config.NumberColumn("Recul limites m"),
                "recul_fond_m": st.column_config.NumberColumn("Recul fond m"),
                "pleine_terre_pct": st.column_config.NumberColumn("Pleine terre %", format="%.1f %%"),
                "biotope_min_pct": st.column_config.NumberColumn("Biotope min %", format="%.1f %%"),
                "stationnement_par_logement": st.column_config.NumberColumn("Places / logement"),
                "oap": st.column_config.CheckboxColumn("OAP"),
                "emplacement_reserve": st.column_config.CheckboxColumn("Emplacement réservé"),
                "constructibilite_limitee": st.column_config.CheckboxColumn("Constructibilité limitée"),
                "sdp_plu_m2": st.column_config.NumberColumn("SDP gabarit m²"),
                "shab_plu_m2": st.column_config.NumberColumn("SHAB gabarit m²"),
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
                "logements_plu": st.column_config.NumberColumn("Logements gabarit"),
                "logements_corriges": st.column_config.NumberColumn("Logements corrigés"),
                "logements_retenus": st.column_config.NumberColumn("Logements retenus"),
                "gabarit_plu_statut": st.column_config.TextColumn("Statut gabarit"),
                "gabarit_plu_source": st.column_config.TextColumn("Source gabarit"),
                "gabarit_couverture_pct": st.column_config.NumberColumn("Couverture règles %", format="%.1f %%"),
                "gabarit_plu_confiance": st.column_config.ProgressColumn("Confiance gabarit", min_value=0, max_value=100),
                "regle_zone_emprise": st.column_config.TextColumn("Règle zone emprise"),
                "regle_zone_hauteur": st.column_config.TextColumn("Règle zone hauteur"),
                "prescription_emprise": st.column_config.TextColumn("Prescription emprise"),
                "prescription_hauteur": st.column_config.TextColumn("Prescription hauteur"),
                "prescription_url": st.column_config.LinkColumn("Prescription"),
                "score": st.column_config.ProgressColumn("Score", min_value=0, max_value=100),
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
                "emprise_graphique_couverture_pct",
                "niveaux_plu",
                "hauteur_plu_m",
                "surface_brute_plu_m2",
                "surface_enveloppe_prudente_m2",
                "emprise_effective_pct",
                "surface_brute_corrigee_m2",
                "sdp_corrigee_m2",
                "shab_corrigee_m2",
                "recul_voie_m",
                "recul_limite_m",
                "recul_fond_m",
                "pleine_terre_pct",
                "biotope_min_pct",
                "stationnement_par_logement",
                "oap",
                "emplacement_reserve",
                "constructibilite_limitee",
                "sdp_plu_m2",
                "shab_plu_m2",
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
                "logements_plu",
                "logements_corriges",
                "logements_retenus",
                "gabarit_plu_statut",
                "gabarit_plu_source",
                "gabarit_couverture_pct",
                "gabarit_plu_confiance",
                "regle_zone_emprise",
                "regle_zone_hauteur",
                "prescription_emprise",
                "prescription_hauteur",
                "prescription_url",
                "score",
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
- **Archive CNIG officielle** : récupération du document complet par partition pour accéder aux attributs supplémentaires publiés.
- **Base de règles par zone** : emprise, hauteur, reculs, pleine terre et stationnement sont pré-interprétés une fois par zone.
- **Règlement graphique** : croisement spatial avec les prescriptions CNIG 38 (emprise), 39 (hauteur), 15 (implantation/recul), 42 (biotope), 18 (OAP), 02 (limitations) et 05 (emplacements réservés).
- **Calcul par sous-surfaces** : une prescription locale ne remplace la règle de zone que sur son périmètre réel.
- **Enveloppe prudente** : les reculs extraits peuvent réduire la surface mobilisable avant calcul de la surface brute.
- **Emprise effective** : plafonnée également par la pleine terre / le biotope lorsqu'une valeur minimale est trouvée.
- **Capacité corrigée** : surface brute corrigée → SDP → SHAB → logements.
- **Propriétaires personnes morales** : tentative de rapprochement avec le fichier open data DGFiP des parcelles détenues par des personnes morales ; affichage du nom, de la forme juridique et du SIREN lorsqu'ils sont disponibles.
- **Compatibilité anciens PLU** : lecture de `DESTDOMI` (01 habitat ; 03 mixte habitat/activité) puis analyse des libellés explicites.
- **Exclusions de zonage** : A, N et AUs sont écartées ; U et AUc sont analysées selon la destination logement.
- **Mode haute précision** : une zone ambiguë sans preuve explicite d'habitat est exclue.
- **Terrain nu / bâti** : déterminé par croisement spatial avec les bâtiments cadastraux.
- **Collectif existant** : les parcelles identifiées par la BDNB comme déjà occupées par du « Résidentiel collectif » sont exclues.
- **Autres parcelles bâties** : elles restent éligibles, quelle que soit la taille ou l'emprise du bâtiment.
- **Capacité logements** : SDP = surface brute × ratio SDP ; SHAB = SDP × ratio SHAB ; logements = SHAB ÷ ratio SHAB/logement.
- **Adresse** : recherchée pour les parcelles sélectionnées via le géocodage inverse de la Géoplateforme.
- **Courrier** : généré à partir du modèle SAGEC fourni.

### Limite actuelle
Le filtre Habitat exploite désormais les destinations structurées du PLU lorsqu'elles sont disponibles.
Il ne remplace toutefois pas une faisabilité complète : le nombre de logements reste une estimation de
**présélection**. Une prochaine version pourra extraire automatiquement du règlement : emprise au sol,
hauteur, retraits, pleine terre, stationnement, OAP, prescriptions et servitudes.
        """
    )
