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

import pandas as pd
import pydeck as pdk
import requests
import streamlit as st
from docx import Document
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
    page_title="Prospecteur Foncier V3.2.1",
    page_icon="🏗️",
    layout="wide",
)

# --------------------------
# Helpers réseau
# --------------------------
BDNB_API = "https://api.bdnb.io/v1/bdnb/donnees/batiment_groupe_complet"

HEADERS = {
    "User-Agent": "ProspecteurFoncier/3.2.1 (Streamlit; donnees publiques)",
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
            "fields": "nom,code,codesPostaux,population",
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

def gpu_get(layer, params, timeout=90):
    url = f"{GPU_API}/{layer}"
    r = requests.get(url, params=params, headers=HEADERS, timeout=timeout)
    if r.ok:
        return r.json()

    # Certaines requêtes géométriques sont mieux acceptées en POST.
    r2 = requests.post(url, json=params, headers=HEADERS, timeout=timeout)
    r2.raise_for_status()
    return r2.json()


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def fetch_gpu_document(insee, commune_geojson):
    features = commune_geojson.get("features", [])
    if not features:
        return None, []

    commune_geom = shape(features[0]["geometry"])
    point = commune_geom.representative_point()
    geom = json.dumps(mapping(point), separators=(",", ":"))

    data = gpu_get("document", {"geom": geom})
    docs = data.get("features", []) or []

    # Retirer les documents non opérationnels pour le zonage local si possible.
    local_docs = []
    for f in docs:
        p = f.get("properties", {}) or {}
        typedoc = str(p.get("typedoc") or p.get("type_doc") or "").upper()
        if "SCOT" in typedoc:
            continue
        local_docs.append(f)

    candidates = local_docs or docs
    if not candidates:
        # Fallback utile pour de nombreux PLU communaux.
        return f"DU_{insee}", []

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
    if not partition:
        partition = f"DU_{insee}"
    return partition, candidates


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def fetch_gpu_zones(partition):
    # API GPU : limite haute de plusieurs milliers d'objets.
    all_features = []
    start = 0
    while True:
        params = {"partition": partition}
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

    # 0. Type de zone : on exclut les zones non immédiatement destinées à bâtir du logement.
    if typezone in {"A", "N"}:
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — zone agricole/naturelle",
            "habitat_preuve": f"TYPEZONE={typezone}",
            "habitat_confiance": 100,
            "typezone_normalise": typezone,
        }

    if typezone == "AUs":
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — zone AU bloquée",
            "habitat_preuve": "TYPEZONE=AUs : ouverture subordonnée à modification/révision du PLU",
            "habitat_confiance": 100,
            "typezone_normalise": typezone,
        }

    if typezone in {"AUc", "AU"} and not include_au:
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — zone AU désactivée dans les critères",
            "habitat_preuve": f"TYPEZONE={typezone}",
            "habitat_confiance": 100,
            "typezone_normalise": typezone,
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
        }

    if destoui & housing_codes:
        return {
            "habitat_eligible": True,
            "habitat_statut": "Habitat autorisé",
            "habitat_preuve": f"DESTOUI={','.join(sorted(destoui & housing_codes))}",
            "habitat_confiance": 100,
            "typezone_normalise": typezone,
        }

    if destcdt & housing_codes:
        if include_conditionnel:
            return {
                "habitat_eligible": True,
                "habitat_statut": "Habitat autorisé sous conditions",
                "habitat_preuve": f"DESTCDT={','.join(sorted(destcdt & housing_codes))}",
                "habitat_confiance": 90,
                "typezone_normalise": typezone,
            }
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — habitat seulement sous conditions",
            "habitat_preuve": f"DESTCDT={','.join(sorted(destcdt & housing_codes))}",
            "habitat_confiance": 90,
            "typezone_normalise": typezone,
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
        }
    if dd in {"3", "03"}:
        return {
            "habitat_eligible": True,
            "habitat_statut": "Habitat/mixte — vocation dominante",
            "habitat_preuve": "DESTDOMI=03 (mixte habitat/activité)",
            "habitat_confiance": 90,
            "typezone_normalise": typezone,
        }
    if dd in {"2", "02", "4", "04", "5", "05", "7", "07", "8", "08", "9", "09", "10"}:
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — vocation dominante non habitat",
            "habitat_preuve": f"DESTDOMI={destdomi}",
            "habitat_confiance": 95,
            "typezone_normalise": typezone,
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
        }

    if negative and not positive:
        return {
            "habitat_eligible": False,
            "habitat_statut": "Exclue — libellé PLU non résidentiel",
            "habitat_preuve": f"{libelle} — {libelong}".strip(" —"),
            "habitat_confiance": 80,
            "typezone_normalise": typezone,
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

        # Méthode de capacité définie par l'utilisateur :
        # Surface brute -> SDP -> SHAB -> nombre de logements.
        # À ce stade, faute de moteur de gabarit PLU complet, la surface cadastrale
        # est utilisée comme base de "surface brute" de présélection.
        surface_brute = float(surface)
        sdp_estimee = surface_brute * (float(ratio_sdp_pct) / 100.0)
        shab_estimee = sdp_estimee * (float(ratio_shab_pct) / 100.0)
        logements = int(max(0, math.floor(shab_estimee / max(float(shab_par_logement), 1.0))))

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
                "surface_brute_m2": round(surface_brute),
                "sdp_estimee_m2": round(sdp_estimee),
                "shab_estimee_m2": round(shab_estimee),
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
                "sdp_m2": round(sdp_estimee),
                "shab_m2": round(shab_estimee),
                "zone": zr["libelle"],
                "typezone": zr["typezone"],
                "habitat": zr["habitat_statut"],
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
        "reglement_url",
        "date_zone",
        "logements_estimes",
        "score",
        "latitude",
        "longitude",
        "adresse",
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
st.title("🏗️ Prospecteur Foncier — V3.2.1 — Filtre Habitat")
st.caption(
    "Cadastre réel + PLU/PLUi + filtre Habitat renforcé + calcul SDP/SHAB + exclusion du collectif existant."
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

st.subheader("1. Critères de recherche")

c1, c2 = st.columns(2)
with c1:
    min_log = st.number_input("Logements minimum", min_value=0, value=20, step=1)
with c2:
    max_log = st.number_input("Logements maximum", min_value=0, value=50, step=1)

st.markdown("#### Ratios de calcul de capacité")
r1, r2, r3 = st.columns(3)
with r1:
    ratio_sdp_pct = st.number_input(
        "SDP / surface brute (%)",
        min_value=1.0,
        max_value=300.0,
        value=80.0,
        step=1.0,
        help="Par défaut : SDP estimée = 80 % de la surface brute.",
    )
with r2:
    ratio_shab_pct = st.number_input(
        "SHAB / SDP (%)",
        min_value=1.0,
        max_value=100.0,
        value=80.0,
        step=1.0,
        help="Par défaut : SHAB estimée = 80 % de la SDP.",
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

st.caption(
    f"Calcul utilisé : SDP = surface brute × {ratio_sdp_pct:.0f} % ; "
    f"SHAB = SDP × {ratio_shab_pct:.0f} % ; "
    f"logements = SHAB ÷ {shab_par_logement:.0f} m²."
)

c5, c6, c7 = st.columns(3)
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
with c7:
    habitat_only = st.checkbox(
        "Conserver uniquement les parcelles destinées à l'habitat",
        value=True,
        help=(
            "Filtre haute précision utilisant les destinations CNIG du PLU/PLUi, "
            "l'ancienne vocation dominante et les libellés de zone."
        ),
    )

h1, h2 = st.columns(2)
with h1:
    include_conditionnel = st.checkbox(
        "Inclure l'habitat autorisé sous conditions",
        value=True,
        help=(
            "Conserve les zones dont le standard CNIG indique Habitation/Logement "
            "dans DESTCDT (destination autorisée sous conditions)."
        ),
    )
with h2:
    st.info(
        "Mode strict : si le PLU ne fournit pas de preuve suffisamment explicite que le logement "
        "est autorisé, la zone est écartée."
    )

st.caption(
    "Filtre bâti : une parcelle déjà construite reste éligible, sauf si la BDNB identifie "
    "un bâtiment en « Résidentiel collectif » sur cette parcelle."
)

st.warning(
    "Le filtre Habitat réduit fortement les faux positifs : il exploite d'abord les champs structurés "
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
            partition, documents = fetch_gpu_document(insee, commune_geojson)
        with st.spinner(f"Chargement du zonage PLU/PLUi ({partition})…"):
            zones_geojson = fetch_gpu_zones(partition)

        if not zones_geojson.get("features"):
            municipality = fetch_gpu_municipality(insee)
            st.error(
                "Aucun zonage PLU/PLUi exploitable n'a été renvoyé par le Géoportail de l'Urbanisme "
                f"pour cette commune. Partition testée : {partition}."
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
    # Si le filtre Habitat n'a trouvé aucune parcelle, ne pas laisser Pandas provoquer un KeyError.
    if results.empty or "logements_estimes" not in results.columns:
        st.subheader("2. Résultats sur le cadastre réel")
        st.warning(
            "Aucune parcelle n'a été retenue avec le filtre Habitat actuel pour cette commune. "
            "Ce n'est plus une erreur de l'application : cela signifie soit que les métadonnées du PLU "
            "ne permettent pas d'identifier explicitement la destination logement, soit que le filtre est trop strict."
        )
        st.info(
            "Essaie d'abord de décocher « Conserver uniquement les parcelles destinées à l'habitat » "
            "pour vérifier que le cadastre et le zonage sont bien chargés. "
            "Ensuite nous pourrons adapter le moteur Habitat à la structure du PLU de cette commune."
        )
        st.stop()

    # Appliquer les critères courants sans relancer les appels réseau.
    filtered = results[
        (results["logements_estimes"] >= min_log)
        & (results["logements_estimes"] <= max_log)
    ].copy()

    # Filtre Habitat haute précision.
    if habitat_only:
        filtered = filtered[filtered["habitat_eligible"] == True].copy()

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
    m2.metric("Parcelles analysées", len(results))
    m3.metric(
        "Habitat confirmé / probable",
        int(filtered["habitat_eligible"].sum()) if "habitat_eligible" in filtered.columns else 0,
    )
    m4.metric("Document urbanisme", st.session_state.get("analysis_partition", "—"))

    if st.session_state.get("bdnb_error"):
        st.warning(
            "La BDNB n'a pas répondu pendant cette analyse. Dans ce cas, le logiciel ne peut pas "
            "garantir l'exclusion des résidences collectives. Détail : "
            + st.session_state["bdnb_error"]
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
                        "Surface brute : {surface_m2} m²<br/>"
                        "SDP : {sdp_m2} m²<br/>"
                        "SHAB : {shab_m2} m²<br/>"
                        "Habitat : {habitat}<br/>"
                        "Confiance habitat : {confiance_habitat} %<br/>"
                        "Collectif existant : {collectif}<br/>"
                        "Zone : {typezone} {zone}<br/>"
                        "Potentiel : {logements} logements"
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
                "surface_m2",
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
                "classe_zone",
                "logements_estimes",
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
                "surface_m2": st.column_config.NumberColumn("Surface brute m²"),
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
                "classe_zone": st.column_config.TextColumn("Qualification"),
                "logements_estimes": st.column_config.NumberColumn("Logements estimés"),
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
                "surface_m2",
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
                "classe_zone",
                "logements_estimes",
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
- **Filtre Habitat structuré** : lecture des champs CNIG `DESTOUI`, `DESTCDT`, `DESTNON` ; les codes 20/21 correspondent à Habitation/Logement.
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
