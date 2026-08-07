
import io
import zipfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document

APP_DIR = Path(__file__).parent
DEFAULT_DATA = APP_DIR / "data" / "parcelles_exemple.csv"
LETTER_TEMPLATE = APP_DIR / "assets" / "lettre_sagec_modele.docx"

st.set_page_config(page_title="Prospecteur Foncier", page_icon="🏗️", layout="wide")

st.title("🏗️ Prospecteur Foncier")
st.caption("MVP – ciblage de parcelles constructibles, sélection et génération de courriers personnalisés.")

# ---------------------------
# Helpers
# ---------------------------
def load_data(uploaded):
    if uploaded is not None:
        return pd.read_csv(uploaded)
    return pd.read_csv(DEFAULT_DATA)

def bool_label(v):
    return "Oui" if bool(v) else "Non"

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
    doc = Document(LETTER_TEMPLATE)

    adresse = str(row["adresse"]).strip()
    ville = str(row["ville"]).strip()
    section = str(row["section"]).strip()
    numero = str(row["numero"]).strip()

    # Replacements that match the supplied SAGEC example as closely as possible.
    replacements = {
        "Adresse….": adresse,
        "Adresse...": adresse,
        "Ville……": ville,
        "Ville......": ville,
        "Anglet, le 19/05/2026.": f"{ville_signature}, le {date.today().strftime('%d/%m/%Y')}.",
        "Objet : Votre propriété à …………….": f"Objet : Votre propriété à {adresse}",
        "Objet : Votre propriété à .................": f"Objet : Votre propriété à {adresse}",
        "votre propriété située à ………": f"votre propriété située à {adresse}",
        "votre propriété située à ........": f"votre propriété située à {adresse}",
        "cadastrée section………………………,": f"cadastrée section {section} n° {numero},",
        "cadastrée section.................................,": f"cadastrée section {section} n° {numero},",
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

    # Also handle embedded lines by targeted fallback.
    for p in doc.paragraphs:
        t = p.text
        if "Objet : Votre propriété à" in t and adresse not in t:
            p.text = f"Objet : Votre propriété à {adresse}"
        if "Dans le cadre de nos recherches foncières" in t:
            # Keep original wording but ensure address/parcel are personalized
            p.text = (
                f"Dans le cadre de nos recherches foncières, nous avons identifié que votre propriété "
                f"située à {adresse}, cadastrée section {section} n° {numero}, comme pouvant présenter "
                f"un fort potentiel pour le développement d'une opération immobilière."
            )

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# ---------------------------
# Sidebar
# ---------------------------
with st.sidebar:
    st.header("Données")
    uploaded = st.file_uploader(
        "Importer les parcelles (CSV)",
        type=["csv"],
        help="Colonnes attendues : ville, adresse, section, numero, surface_m2, zone_plu, constructible, terrain_bati, pc_obtenu, logements_estimes, score."
    )
    st.info(
        "Le MVP utilise un CSV d'exemple. La connexion directe au cadastre, au PLU/PLUi "
        "et aux autorisations d'urbanisme est prévue dans l'étape suivante."
    )

df = load_data(uploaded)

required = {
    "ville", "adresse", "section", "numero", "surface_m2", "zone_plu",
    "constructible", "terrain_bati", "pc_obtenu", "logements_estimes", "score"
}
missing = required - set(df.columns)
if missing:
    st.error(f"Colonnes manquantes : {', '.join(sorted(missing))}")
    st.stop()

# Normalize booleans if imported as strings
for c in ["constructible", "terrain_bati", "pc_obtenu"]:
    if df[c].dtype == object:
        df[c] = df[c].astype(str).str.lower().map(
            {"true": True, "false": False, "oui": True, "non": False, "1": True, "0": False}
        ).fillna(False)

# ---------------------------
# Search criteria
# ---------------------------
st.subheader("1. Critères de recherche")
c1, c2, c3, c4 = st.columns(4)
with c1:
    villes = sorted(df["ville"].dropna().astype(str).unique())
    ville = st.selectbox("Ville", villes)
with c2:
    min_log = st.number_input("Logements minimum", min_value=0, value=20, step=1)
with c3:
    max_log = st.number_input("Logements maximum", min_value=0, value=50, step=1)
with c4:
    min_surface = st.number_input("Surface minimale (m²)", min_value=0, value=1000, step=100)

c5, c6, c7 = st.columns(3)
with c5:
    pc = st.selectbox("Permis de construire", ["Indifférent", "Avec PC", "Sans PC"])
with c6:
    terrain = st.selectbox("Type de terrain", ["Indifférent", "Terrain nu", "Terrain bâti"])
with c7:
    score_min = st.slider("Score minimum", 0, 100, 60)

# Filter: non-constructible parcels are automatically removed.
f = df[
    (df["ville"].astype(str) == ville)
    & (df["constructible"] == True)
    & (df["logements_estimes"] >= min_log)
    & (df["logements_estimes"] <= max_log)
    & (df["surface_m2"] >= min_surface)
    & (df["score"] >= score_min)
].copy()

if pc == "Avec PC":
    f = f[f["pc_obtenu"] == True]
elif pc == "Sans PC":
    f = f[f["pc_obtenu"] == False]

if terrain == "Terrain nu":
    f = f[f["terrain_bati"] == False]
elif terrain == "Terrain bâti":
    f = f[f["terrain_bati"] == True]

f = f.sort_values(["score", "logements_estimes"], ascending=[False, False])

st.subheader("2. Parcelles correspondant aux critères")
m1, m2, m3 = st.columns(3)
m1.metric("Parcelles retenues", len(f))
m2.metric("Surface totale", f"{int(f['surface_m2'].sum()):,} m²".replace(",", " "))
m3.metric("Logements théoriques", int(f["logements_estimes"].sum()))

if f.empty:
    st.warning("Aucune parcelle ne correspond aux critères actuels.")
    st.stop()

display = f[[
    "adresse", "section", "numero", "surface_m2", "zone_plu",
    "logements_estimes", "pc_obtenu", "terrain_bati", "score"
]].copy()
display["pc_obtenu"] = display["pc_obtenu"].map(bool_label)
display["terrain_bati"] = display["terrain_bati"].map(bool_label)

st.dataframe(
    display.rename(columns={
        "adresse": "Adresse",
        "section": "Section",
        "numero": "N°",
        "surface_m2": "Surface m²",
        "zone_plu": "Zone PLU",
        "logements_estimes": "Logements estimés",
        "pc_obtenu": "PC obtenu",
        "terrain_bati": "Bâti",
        "score": "Score"
    }),
    use_container_width=True,
    hide_index=True
)

# ---------------------------
# Select parcels
# ---------------------------
st.subheader("3. Sélection des parcelles à prospecter")
options = {
    f"{r['adresse']} — {r['section']} {r['numero']} — {int(r['logements_estimes'])} logts — score {int(r['score'])}": i
    for i, r in f.iterrows()
}
selected_labels = st.multiselect("Choisir les parcelles", list(options.keys()))
selected_idx = [options[x] for x in selected_labels]
selected = f.loc[selected_idx] if selected_idx else f.iloc[0:0]

# ---------------------------
# Letter generation
# ---------------------------
st.subheader("4. Génération des courriers personnalisés")
l1, l2 = st.columns(2)
with l1:
    signataire = st.text_input("Signataire", "Nicolas PEDROT")
    fonction = st.text_input("Fonction", "Directeur")
with l2:
    email = st.text_input("E-mail", "Nicolas.pedrot@sagec.fr")
    ville_signature = st.text_input("Ville de signature", "Anglet")

st.caption(
    "Le texte reprend la lettre SAGEC fournie comme modèle. "
    "Le destinataire reste « Madame, Monsieur » et le courrier utilise l'adresse de la parcelle/maison."
)

if st.button("Générer les courriers", type="primary", disabled=selected.empty):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for _, row in selected.iterrows():
            letter = generate_letter(row, signataire, fonction, email, ville_signature)
            safe_addr = "".join(ch if ch.isalnum() else "_" for ch in str(row["adresse"])).strip("_")
            filename = f"Lettre_{row['ville']}_{row['section']}{row['numero']}_{safe_addr}.docx"
            z.writestr(filename, letter.getvalue())
    zip_buffer.seek(0)
    st.success(f"{len(selected)} courrier(s) généré(s).")
    st.download_button(
        "Télécharger les courriers (.zip)",
        data=zip_buffer.getvalue(),
        file_name=f"courriers_prospection_{ville}.zip",
        mime="application/zip"
    )

# ---------------------------
# Future integrations
# ---------------------------
with st.expander("Étape suivante : connexion aux données publiques"):
    st.markdown(
        """
        Cette V1 est volontairement exploitable sans API externe. Pour passer à la version opérationnelle :
        - récupérer automatiquement les parcelles cadastrales de la commune ;
        - superposer le zonage PLU/PLUi en vigueur ;
        - exclure automatiquement les zones non constructibles ;
        - intégrer les règles de gabarit/emprise/retrait/hauteur ;
        - détecter les permis de construire connus ;
        - calculer un potentiel de logements avec un indice de confiance ;
        - résoudre l'adresse postale correspondant à chaque parcelle ;
        - conserver l'historique des courriers et relances.
        """
    )
