import streamlit as st
import pandas as pd
import requests
import time
import io
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Coordonnées GPS Parcelles",
    page_icon="📍",
    layout="wide"
)

st.title("📍 Géolocalisation de Parcelles Cadastrales")
st.markdown("Importez un fichier Excel contenant vos parcelles pour obtenir leurs coordonnées GPS.")

# ─── helpers ────────────────────────────────────────────────────────────────

def parse_parcelle(raw: str) -> dict | None:
    """
    Decode a parcelle reference like '000ZE0003'.
    Returns dict with section, numero keys, or None if unparseable.
    IGN format: first 3 chars = commune prefix (often '000'), next 2 = section letters, last 4 = parcel number.
    """
    s = str(raw).strip().upper().replace(" ", "")
    if len(s) < 6:
        return None
    # Try to find the section (2 alpha chars) and numero (numeric suffix)
    # Common patterns: 000ZE0003 → section=ZE, numero=0003
    #                  0A0001 → section=0A, numero=0001
    import re
    m = re.match(r'^(\d{0,3})([A-Z]{1,2})(\d+)$', s)
    if m:
        return {"prefix": m.group(1), "section": m.group(2), "numero": m.group(3).zfill(4)}
    return None


def fetch_parcel_coords(postal_code: str, city: str, parcelle: str) -> dict:
    """
    Query the IGN Géoportail WFS API for a cadastral parcel and return centroid lat/lon.
    Falls back to commune geocoding if parcel not found.
    """
    parsed = parse_parcelle(parcelle)
    result = {
        "latitude": None,
        "longitude": None,
        "status": "Non trouvée",
        "source": "",
        "adresse": "",
    }

    # ── Step 1: try IGN cadastre WFS ────────────────────────────────────────
    if parsed:
        section = parsed["section"]
        numero = parsed["numero"]
        # The IGN WFS uses INSEE code; derive it from postal_code + city name via geocoder
        # We'll try the public cadastre API first (data.gouv.fr / API Carto)
        api_url = "https://apicarto.ign.fr/api/cadastre/parcelle"
        params = {
            "code_postal": str(postal_code),
            "nom_commune": city,
            "section": section,
            "numero": numero,
            "_limit": 1,
        }
        try:
            r = requests.get(api_url, params=params, timeout=10)
            if r.status_code == 200:
                data = r.json()
                features = data.get("features", [])
                if features:
                    geom = features[0].get("geometry", {})
                    coords = geom.get("coordinates", [])
                    # centroid of polygon
                    if geom.get("type") == "Polygon" and coords:
                        ring = coords[0]
                        lon = sum(p[0] for p in ring) / len(ring)
                        lat = sum(p[1] for p in ring) / len(ring)
                        props = features[0].get("properties", {})
                        result.update({
                            "latitude": round(lat, 6),
                            "longitude": round(lon, 6),
                            "status": "✅ Trouvée",
                            "source": "IGN Cadastre",
                            "adresse": props.get("numero", "") + " " + props.get("feuille", ""),
                        })
                        return result
                    elif geom.get("type") == "MultiPolygon" and coords:
                        ring = coords[0][0]
                        lon = sum(p[0] for p in ring) / len(ring)
                        lat = sum(p[1] for p in ring) / len(ring)
                        result.update({
                            "latitude": round(lat, 6),
                            "longitude": round(lon, 6),
                            "status": "✅ Trouvée",
                            "source": "IGN Cadastre",
                        })
                        return result
        except Exception as e:
            pass  # fall through to next method

    # ── Step 2: try API Carto with INSEE code lookup ─────────────────────────
    if parsed:
        section = parsed["section"]
        numero = parsed["numero"]
        # Get INSEE code via geo.api.gouv.fr
        insee = None
        try:
            geo_url = "https://geo.api.gouv.fr/communes"
            geo_r = requests.get(geo_url, params={"codePostal": postal_code, "nom": city, "fields": "code", "limit": 1}, timeout=8)
            if geo_r.status_code == 200 and geo_r.json():
                insee = geo_r.json()[0].get("code")
        except Exception:
            pass

        if insee:
            api_url = "https://apicarto.ign.fr/api/cadastre/parcelle"
            params = {
                "code_insee": insee,
                "section": section,
                "numero": numero,
                "_limit": 1,
            }
            try:
                r = requests.get(api_url, params=params, timeout=10)
                if r.status_code == 200:
                    features = r.json().get("features", [])
                    if features:
                        geom = features[0].get("geometry", {})
                        coords = geom.get("coordinates", [])
                        if geom.get("type") in ("Polygon", "MultiPolygon") and coords:
                            ring = coords[0][0] if geom["type"] == "MultiPolygon" else coords[0]
                            lon = sum(p[0] for p in ring) / len(ring)
                            lat = sum(p[1] for p in ring) / len(ring)
                            result.update({
                                "latitude": round(lat, 6),
                                "longitude": round(lon, 6),
                                "status": "✅ Trouvée",
                                "source": "IGN Cadastre (INSEE)",
                                "adresse": insee,
                            })
                            return result
            except Exception:
                pass

    # ── Step 3: fallback — geocode the commune only ──────────────────────────
    try:
        geo_url = "https://api-adresse.data.gouv.fr/search/"
        r = requests.get(geo_url, params={"q": f"{city} {postal_code}", "limit": 1, "type": "municipality"}, timeout=8)
        if r.status_code == 200:
            feats = r.json().get("features", [])
            if feats:
                lon, lat = feats[0]["geometry"]["coordinates"]
                result.update({
                    "latitude": round(lat, 6),
                    "longitude": round(lon, 6),
                    "status": "⚠️ Commune seulement",
                    "source": "Géocodage commune",
                    "adresse": feats[0]["properties"].get("label", ""),
                })
    except Exception:
        pass

    return result


def build_output_excel(df_result: pd.DataFrame) -> bytes:
    """Build a nicely formatted Excel output."""
    wb_out = load_workbook(data_only=True) if False else __import__('openpyxl').Workbook()
    ws = wb_out.active
    ws.title = "Résultats GPS"

    headers = ["Code Postal", "Commune", "Parcelle", "Latitude", "Longitude", "Statut", "Source", "Lien Google Maps"]
    header_fill = PatternFill("solid", start_color="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Alternating row colors
    fill_even = PatternFill("solid", start_color="EBF3FB")
    fill_odd = PatternFill("solid", start_color="FFFFFF")

    for row_idx, row in df_result.iterrows():
        excel_row = row_idx + 2
        fill = fill_even if row_idx % 2 == 0 else fill_odd
        values = [
            row.get("postal_code", ""),
            row.get("city", ""),
            row.get("parcelle", ""),
            row.get("latitude", ""),
            row.get("longitude", ""),
            row.get("status", ""),
            row.get("source", ""),
        ]
        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=excel_row, column=col_idx, value=val)
            cell.fill = fill
            cell.border = border
            cell.alignment = Alignment(horizontal="center")

        # Google Maps link
        lat = row.get("latitude")
        lon = row.get("longitude")
        if lat and lon:
            link_cell = ws.cell(row=excel_row, column=8)
            url = f"https://www.google.com/maps?q={lat},{lon}"
            link_cell.hyperlink = url
            link_cell.value = "📍 Voir sur Maps"
            link_cell.font = Font(color="1155CC", underline="single")
            link_cell.fill = fill
            link_cell.border = border
            link_cell.alignment = Alignment(horizontal="center")

    # Column widths
    widths = [14, 22, 16, 12, 12, 26, 22, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    return buf.read()


# ─── UI ─────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.header("ℹ️ Informations")
    st.markdown("""
**Format du fichier Excel attendu :**

| postal_code | city | parcelle |
|-------------|------|----------|
| 85320 | LA BRETONNIERE | 000ZE0003 |

**Sources utilisées :**
- 🗺️ API Cadastre IGN (apicarto.ign.fr)
- 📍 API Adresse data.gouv.fr (fallback)

**Statuts possibles :**
- ✅ Trouvée — coordonnées exactes
- ⚠️ Commune seulement — parcelle introuvable
- ❌ Erreur — problème de connexion
    """)
    st.divider()
    delay = st.slider("Délai entre requêtes (ms)", 100, 1000, 300, 50,
                      help="Évite la surcharge des APIs publiques")

uploaded = st.file_uploader("📂 Importez votre fichier Excel (.xlsx)", type=["xlsx", "xls"])

if uploaded:
    try:
        df = pd.read_excel(uploaded, dtype=str)
        df.columns = [c.strip().lower() for c in df.columns]

        required = {"postal_code", "city", "parcelle"}
        missing = required - set(df.columns)
        if missing:
            st.error(f"Colonnes manquantes dans le fichier : {', '.join(missing)}")
            st.stop()

        df = df[["postal_code", "city", "parcelle"]].dropna(how="all")
        df = df.fillna("")

        st.success(f"✅ {len(df)} parcelle(s) chargée(s)")
        st.dataframe(df, use_container_width=True, height=200)

        if st.button("🚀 Lancer la géolocalisation", type="primary", use_container_width=True):
            results = []
            progress = st.progress(0, text="Initialisation…")
            status_area = st.empty()

            for i, row in df.iterrows():
                postal = str(row["postal_code"]).strip()
                city = str(row["city"]).strip()
                parcelle = str(row["parcelle"]).strip()

                progress.progress((i + 1) / len(df), text=f"Traitement {i+1}/{len(df)} — {parcelle}")
                status_area.info(f"🔍 Recherche : {parcelle} ({city} {postal})")

                coords = fetch_parcel_coords(postal, city, parcelle)
                results.append({**row.to_dict(), **coords})
                time.sleep(delay / 1000)

            progress.empty()
            status_area.empty()

            df_result = pd.DataFrame(results)

            # Summary metrics
            found = (df_result["status"].str.startswith("✅")).sum()
            partial = (df_result["status"].str.startswith("⚠️")).sum()
            not_found = len(df_result) - found - partial

            col1, col2, col3 = st.columns(3)
            col1.metric("✅ Trouvées", found)
            col2.metric("⚠️ Communes seulement", partial)
            col3.metric("❌ Non trouvées", not_found)

            st.subheader("📋 Résultats")
            display_cols = ["postal_code", "city", "parcelle", "latitude", "longitude", "status", "source"]
            st.dataframe(df_result[display_cols], use_container_width=True)

            # Map preview
            map_df = df_result.dropna(subset=["latitude", "longitude"])
            if not map_df.empty:
                st.subheader("🗺️ Aperçu cartographique")
                st.map(map_df.rename(columns={"latitude": "lat", "longitude": "lon"})[["lat", "lon"]])

            # Download
            st.subheader("⬇️ Télécharger les résultats")
            excel_bytes = build_output_excel(df_result)
            st.download_button(
                label="📥 Télécharger Excel (.xlsx)",
                data=excel_bytes,
                file_name="parcelles_gps.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

            csv_bytes = df_result.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Télécharger CSV",
                data=csv_bytes,
                file_name="parcelles_gps.csv",
                mime="text/csv",
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {e}")
else:
    st.info("👆 Commencez par importer votre fichier Excel.")
    with st.expander("📝 Voir un exemple de fichier attendu"):
        sample = pd.DataFrame({
            "postal_code": ["85320", "85320", "85320"],
            "city": ["LA BRETONNIERE", "PEAULT", "PEAULT"],
            "parcelle": ["000ZE0003", "000ZK0201", "000ZK0038"],
        })
        st.dataframe(sample, use_container_width=True)
