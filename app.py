import streamlit as st
import pandas as pd
import requests
import re
from difflib import get_close_matches

st.set_page_config(page_title="Parcelle → GPS (IGN FIX)", layout="wide")

st.title("📍 Parcelle → GPS (IGN OFFICIAL API)")

# -----------------------------
# FILE READING
# -----------------------------
def read_file(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_file)
        except:
            df = pd.read_csv(uploaded_file, sep=";")

        if len(df.columns) == 1:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep="\t")
    else:
        df = pd.read_excel(uploaded_file)

    df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]
    return df

# -----------------------------
# PARCEL PARSING
# -----------------------------
def parse_parcelle(value):
    value = str(value).upper()
    value = re.sub(r"[^A-Z0-9]", "", value)

    match = re.search(r"([A-Z]{1,2})(\d+)$", value)
    if match:
        section = match.group(1)
        numero = str(int(match.group(2)))
        return section, numero

    return None, None

# -----------------------------
# CITY CORRECTION
# -----------------------------
@st.cache_data
def correct_city(city, postal_code):
    url = "https://geo.api.gouv.fr/communes"

    try:
        r = requests.get(url, params={"codePostal": postal_code}, timeout=10)
        communes = r.json()

        for c in communes:
            if c["nom"].lower() == city.lower():
                return c["nom"]

        names = [c["nom"] for c in communes]
        match = get_close_matches(city, names, n=1, cutoff=0.6)

        if match:
            return match[0]

        return city

    except:
        return city

# -----------------------------
# IGN PARCEL SEARCH (WORKING)
# -----------------------------
def get_coords(city, postal_code, parcelle):

    query = f"{parcelle} {city}"

    url = "https://data.geopf.fr/geocodage/search"

    params = {
        "q": query,
        "limit": 1,
        "index": "parcel"   # 🔥 critical
    }

    try:
        r = requests.get(url, params=params, timeout=10)

        if r.status_code != 200:
            return None, None

        data = r.json()

        if not data.get("features"):
            return None, None

        coords = data["features"][0]["geometry"]["coordinates"]

        return coords[1], coords[0]

    except:
        return None, None

# -----------------------------
# UI
# -----------------------------
uploaded_file = st.file_uploader("📂 Upload Excel or CSV", type=["xlsx", "csv"])

if uploaded_file:

    df = read_file(uploaded_file)

    st.write("Detected columns:", df.columns.tolist())
    st.dataframe(df.head())

    required_cols = ["postal_code", "city", "parcelle"]

    if not all(col in df.columns for col in required_cols):
        st.error(f"❌ File must contain columns: {required_cols}")

    else:
        if st.button("🚀 Process"):

            lats, lons, status, corrected_city = [], [], [], []

            for _, row in df.iterrows():

                city_corr = correct_city(row["city"], row["postal_code"])

                section, numero = parse_parcelle(row["parcelle"])

                if not section:
                    lats.append(None)
                    lons.append(None)
                    status.append("❌ Invalid parcelle format")
                    corrected_city.append(city_corr)
                    continue

                lat, lon = get_coords(city_corr, row["postal_code"], row["parcelle"])

                if lat is None:
                    lats.append(None)
                    lons.append(None)
                    status.append("❌ Parcel not found")
                else:
                    lats.append(lat)
                    lons.append(lon)
                    status.append("✅ Exact (IGN)")

                corrected_city.append(city_corr)

            df["latitude"] = lats
            df["longitude"] = lons
            df["status"] = status
            df["corrected_city"] = corrected_city

            st.success("✅ Processing complete")
            st.dataframe(df)

            csv = df.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Download results", csv, "results.csv", "text/csv")
