import os
import io
import math
import json
import time
import logging
from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd
import requests

# =========================
# Configuration
# =========================
START_YEAR = 2002
END_YEAR = 2023
OUTPUT_XLSX = "etude_quantitative_livrable2_panel_final.xlsx"
OUTPUT_PANEL_CSV = "panel_final.csv"
CACHE_DIR = "cache"
TIMEOUT = 30

os.makedirs(CACHE_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

COUNTRIES = [
    "France", "Germany", "Italy", "Spain", "Portugal", "Greece", "Ireland", "Netherlands", "Belgium", "Austria",
    "Denmark", "Sweden", "Finland", "Norway", "Switzerland", "United Kingdom", "United States", "Canada", "Japan",
    "Australia", "New Zealand", "Brazil", "Mexico", "Chile", "South Africa", "Turkey", "India", "China", "Indonesia", "Korea, Rep."
]

WDI_INDICATORS = {
    "NY.GDP.PCAP.CD": "gdp_per_capita",
    "NY.GDP.MKTP.KD.ZG": "gdp_growth",
    "FP.CPI.TOTL.ZG": "inflation",
    "NE.CON.GOVT.ZS": "gov_consumption_gdp",
    "BN.CAB.XOKA.GD.ZS": "current_account_gdp",
    "DT.DOD.DECT.GN.ZS": "external_debt_gni",
    "SL.UEM.TOTL.ZS": "unemployment",
    "NE.TRD.GNFS.ZS": "trade_openess",
    "EN.ATM.CO2E.PC": "co2_per_capita",
    "EG.FEC.RNEW.ZS": "renewable_energy",
    "AG.LND.FRST.ZS": "forest_area",
    "EN.ATM.PM25.MC.M3": "pm25",
    "SP.DYN.LE00.IN": "life_expectancy",
    "SE.TER.ENRR": "tertiary_enrollment",
}

WGI_INDICATORS = {
    "VA.EST": "voice_accountability",
    "PV.EST": "political_stability",
    "GE.EST": "government_effectiveness",
    "RQ.EST": "regulatory_quality",
    "RL.EST": "rule_of_law",
    "CC.EST": "control_corruption",
}

RATING_MAP = {
    "AAA": 21, "AA+": 20, "AA": 19, "AA-": 18, "A+": 17, "A": 16, "A-": 15,
    "BBB+": 14, "BBB": 13, "BBB-": 12, "BB+": 11, "BB": 10, "BB-": 9,
    "B+": 8, "B": 7, "B-": 6, "CCC+": 5, "CCC": 4, "CCC-": 3, "CC": 2, "C": 1,
    "D": 0, "SD": 0, "RD": 0,
    "Aaa": 21, "Aa1": 20, "Aa2": 19, "Aa3": 18, "A1": 17, "A2": 16, "A3": 15,
    "Baa1": 14, "Baa2": 13, "Baa3": 12, "Ba1": 11, "Ba2": 10, "Ba3": 9,
    "B1": 8, "B2": 7, "B3": 6, "Caa1": 5, "Caa2": 4, "Caa3": 3, "Ca": 2,
}


def safe_get_json(url: str, cache_name: Optional[str] = None):
    if cache_name:
        cache_path = os.path.join(CACHE_DIR, cache_name)
        if os.path.exists(cache_path):
            logging.info(f"Cache hit: {cache_name}")
            with open(cache_path, "r", encoding="utf-8") as f:
                return json.load(f)
    try:
        r = requests.get(url, timeout=TIMEOUT)
        r.raise_for_status()
        data = r.json()
        if cache_name:
            with open(cache_path, "w", encoding="utf-8") as f:
                json.dump(data, f)
        return data
    except Exception as e:
        logging.warning(f"GET JSON failed {url}: {e}")
        return None


def safe_get_text(url: str, cache_name: Optional[str] = None):
    if cache_name:
        cache_path = os.path.join(CACHE_DIR, cache_name)
        if os.path.exists(cache_path):
            logging.info(f"Cache hit: {cache_name}")
            with open(cache_path, "r", encoding="utf-8") as f:
                return f.read()
    try:
        r = requests.get(url, timeout=TIMEOUT)
        r.raise_for_status()
        text = r.text
        if cache_name:
            with open(cache_path, "w", encoding="utf-8") as f:
                f.write(text)
        return text
    except Exception as e:
        logging.warning(f"GET text failed {url}: {e}")
        return None


def get_country_meta() -> pd.DataFrame:
    url = "https://api.worldbank.org/v2/country?format=json&per_page=400"
    data = safe_get_json(url, "wb_countries.json")
    rows = []
    if data and len(data) > 1:
        for c in data[1]:
            rows.append({
                "country_name": c.get("name"),
                "iso3": c.get("id"),
                "region": (c.get("region") or {}).get("value"),
                "income_group": (c.get("incomeLevel") or {}).get("value"),
            })
    df = pd.DataFrame(rows)
    return df[df["country_name"].isin(COUNTRIES)].copy()


def get_wdi_data(iso3_list: List[str]) -> pd.DataFrame:
    all_rows = []
    for indicator in WDI_INDICATORS:
        iso_chunk = ";".join(iso3_list)
        url = (
            f"https://api.worldbank.org/v2/country/{iso_chunk}/indicator/{indicator}"
            f"?format=json&per_page=20000"
        )
        data = safe_get_json(url, f"wdi_{indicator.replace('.', '_')}.json")
        if not data or len(data) < 2:
            continue
        for row in data[1]:
            year = int(row["date"])
            if START_YEAR <= year <= END_YEAR:
                all_rows.append({
                    "iso3": row.get("countryiso3code"),
                    "year": year,
                    "indicator_code": indicator,
                    "indicator_name": WDI_INDICATORS[indicator],
                    "value": row.get("value"),
                    "source_url": url,
                })
    return pd.DataFrame(all_rows)


def get_wgi_data(iso3_list: List[str]) -> pd.DataFrame:
    rows = []
    for ind, ind_name in WGI_INDICATORS.items():
        # Using WGI bulk CSV from World Bank API-like endpoint
        url = f"https://api.worldbank.org/v2/en/sources/3/country/{';'.join(iso3_list)}/indicator/{ind}?format=json&per_page=20000"
        data = safe_get_json(url, f"wgi_{ind.replace('.', '_')}.json")
        if not data or len(data) < 2:
            continue
        for r in data[1]:
            try:
                year = int(r.get("date"))
            except Exception:
                continue
            if START_YEAR <= year <= END_YEAR:
                rows.append({
                    "iso3": r.get("countryiso3code"),
                    "year": year,
                    "indicator_code": ind,
                    "indicator_name": ind_name,
                    "value": r.get("value"),
                    "source_url": url,
                })
    return pd.DataFrame(rows)


def normalize_rating_text(rt: str) -> str:
    if rt is None or (isinstance(rt, float) and math.isnan(rt)):
        return ""
    t = str(rt).strip()
    for suffix in ["(EXP)", "u", "*", "(P)"]:
        t = t.replace(suffix, "")
    return t.strip()


def parse_ratings_csv(url: str, agency: str, country_meta: pd.DataFrame) -> pd.DataFrame:
    text = safe_get_text(url, f"ratings_{agency}.csv")
    if not text:
        return pd.DataFrame(columns=["agency", "country_name", "iso3", "rating_date", "year", "rating_text", "rating_score", "source_url"])
    try:
        df = pd.read_csv(io.StringIO(text))
    except Exception as e:
        logging.warning(f"Failed reading ratings CSV {agency}: {e}")
        return pd.DataFrame(columns=["agency", "country_name", "iso3", "rating_date", "year", "rating_text", "rating_score", "source_url"])

    cols = {c.lower(): c for c in df.columns}
    # heuristics for column names on ratingshistory datasets
    c_country = next((cols[c] for c in cols if "country" in c), None)
    c_date = next((cols[c] for c in cols if "date" in c), None)
    c_rating = next((cols[c] for c in cols if "rating" in c), None)

    if not c_country or not c_date or not c_rating:
        logging.warning(f"Ratings CSV structure unexpected for {agency}.")
        return pd.DataFrame(columns=["agency", "country_name", "iso3", "rating_date", "year", "rating_text", "rating_score", "source_url"])

    out = df[[c_country, c_date, c_rating]].copy()
    out.columns = ["country_name", "rating_date", "rating_text"]
    out["rating_date"] = pd.to_datetime(out["rating_date"], errors="coerce")
    out = out.dropna(subset=["rating_date"])
    out["year"] = out["rating_date"].dt.year
    out = out[(out["year"] >= START_YEAR) & (out["year"] <= END_YEAR)]

    out["rating_text"] = out["rating_text"].apply(normalize_rating_text)
    out["rating_score"] = out["rating_text"].map(RATING_MAP)

    mapper = country_meta.set_index("country_name")["iso3"].to_dict()
    out["iso3"] = out["country_name"].map(mapper)
    out["agency"] = agency
    out["source_url"] = url
    out = out[["agency", "country_name", "iso3", "rating_date", "year", "rating_text", "rating_score", "source_url"]]
    return out


def aggregate_annual_ratings(ratings_raw: pd.DataFrame, iso3_list: List[str]) -> pd.DataFrame:
    base = pd.MultiIndex.from_product([iso3_list, list(range(START_YEAR, END_YEAR + 1))], names=["iso3", "year"]).to_frame(index=False)

    if ratings_raw.empty:
        out = base.copy()
        for c in ["sp_rating_text", "sp_rating_score", "moodys_rating_text", "moodys_rating_score", "fitch_rating_text", "fitch_rating_score", "avg_rating_score", "rating_dispersion"]:
            out[c] = pd.NA
        return out

    agency_map = {"sp": "sp", "s&p": "sp", "moodys": "moodys", "moodys": "moodys", "fitch": "fitch"}
    ratings_raw = ratings_raw.copy()
    ratings_raw["agency_norm"] = ratings_raw["agency"].str.lower().map(lambda x: agency_map.get(x, x))

    latest = ratings_raw.sort_values("rating_date").groupby(["iso3", "year", "agency_norm"], as_index=False).tail(1)
    ptxt = latest.pivot_table(index=["iso3", "year"], columns="agency_norm", values="rating_text", aggfunc="first")
    pval = latest.pivot_table(index=["iso3", "year"], columns="agency_norm", values="rating_score", aggfunc="first")

    merged = base.merge(ptxt.add_suffix("_rating_text"), on=["iso3", "year"], how="left")
    merged = merged.merge(pval.add_suffix("_rating_score"), on=["iso3", "year"], how="left")

    rename_map = {
        "sp_rating_text": "sp_rating_text",
        "moodys_rating_text": "moodys_rating_text",
        "fitch_rating_text": "fitch_rating_text",
        "sp_rating_score": "sp_rating_score",
        "moodys_rating_score": "moodys_rating_score",
        "fitch_rating_score": "fitch_rating_score",
    }
    for col in rename_map:
        if col not in merged.columns:
            merged[col] = pd.NA

    score_cols = ["sp_rating_score", "moodys_rating_score", "fitch_rating_score"]
    merged["avg_rating_score"] = merged[score_cols].mean(axis=1, skipna=True)
    merged["rating_dispersion"] = merged[score_cols].std(axis=1, skipna=True)
    return merged


def main():
    logging.info("Building quantitative panel...")
    countries = get_country_meta()
    iso3_list = countries["iso3"].dropna().unique().tolist()

    wdi_raw = get_wdi_data(iso3_list)
    wgi_raw = get_wgi_data(iso3_list)

    rating_urls = {
        "fitch": "https://www.ratingshistory.info/data/fitch-sovereign-ratings.csv",
        "moodys": "https://www.ratingshistory.info/data/moodys-sovereign-ratings.csv",
        "sp": "https://www.ratingshistory.info/data/sp-sovereign-ratings.csv",
    }
    ratings_raw = pd.concat(
        [parse_ratings_csv(u, a, countries) for a, u in rating_urls.items()],
        ignore_index=True,
    )

    ratings_annual = aggregate_annual_ratings(ratings_raw, iso3_list)

    panel = pd.MultiIndex.from_product([iso3_list, list(range(START_YEAR, END_YEAR + 1))], names=["iso3", "year"]).to_frame(index=False)
    panel = panel.merge(countries[["country_name", "iso3", "region", "income_group"]], on="iso3", how="left")
    panel = panel.merge(ratings_annual, on=["iso3", "year"], how="left")

    wdi_wide = wdi_raw.pivot_table(index=["iso3", "year"], columns="indicator_name", values="value", aggfunc="mean").reset_index()
    panel = panel.merge(wdi_wide, on=["iso3", "year"], how="left")

    wgi_wide = wgi_raw.pivot_table(index=["iso3", "year"], columns="indicator_name", values="value", aggfunc="mean").reset_index()
    panel = panel.merge(wgi_wide, on=["iso3", "year"], how="left")

    desc_vars = [c for c in panel.columns if c not in ["country_name", "iso3", "year", "region", "income_group"]]
    desc = []
    for v in desc_vars:
        s = pd.to_numeric(panel[v], errors="coerce")
        desc.append({
            "variable": v,
            "n": int(s.notna().sum()),
            "mean": s.mean(),
            "std": s.std(),
            "min": s.min(),
            "p25": s.quantile(0.25),
            "median": s.median(),
            "p75": s.quantile(0.75),
            "max": s.max(),
            "missing_rate": s.isna().mean(),
        })
    descriptive_stats = pd.DataFrame(desc)

    missing = []
    for v in desc_vars:
        miss = panel[panel[v].isna()].groupby("iso3").size().sort_values(ascending=False)
        missing.append({
            "variable": v,
            "missing_count": int(panel[v].isna().sum()),
            "missing_rate": panel[v].isna().mean(),
            "countries_most_missing": ", ".join(miss.head(5).index.tolist()) if len(miss) else "",
        })
    missing_values = pd.DataFrame(missing)

    must_have = ["avg_rating_score", "gdp_per_capita", "gdp_growth", "inflation", "government_effectiveness", "rule_of_law"]
    regression_ready = panel.dropna(subset=must_have).copy()

    panel.to_csv(OUTPUT_PANEL_CSV, index=False)

    sources = []
    for k, v in WDI_INDICATORS.items():
        sources.append({"variable": v, "source": "World Bank WDI", "URL/API": f"https://api.worldbank.org/v2/indicator/{k}", "fréquence": "annual", "période disponible": "varies", "commentaire qualité": "Public data, may contain missing values"})
    for k, v in WGI_INDICATORS.items():
        sources.append({"variable": v, "source": "World Bank WGI", "URL/API": f"https://api.worldbank.org/v2/sources/3/indicator/{k}", "fréquence": "annual", "période disponible": "varies", "commentaire qualité": "Perception-based governance indicators"})
    for agency, url in rating_urls.items():
        sources.append({"variable": f"{agency}_rating", "source": "ratingshistory.info", "URL/API": url, "fréquence": "event -> annualized", "période disponible": "varies", "commentaire qualité": "Public scraped file, verify coverage"})
    sources_df = pd.DataFrame(sources)

    esg_codes = {"EN.ATM.CO2E.PC", "EG.FEC.RNEW.ZS", "AG.LND.FRST.ZS", "EN.ATM.PM25.MC.M3", "SP.DYN.LE00.IN", "SE.TER.ENRR"}
    esg_public_raw = wdi_raw[wdi_raw["indicator_code"].isin(esg_codes)].copy()
    esg_public_raw["pillar"] = esg_public_raw["indicator_code"].map({
        "EN.ATM.CO2E.PC": "E", "EG.FEC.RNEW.ZS": "E", "AG.LND.FRST.ZS": "E", "EN.ATM.PM25.MC.M3": "E",
        "SP.DYN.LE00.IN": "S", "SE.TER.ENRR": "S"
    })

    readme_df = pd.DataFrame([
        ["objectif", "Construire un panel pays-année pour l'analyse des notations souveraines et facteurs ESG/gouvernance."],
        ["date_generation", datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")],
        ["sources", "World Bank WDI/WGI APIs; ratingshistory.info CSV publics"],
        ["limites", "Données publiques hétérogènes; trous de couverture; ratings SP/Moodys/Fitch selon disponibilité."],
    ], columns=["champ", "valeur"])

    notes_methodo = pd.DataFrame([
        ["Choix de construction", "Panel équilibré iso3 x année puis fusion à gauche des données disponibles."],
        ["Mapping ratings", "Mapping ordinal de AAA/Aaa=21 à D/SD/RD=0 selon consigne."],
        ["Limites", "Disponibilité variable selon indicateur/pays/année; éventuels changements méthodologiques des sources."],
        ["Interprétation", "Les analyses issues de cette base relèvent de corrélations et non d'une causalité stricte."],
    ], columns=["theme", "detail"])

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        readme_df.to_excel(writer, sheet_name="README", index=False)
        sources_df.to_excel(writer, sheet_name="Sources", index=False)
        countries.to_excel(writer, sheet_name="Countries", index=False)
        ratings_raw.to_excel(writer, sheet_name="Ratings_raw", index=False)
        ratings_annual.to_excel(writer, sheet_name="Ratings_annual", index=False)
        wdi_raw.to_excel(writer, sheet_name="WDI_raw", index=False)
        wgi_raw.to_excel(writer, sheet_name="WGI_raw", index=False)
        esg_public_raw.to_excel(writer, sheet_name="ESG_public_raw", index=False)
        panel.to_excel(writer, sheet_name="Panel_final", index=False)
        descriptive_stats.to_excel(writer, sheet_name="Descriptive_stats", index=False)
        missing_values.to_excel(writer, sheet_name="Missing_values", index=False)
        regression_ready.to_excel(writer, sheet_name="Regression_ready", index=False)
        notes_methodo.to_excel(writer, sheet_name="Notes_methodologiques", index=False)

    logging.info("Done. Outputs created.")


if __name__ == "__main__":
    main()
