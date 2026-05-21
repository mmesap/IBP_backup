import re
from pathlib import Path
from typing import Dict, List, Optional

import numpy as np
import pandas as pd


# ============================================================
# CONFIGURACIÓN
# ============================================================

# Ruta absoluta del archivo fuente
INPUT_FILE = Path("/Users/elizabethmesa/ibp/Data/Base global/GlobalDataUpdated-30-3-2026.xlsx")

# Carpeta donde está guardado este script
SCRIPT_DIR = Path(__file__).resolve().parent

# Carpeta de salida dentro del mismo directorio del script
OUTPUT_DIR = SCRIPT_DIR / "diagnostico_inicial"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Nombre opcional de hoja. Si se deja en None, usa la primera hoja.
SHEET_NAME = None


# ============================================================
# UTILIDADES
# ============================================================

def normalize_colname(name: str) -> str:
    """Normalize column names for internal matching only."""
    if name is None:
        return ""
    text = str(name).strip().lower()
    replacements = {
        "á": "a",
        "é": "e",
        "í": "i",
        "ó": "o",
        "ú": "u",
        "ü": "u",
        "ñ": "n",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"\s+", " ", text)
    return text


def build_column_map(columns: List[str]) -> Dict[str, str]:
    """Map normalized column names to original column names."""
    return {normalize_colname(col): col for col in columns}


def find_first_matching_column(
    colmap: Dict[str, str],
    candidates: List[str]
) -> Optional[str]:
    """Return first matching original column name from candidate aliases."""
    for cand in candidates:
        key = normalize_colname(cand)
        if key in colmap:
            return colmap[key]
    return None


def as_clean_string(series: pd.Series) -> pd.Series:
    """Convert series to cleaned string values preserving nulls."""
    s = series.copy()
    s = s.astype("string")
    s = s.str.strip()
    s = s.replace(
        {
            "": pd.NA,
            "nan": pd.NA,
            "None": pd.NA,
            "none": pd.NA,
            "NaN": pd.NA,
            "<NA>": pd.NA,
        }
    )
    return s


def normalize_text_basic(series: pd.Series) -> pd.Series:
    """Basic normalization for duplicate detection."""
    s = as_clean_string(series).str.upper()
    s = s.str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("utf-8")
    s = s.str.replace(r"\s+", " ", regex=True)
    s = s.str.replace(r"[^\w\s&/@.-]", "", regex=True)
    s = s.str.strip()
    return s


def normalize_email(series: pd.Series) -> pd.Series:
    s = as_clean_string(series).str.lower()
    s = s.str.replace(r"\s+", "", regex=True)
    return s


def normalize_linkedin(series: pd.Series) -> pd.Series:
    s = as_clean_string(series).str.lower()
    s = s.str.replace(r"\s+", "", regex=True)
    s = s.str.replace(r"^https?://", "", regex=True)
    s = s.str.replace(r"^www\.", "", regex=True)
    s = s.str.rstrip("/")
    return s


def normalize_phone(series: pd.Series) -> pd.Series:
    s = as_clean_string(series)
    s = s.str.replace(r"\.0$", "", regex=True)
    s = s.str.replace(r"[^\d]", "", regex=True)
    s = s.replace("", pd.NA)
    return s


def valid_email(series: pd.Series) -> pd.Series:
    s = normalize_email(series)
    return s.str.match(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$", na=False)


def valid_linkedin(series: pd.Series) -> pd.Series:
    s = normalize_linkedin(series)
    return s.str.contains("linkedin.com/", na=False)


def valid_phone(series: pd.Series) -> pd.Series:
    s = normalize_phone(series)
    return s.str.match(r"^\d{7,15}$", na=False)


def count_duplicate_rows_by_keys(df: pd.DataFrame, keys: List[str]) -> Dict[str, int]:
    """Return duplicate counts for a given list of key columns."""
    subset = df[keys].copy()
    subset = subset.dropna(how="any")
    if subset.empty:
        return {
            "rows_evaluated": 0,
            "duplicate_rows": 0,
            "duplicate_groups": 0,
        }

    dup_mask = subset.duplicated(keep=False)
    duplicate_rows = int(dup_mask.sum())
    duplicate_groups = int(subset[dup_mask].drop_duplicates().shape[0])

    return {
        "rows_evaluated": int(subset.shape[0]),
        "duplicate_rows": duplicate_rows,
        "duplicate_groups": duplicate_groups,
    }


# ============================================================
# CARGA
# ============================================================

if not INPUT_FILE.exists():
    raise FileNotFoundError(f"No se encontró el archivo de entrada: {INPUT_FILE}")

if SHEET_NAME is None:
    df = pd.read_excel(INPUT_FILE)
else:
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

original_columns = list(df.columns)
colmap = build_column_map(original_columns)

# ============================================================
# ALIAS DE COLUMNAS
# ============================================================

company_col = find_first_matching_column(
    colmap,
    [
        "nombre comercial empresa",
        "nombre comercial",
        "empresa",
        "company",
    ],
)

legal_name_col = find_first_matching_column(
    colmap,
    [
        "razon social",
        "razón social",
    ],
)

country_col = find_first_matching_column(
    colmap,
    [
        "pais",
        "país",
    ],
)

state_col = find_first_matching_column(
    colmap,
    [
        "departamento / estado",
        "departamento/estado",
        "departamento",
        "estado",
    ],
)

city_col = find_first_matching_column(
    colmap,
    [
        "ciudad / provincias",
        "ciudad/provincias",
        "ciudad",
        "provincia",
    ],
)

industry_col = find_first_matching_column(
    colmap,
    [
        "industria",
    ],
)

sector_col = find_first_matching_column(
    colmap,
    [
        "sector",
    ],
)

contact_name_col = find_first_matching_column(
    colmap,
    [
        "interlocutor 1",
        "nombre contacto",
        "contacto",
    ],
)

position_col = find_first_matching_column(
    colmap,
    [
        "cargo",
    ],
)

area_col = find_first_matching_column(
    colmap,
    [
        "area",
        "área",
    ],
)

management_col = find_first_matching_column(
    colmap,
    [
        "management level",
        "managementlevel",
    ],
)

corp_email_col = find_first_matching_column(
    colmap,
    [
        "correo electronico corporativo",
        "correo electrónico corporativo",
        "email corporativo",
    ],
)

personal_email_col = find_first_matching_column(
    colmap,
    [
        "correo electronico personal",
        "correo electrónico personal",
        "email personal",
    ],
)

linkedin_company_col = find_first_matching_column(
    colmap,
    [
        "url linkedin empresa",
        "url linkedin compania",
        "linkedin empresa",
    ],
)

linkedin_contact_col = find_first_matching_column(
    colmap,
    [
        "url linkedin 1",
        "url linkedin contacto",
        "linkedin contacto",
    ],
)

cellphone_col = find_first_matching_column(
    colmap,
    [
        "celular",
    ],
)

phone1_col = find_first_matching_column(
    colmap,
    [
        "telefono 1",
        "teléfono 1",
    ],
)

phone2_col = find_first_matching_column(
    colmap,
    [
        "telefono 2",
        "teléfono 2",
    ],
)

company_phone_col = find_first_matching_column(
    colmap,
    [
        "telefono de empresa",
        "teléfono de empresa",
    ],
)

campaign_col = find_first_matching_column(
    colmap,
    [
        "campana",
        "campaña",
    ],
)

# ============================================================
# ESTADÍSTICAS GENERALES
# ============================================================

general_stats = {
    "input_file": str(INPUT_FILE),
    "output_dir": str(OUTPUT_DIR),
    "sheet_used": SHEET_NAME if SHEET_NAME is not None else "Primera hoja",
    "total_rows": int(df.shape[0]),
    "total_columns": int(df.shape[1]),
    "empty_rows": int(df.isna().all(axis=1).sum()),
}

general_stats_df = pd.DataFrame(
    [{"metric": k, "value": v} for k, v in general_stats.items()]
)

# ============================================================
# PERFIL POR COLUMNA
# ============================================================

profile_rows = []
for col in df.columns:
    series = df[col]
    missing = int(series.isna().sum())
    non_missing = int(series.notna().sum())
    unique_non_null = int(series.dropna().astype("string").nunique())
    completeness_pct = round((non_missing / len(df) * 100), 2) if len(df) else 0.0
    sample_values = (
        series.dropna().astype("string").drop_duplicates().head(5).tolist()
    )
    profile_rows.append(
        {
            "column": col,
            "dtype": str(series.dtype),
            "rows": int(len(series)),
            "missing": missing,
            "non_missing": non_missing,
            "completeness_pct": completeness_pct,
            "unique_non_null": unique_non_null,
            "sample_values": " | ".join(map(str, sample_values)),
        }
    )

column_profile_df = pd.DataFrame(profile_rows).sort_values(
    by=["completeness_pct", "column"], ascending=[True, True]
)

# ============================================================
# MÉTRICAS DE NEGOCIO
# ============================================================

metrics_rows = []

def add_metric(name: str, value):
    metrics_rows.append({"metric": name, "value": value})

if company_col:
    add_metric("company_column_used", company_col)
    company_norm = normalize_text_basic(df[company_col])
    add_metric("unique_companies_by_name", int(company_norm.dropna().nunique()))

if country_col:
    add_metric("country_column_used", country_col)

if company_col and country_col:
    company_country = pd.DataFrame(
        {
            "company": normalize_text_basic(df[company_col]),
            "country": normalize_text_basic(df[country_col]),
        }
    ).dropna()
    add_metric(
        "unique_companies_by_name_country",
        int(company_country.drop_duplicates().shape[0]),
    )

if contact_name_col:
    add_metric("contact_column_used", contact_name_col)
    contact_norm = normalize_text_basic(df[contact_name_col])
    add_metric("unique_contacts_by_name", int(contact_norm.dropna().nunique()))

if corp_email_col:
    corp_email_norm = normalize_email(df[corp_email_col])
    add_metric("records_with_corporate_email", int(corp_email_norm.notna().sum()))
    add_metric("unique_corporate_emails", int(corp_email_norm.dropna().nunique()))

if personal_email_col:
    personal_email_norm = normalize_email(df[personal_email_col])
    add_metric("records_with_personal_email", int(personal_email_norm.notna().sum()))
    add_metric("unique_personal_emails", int(personal_email_norm.dropna().nunique()))

if linkedin_contact_col:
    linkedin_contact_norm = normalize_linkedin(df[linkedin_contact_col])
    add_metric("records_with_contact_linkedin", int(linkedin_contact_norm.notna().sum()))
    add_metric("unique_contact_linkedins", int(linkedin_contact_norm.dropna().nunique()))

if linkedin_company_col:
    linkedin_company_norm = normalize_linkedin(df[linkedin_company_col])
    add_metric("records_with_company_linkedin", int(linkedin_company_norm.notna().sum()))
    add_metric("unique_company_linkedins", int(linkedin_company_norm.dropna().nunique()))

phone_presence = 0
for c in [cellphone_col, phone1_col, phone2_col, company_phone_col]:
    if c:
        phone_presence += normalize_phone(df[c]).notna().astype(int)
add_metric("records_with_at_least_one_phone", int((phone_presence > 0).sum()))

business_metrics_df = pd.DataFrame(metrics_rows)

# ============================================================
# CALIDAD DE DATOS
# ============================================================

quality_rows = []

def add_quality_metric(name: str, value):
    quality_rows.append({"metric": name, "value": value})

if corp_email_col:
    corp_email_clean = normalize_email(df[corp_email_col])
    invalid_corp_email = corp_email_clean.notna() & ~valid_email(df[corp_email_col])
    add_quality_metric("invalid_corporate_emails", int(invalid_corp_email.sum()))

if personal_email_col:
    personal_email_clean = normalize_email(df[personal_email_col])
    invalid_personal_email = personal_email_clean.notna() & ~valid_email(df[personal_email_col])
    add_quality_metric("invalid_personal_emails", int(invalid_personal_email.sum()))

if linkedin_contact_col:
    contact_link_clean = normalize_linkedin(df[linkedin_contact_col])
    invalid_contact_link = contact_link_clean.notna() & ~valid_linkedin(df[linkedin_contact_col])
    add_quality_metric("invalid_contact_linkedins", int(invalid_contact_link.sum()))

if linkedin_company_col:
    company_link_clean = normalize_linkedin(df[linkedin_company_col])
    invalid_company_link = company_link_clean.notna() & ~valid_linkedin(df[linkedin_company_col])
    add_quality_metric("invalid_company_linkedins", int(invalid_company_link.sum()))

for label, col in [
    ("invalid_cellphones", cellphone_col),
    ("invalid_phone1", phone1_col),
    ("invalid_phone2", phone2_col),
    ("invalid_company_phone", company_phone_col),
]:
    if col:
        phone_clean = normalize_phone(df[col])
        invalid_phone = phone_clean.notna() & ~valid_phone(df[col])
        add_quality_metric(label, int(invalid_phone.sum()))

quality_metrics_df = pd.DataFrame(quality_rows)

# ============================================================
# DUPLICADOS
# ============================================================

duplicate_rows = []

def add_duplicate_metric(entity: str, key_logic: str, result: Dict[str, int]):
    duplicate_rows.append(
        {
            "entity": entity,
            "key_logic": key_logic,
            **result,
        }
    )

if company_col:
    temp = pd.DataFrame({"company": normalize_text_basic(df[company_col])})
    add_duplicate_metric(
        "company",
        "company_name",
        count_duplicate_rows_by_keys(temp, ["company"]),
    )

if company_col and country_col:
    temp = pd.DataFrame(
        {
            "company": normalize_text_basic(df[company_col]),
            "country": normalize_text_basic(df[country_col]),
        }
    )
    add_duplicate_metric(
        "company",
        "company_name + country",
        count_duplicate_rows_by_keys(temp, ["company", "country"]),
    )

if contact_name_col:
    temp = pd.DataFrame({"contact": normalize_text_basic(df[contact_name_col])})
    add_duplicate_metric(
        "contact",
        "contact_name",
        count_duplicate_rows_by_keys(temp, ["contact"]),
    )

if corp_email_col:
    temp = pd.DataFrame({"corp_email": normalize_email(df[corp_email_col])})
    add_duplicate_metric(
        "contact",
        "corporate_email",
        count_duplicate_rows_by_keys(temp, ["corp_email"]),
    )

if linkedin_contact_col:
    temp = pd.DataFrame({"linkedin_contact": normalize_linkedin(df[linkedin_contact_col])})
    add_duplicate_metric(
        "contact",
        "contact_linkedin",
        count_duplicate_rows_by_keys(temp, ["linkedin_contact"]),
    )

if company_col and contact_name_col:
    temp = pd.DataFrame(
        {
            "company": normalize_text_basic(df[company_col]),
            "contact": normalize_text_basic(df[contact_name_col]),
        }
    )
    add_duplicate_metric(
        "company_contact",
        "company_name + contact_name",
        count_duplicate_rows_by_keys(temp, ["company", "contact"]),
    )

duplicate_metrics_df = pd.DataFrame(duplicate_rows)

# ============================================================
# TOPS CATEGÓRICOS
# ============================================================

def top_values(df: pd.DataFrame, col: Optional[str], label: str, top_n: int = 20) -> pd.DataFrame:
    if not col:
        return pd.DataFrame(columns=["dimension", "value", "count"])
    s = as_clean_string(df[col])
    vc = s.value_counts(dropna=True).head(top_n).reset_index()
    vc.columns = ["value", "count"]
    vc.insert(0, "dimension", label)
    return vc

tops_frames = [
    top_values(df, country_col, "country"),
    top_values(df, city_col, "city"),
    top_values(df, industry_col, "industry"),
    top_values(df, sector_col, "sector"),
    top_values(df, management_col, "management_level"),
    top_values(df, area_col, "area"),
    top_values(df, campaign_col, "campaign"),
]

tops_df = pd.concat([frame for frame in tops_frames if not frame.empty], ignore_index=True)

# ============================================================
# MUESTRAS DE REGISTROS PROBLEMÁTICOS
# ============================================================

samples = {}

if corp_email_col:
    mask = normalize_email(df[corp_email_col]).notna() & ~valid_email(df[corp_email_col])
    samples["invalid_corporate_emails"] = df.loc[mask, [corp_email_col]].head(200)

if personal_email_col:
    mask = normalize_email(df[personal_email_col]).notna() & ~valid_email(df[personal_email_col])
    samples["invalid_personal_emails"] = df.loc[mask, [personal_email_col]].head(200)

if linkedin_contact_col:
    mask = normalize_linkedin(df[linkedin_contact_col]).notna() & ~valid_linkedin(df[linkedin_contact_col])
    samples["invalid_contact_linkedins"] = df.loc[mask, [linkedin_contact_col]].head(200)

if company_col:
    temp = pd.DataFrame({"company": normalize_text_basic(df[company_col])})
    dup_mask = temp.duplicated(subset=["company"], keep=False) & temp["company"].notna()
    cols = [c for c in [company_col, country_col, legal_name_col] if c]
    samples["duplicate_companies_by_name"] = df.loc[dup_mask, cols].head(500)

if company_col and country_col:
    temp = pd.DataFrame(
        {
            "company": normalize_text_basic(df[company_col]),
            "country": normalize_text_basic(df[country_col]),
        }
    )
    dup_mask = temp.duplicated(subset=["company", "country"], keep=False)
    dup_mask = dup_mask & temp["company"].notna() & temp["country"].notna()
    cols = [c for c in [company_col, country_col, legal_name_col] if c]
    samples["duplicate_companies_by_name_country"] = df.loc[dup_mask, cols].head(500)

if contact_name_col and company_col:
    temp = pd.DataFrame(
        {
            "company": normalize_text_basic(df[company_col]),
            "contact": normalize_text_basic(df[contact_name_col]),
        }
    )
    dup_mask = temp.duplicated(subset=["company", "contact"], keep=False)
    dup_mask = dup_mask & temp["company"].notna() & temp["contact"].notna()
    cols = [
        c for c in
        [company_col, contact_name_col, corp_email_col, cellphone_col, linkedin_contact_col]
        if c
    ]
    samples["duplicate_company_contact"] = df.loc[dup_mask, cols].head(500)

# ============================================================
# EXPORTACIÓN
# ============================================================

excel_output = OUTPUT_DIR / "diagnostico_resumen.xlsx"
with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
    general_stats_df.to_excel(writer, sheet_name="general_stats", index=False)
    business_metrics_df.to_excel(writer, sheet_name="business_metrics", index=False)
    quality_metrics_df.to_excel(writer, sheet_name="quality_metrics", index=False)
    duplicate_metrics_df.to_excel(writer, sheet_name="duplicate_metrics", index=False)
    column_profile_df.to_excel(writer, sheet_name="column_profile", index=False)
    tops_df.to_excel(writer, sheet_name="tops", index=False)

    for sheet_name, sample_df in samples.items():
        safe_sheet = sheet_name[:31]
        sample_df.to_excel(writer, sheet_name=safe_sheet, index=False)

general_stats_df.to_csv(OUTPUT_DIR / "general_stats.csv", index=False)
business_metrics_df.to_csv(OUTPUT_DIR / "business_metrics.csv", index=False)
quality_metrics_df.to_csv(OUTPUT_DIR / "quality_metrics.csv", index=False)
duplicate_metrics_df.to_csv(OUTPUT_DIR / "duplicate_metrics.csv", index=False)
column_profile_df.to_csv(OUTPUT_DIR / "column_profile.csv", index=False)
tops_df.to_csv(OUTPUT_DIR / "tops.csv", index=False)

summary_txt = OUTPUT_DIR / "README_resultados.txt"
with open(summary_txt, "w", encoding="utf-8") as f:
    f.write("DIAGNÓSTICO INICIAL DE BASE IBP\n")
    f.write("=" * 40 + "\n\n")
    f.write(f"Archivo analizado: {INPUT_FILE}\n")
    f.write(f"Carpeta de salida: {OUTPUT_DIR}\n")
    f.write(f"Filas: {df.shape[0]}\n")
    f.write(f"Columnas: {df.shape[1]}\n\n")
    f.write("Archivos generados:\n")
    f.write("- diagnostico_resumen.xlsx\n")
    f.write("- general_stats.csv\n")
    f.write("- business_metrics.csv\n")
    f.write("- quality_metrics.csv\n")
    f.write("- duplicate_metrics.csv\n")
    f.write("- column_profile.csv\n")
    f.write("- tops.csv\n")

print("Diagnóstico completado correctamente.")
print(f"Archivo fuente: {INPUT_FILE}")
print(f"Resultados guardados en: {OUTPUT_DIR}")
print(f"Excel resumen: {excel_output}")
