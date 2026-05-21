import os
import re
import unicodedata
import pandas as pd

# ============================================================
# CONFIGURACIÓN
# ============================================================

BASE_DIR = "/Users/elizabethmesa/ibp/Cruces"

# Cambia estos nombres por los reales de tus archivos
MAIN_FILE = os.path.join(BASE_DIR, "base_principal_wolk.xlsx")
COMPANY_EXCLUSION_FILE = os.path.join(BASE_DIR, "empresas_a_eliminar_wolk.xlsx")
CALLS_FILE = os.path.join(BASE_DIR, "reporte_llamadas_wolk.xlsx")

# Si conoces el nombre de la hoja, puedes ponerlo aquí.
# Si no, deja None para que lea la primera hoja.
MAIN_SHEET = None
COMPANY_EXCLUSION_SHEET = None
CALLS_SHEET = None

OUTPUT_CLEAN = os.path.join(BASE_DIR, "base_limpia_wolk.xlsx")
OUTPUT_REMOVED = os.path.join(BASE_DIR, "registros_eliminados_wolk.xlsx")
OUTPUT_SUMMARY = os.path.join(BASE_DIR, "resumen_limpieza_wolk.xlsx")

# Columna del archivo de exclusión de empresas
# Si no estás segura, el script intentará detectar la primera columna.
COMPANY_EXCLUSION_COLUMN = None

# Tags que implican descarte
DISCARD_TAGS = {
    "DESCARTADO",
    "DATOS INVALIDOS",
    "COMPETENCIA",
    "NUMERO EQUIVOCADO",
    "NUMERO ERRADO",
    "SOLICITA NO VOLVER A CONTACTAR",
    "YA NO TRABAJA EN ESTA EMPRESA",
}

# Columnas de teléfonos de la base principal
PHONE_COLUMNS_MAIN = [
    "TELÉFONO DE EMPRESA",
    "CELULAR",
    "TELEFONO 1",
    "TELEFONO 2",
]

# Columnas clave
COMPANY_NAME_COL = "NOMBRE COMERCIAL EMPRESA"
INTERLOCUTOR_COL = "INTERLOCUTOR 1"

# Columnas del reporte de llamadas
CALL_CONTACT_NAME_COL = "contact_name"
CALL_CONTACT_NUMBER_COL = "contact_number"
CALL_TAG1_COL = "tag 1"
CALL_TAG2_COL = "tag 2"

# Si quieres una comparación de teléfonos más conservadora, deja solo [10]
PHONE_SUFFIX_LENGTHS = [10, 8]


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def normalize_text(value):
    """
    Normaliza texto:
    - convierte a string
    - quita tildes
    - pasa a mayúsculas
    - elimina espacios extra
    - elimina signos raros
    """
    if pd.isna(value):
        return ""

    value = str(value).strip().upper()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(c for c in value if not unicodedata.combining(c))
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_phone(value):
    """
    Normaliza teléfonos:
    - convierte a string
    - quita .0 típico de Excel
    - deja solo dígitos
    """
    if pd.isna(value):
        return ""

    value = str(value).strip()

    # Eliminar .0 cuando Excel interpreta números como float
    if re.match(r"^\d+\.0$", value):
        value = value[:-2]

    # Dejar solo dígitos
    value = re.sub(r"\D", "", value)
    return value


def get_phone_suffixes(phone, lengths=None):
    """
    Devuelve variantes del teléfono para comparación por sufijo.
    Ej: si phone = 573001234567 -> puede devolver 3001234567 y 34567 etc.
    """
    if not phone:
        return set()

    if lengths is None:
        lengths = [10]

    suffixes = {phone}
    for length in lengths:
        if len(phone) >= length:
            suffixes.add(phone[-length:])
    return suffixes


def read_excel_flexible(path, sheet_name=None):
    """
    Lee un Excel. Si sheet_name es None, toma la primera hoja.
    """
    if sheet_name is None:
        return pd.read_excel(path)
    return pd.read_excel(path, sheet_name=sheet_name)


def detect_first_column(df):
    """
    Devuelve el nombre de la primera columna no vacía.
    """
    if df.empty:
        raise ValueError("El archivo está vacío.")
    return df.columns[0]


# ============================================================
# CARGA DE ARCHIVOS
# ============================================================

print("Leyendo archivos...")

df_main = read_excel_flexible(MAIN_FILE, MAIN_SHEET)
df_companies = read_excel_flexible(COMPANY_EXCLUSION_FILE, COMPANY_EXCLUSION_SHEET)
df_calls = read_excel_flexible(CALLS_FILE, CALLS_SHEET)

print(f"Base principal: {len(df_main)} registros")
print(f"Lista de empresas: {len(df_companies)} registros")
print(f"Reporte de llamadas: {len(df_calls)} registros")


# ============================================================
# VALIDACIÓN DE COLUMNAS
# ============================================================

required_main_cols = [COMPANY_NAME_COL, INTERLOCUTOR_COL] + PHONE_COLUMNS_MAIN
required_calls_cols = [CALL_CONTACT_NAME_COL, CALL_CONTACT_NUMBER_COL, CALL_TAG1_COL, CALL_TAG2_COL]

missing_main = [c for c in required_main_cols if c not in df_main.columns]
missing_calls = [c for c in required_calls_cols if c not in df_calls.columns]

if missing_main:
    raise ValueError(f"Faltan columnas en la base principal: {missing_main}")

if missing_calls:
    raise ValueError(f"Faltan columnas en el reporte de llamadas: {missing_calls}")

if COMPANY_EXCLUSION_COLUMN is None:
    COMPANY_EXCLUSION_COLUMN = detect_first_column(df_companies)

if COMPANY_EXCLUSION_COLUMN not in df_companies.columns:
    raise ValueError(
        f"La columna '{COMPANY_EXCLUSION_COLUMN}' no existe en el archivo de exclusión de empresas."
    )


# ============================================================
# NORMALIZACIÓN DE CAMPOS EN BASE PRINCIPAL
# ============================================================

print("Normalizando base principal...")

df_main = df_main.copy()
df_main["_row_id"] = range(1, len(df_main) + 1)

df_main["_company_norm"] = df_main[COMPANY_NAME_COL].apply(normalize_text)
df_main["_interlocutor_norm"] = df_main[INTERLOCUTOR_COL].apply(normalize_text)

for col in PHONE_COLUMNS_MAIN:
    df_main[f"_{col}_norm"] = df_main[col].apply(normalize_phone)

# Unimos todos los teléfonos normalizados de la fila en un set
def collect_main_phone_keys(row):
    keys = set()
    for col in PHONE_COLUMNS_MAIN:
        phone = row.get(f"_{col}_norm", "")
        keys.update(get_phone_suffixes(phone, PHONE_SUFFIX_LENGTHS))
    keys.discard("")
    return keys

df_main["_phone_keys"] = df_main.apply(collect_main_phone_keys, axis=1)


# ============================================================
# 1) ELIMINACIÓN POR LISTA DE EMPRESAS
# ============================================================

print("Procesando exclusión por empresas...")

excluded_companies = (
    df_companies[COMPANY_EXCLUSION_COLUMN]
    .dropna()
    .astype(str)
    .map(normalize_text)
)

excluded_companies_set = set(x for x in excluded_companies if x)

mask_company_exclusion = df_main["_company_norm"].isin(excluded_companies_set)

removed_by_company = df_main[mask_company_exclusion].copy()
removed_by_company["_reason"] = "Empresa incluida en lista de exclusión"
removed_by_company["_source"] = "Lista empresas"

df_remaining = df_main[~mask_company_exclusion].copy()

print(f"Eliminados por empresa: {len(removed_by_company)}")
print(f"Registros restantes tras exclusión por empresa: {len(df_remaining)}")


# ============================================================
# 2) FILTRAR LLAMADAS DESCARTABLES
# ============================================================

print("Procesando reporte de llamadas...")

df_calls = df_calls.copy()
df_calls["_tag1_norm"] = df_calls[CALL_TAG1_COL].apply(normalize_text)
df_calls["_tag2_norm"] = df_calls[CALL_TAG2_COL].apply(normalize_text)

mask_discard_calls = (
    df_calls["_tag1_norm"].isin(DISCARD_TAGS) |
    df_calls["_tag2_norm"].isin(DISCARD_TAGS)
)

df_calls_discard = df_calls[mask_discard_calls].copy()

print(f"Llamadas con tags de descarte: {len(df_calls_discard)}")

df_calls_discard["_contact_name_norm"] = df_calls_discard[CALL_CONTACT_NAME_COL].apply(normalize_text)
df_calls_discard["_contact_number_norm"] = df_calls_discard[CALL_CONTACT_NUMBER_COL].apply(normalize_phone)

# Conjuntos de exclusión por nombre y por teléfono
excluded_contact_names = set(
    x for x in df_calls_discard["_contact_name_norm"].dropna().tolist() if x
)

excluded_phone_keys = set()
for phone in df_calls_discard["_contact_number_norm"].dropna().tolist():
    excluded_phone_keys.update(get_phone_suffixes(phone, PHONE_SUFFIX_LENGTHS))

excluded_phone_keys.discard("")

print(f"Nombres de contacto para excluir: {len(excluded_contact_names)}")
print(f"Claves de teléfono para excluir: {len(excluded_phone_keys)}")


# ============================================================
# 3) ELIMINACIÓN POR NOMBRE Y/O TELÉFONO
# ============================================================

print("Aplicando exclusión por llamadas...")

def match_by_phone(phone_keys, excluded_keys):
    if not phone_keys:
        return False
    return len(phone_keys.intersection(excluded_keys)) > 0

mask_name = df_remaining["_interlocutor_norm"].isin(excluded_contact_names)
mask_phone = df_remaining["_phone_keys"].apply(lambda x: match_by_phone(x, excluded_phone_keys))
mask_calls_exclusion = mask_name | mask_phone

removed_by_calls = df_remaining[mask_calls_exclusion].copy()

def build_reason(row):
    reasons = []
    if row["_interlocutor_norm"] in excluded_contact_names:
        reasons.append("Coincidencia por nombre de interlocutor")
    if match_by_phone(row["_phone_keys"], excluded_phone_keys):
        reasons.append("Coincidencia por teléfono")
    return " + ".join(reasons)

removed_by_calls["_reason"] = removed_by_calls.apply(build_reason, axis=1)
removed_by_calls["_source"] = "Reporte llamadas"

df_clean = df_remaining[~mask_calls_exclusion].copy()

print(f"Eliminados por llamadas: {len(removed_by_calls)}")
print(f"Registros finales en base limpia: {len(df_clean)}")


# ============================================================
# CONSOLIDAR ELIMINADOS
# ============================================================

removed_all = pd.concat([removed_by_company, removed_by_calls], ignore_index=True)

# Eliminar columnas auxiliares antes de exportar
helper_cols = [c for c in removed_all.columns if c.startswith("_")]
removed_export = removed_all.drop(columns=helper_cols, errors="ignore")

helper_cols_clean = [c for c in df_clean.columns if c.startswith("_")]
clean_export = df_clean.drop(columns=helper_cols_clean, errors="ignore")


# ============================================================
# RESUMEN
# ============================================================

summary = pd.DataFrame({
    "métrica": [
        "Registros originales",
        "Eliminados por lista de empresas",
        "Eliminados por reporte de llamadas",
        "Total eliminados",
        "Registros finales"
    ],
    "valor": [
        len(df_main),
        len(removed_by_company),
        len(removed_by_calls),
        len(removed_all),
        len(df_clean)
    ]
})


# ============================================================
# EXPORTACIÓN
# ============================================================

print("Guardando resultados...")

with pd.ExcelWriter(OUTPUT_CLEAN, engine="openpyxl") as writer:
    clean_export.to_excel(writer, sheet_name="base_limpia", index=False)

with pd.ExcelWriter(OUTPUT_REMOVED, engine="openpyxl") as writer:
    removed_export.to_excel(writer, sheet_name="eliminados", index=False)

with pd.ExcelWriter(OUTPUT_SUMMARY, engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="resumen", index=False)

print("Proceso finalizado.")
print(f"Base limpia: {OUTPUT_CLEAN}")
print(f"Eliminados: {OUTPUT_REMOVED}")
print(f"Resumen: {OUTPUT_SUMMARY}")