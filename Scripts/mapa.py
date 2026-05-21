import pandas as pd
import numpy as np
import re
import unicodedata

# =========================
# 1) Carga
# =========================
archivo = "./Data/GlobalDataUpdated-15-12-2025.xlsx"  # <-- cambia esto
df = pd.read_excel(archivo)
df.columns = [c.strip() for c in df.columns]

# =========================
# 2) Helpers de limpieza
# =========================
def strip_accents(s: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(ch)
    )

def clean_text(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    s = re.sub(r"\s+", " ", s)
    s = strip_accents(s).upper()
    if s in {"", "NA", "N/A", "NONE", "-", "--", "NULL"}:
        return np.nan
    return s

def clean_email(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip().lower()
    s = re.sub(r"\s+", "", s)
    if s in {"", "na", "n/a", "none", "-", "--", "null"}:
        return np.nan
    # validación mínima
    if "@" not in s:
        return np.nan
    return s

def clean_phone_presence(x):
    """
    Devuelve True si parece haber un celular usable.
    No intenta estandarizar, solo presencia real.
    """
    if pd.isna(x):
        return False
    s = str(x).strip()
    if s == "":
        return False
    s_low = s.lower()
    if s_low in {"na", "n/a", "none", "-", "--", "null"}:
        return False
    # si solo hay símbolos o muy corto, lo tratamos como ausente
    digits = re.sub(r"\D", "", s)
    return len(digits) >= 7  # umbral razonable LATAM

# =========================
# 3) Columnas
# =========================
pais_col = "PAÍS"
industria_col = "INDUSTRIA"
razon_social_col = "RAZON SOCIAL"
nombre_comercial_col = "NOMBRE COMERCIAL EMPRESA"
interlocutor_col = "INTERLOCUTOR 1"
email_corp_col = "CORREO ELECTRÓNICO CORPORATIVO"
email_personal_col = "CORREO ELECTRÓNICO PERSONAL"
celular_col = "CELULAR"

# Normalizar textos principales
for c in [pais_col, industria_col, razon_social_col, nombre_comercial_col, interlocutor_col]:
    if c in df.columns:
        df[c] = df[c].apply(clean_text)

for c in [email_corp_col, email_personal_col]:
    if c in df.columns:
        df[c] = df[c].apply(clean_email)

# =========================
# 4) Normalización de país LATAM + Caribe
# =========================
LATAM_CARIBE = {
    # LATAM
    "ARGENTINA","BOLIVIA","BRASIL","CHILE","COLOMBIA","COSTA RICA","CUBA",
    "ECUADOR","EL SALVADOR","GUATEMALA","HONDURAS","MEXICO","NICARAGUA",
    "PANAMA","PARAGUAY","PERU","REPUBLICA DOMINICANA","URUGUAY","VENEZUELA",
    # Caribe / región ampliada
    "BAHAMAS","BARBADOS","HAITI","JAMAICA","TRINIDAD Y TOBAGO","TRINIDAD AND TOBAGO",
    "ANTIGUA Y BARBUDA","ANTIGUA AND BARBUDA","DOMINICA","GRANADA","GRENADA",
    "SAN CRISTOBAL Y NIEVES","SAINT KITTS AND NEVIS","SAINT KITTS & NEVIS",
    "SANTA LUCIA","SAINT LUCIA","SAN VICENTE Y LAS GRANADINAS","SAINT VINCENT AND THE GRENADINES",
    "GUYANA","SURINAME","BELICE","BELIZE",
    "PUERTO RICO"
}

# Mapeo de variantes comunes a una forma estándar
country_map = {
    "MEXICO D F": "MEXICO",
    "ESTADOS UNIDOS MEXICANOS": "MEXICO",
    "REP DOMINICANA": "REPUBLICA DOMINICANA",
    "R. DOMINICANA": "REPUBLICA DOMINICANA",
    "DOMINICAN REPUBLIC": "REPUBLICA DOMINICANA",
    "TRINIDAD & TOBAGO": "TRINIDAD Y TOBAGO",
    "TRINIDAD AND TOBAGO": "TRINIDAD Y TOBAGO",
    "SAINT KITTS & NEVIS": "SAN CRISTOBAL Y NIEVES",
    "SAINT KITTS AND NEVIS": "SAN CRISTOBAL Y NIEVES",
    "SAINT LUCIA": "SANTA LUCIA",
    "SAINT VINCENT AND THE GRENADINES": "SAN VICENTE Y LAS GRANADINAS",
    "ANTIGUA AND BARBUDA": "ANTIGUA Y BARBUDA",
    "GRENADA": "GRANADA",
    "BELIZE": "BELICE",
}

df[pais_col] = df[pais_col].replace(country_map)

# Opcional: marcar fuera de LATAM+Caribe (por si aparece algo raro)
df["PAIS_REGION_OK"] = df[pais_col].isin(LATAM_CARIBE)

# =========================
# 5) Keys rápidas (empresa y contacto)
# =========================
# Empresa key: razón social si existe; si no, nombre comercial. Se concatena con país.
df["empresa_base"] = df[razon_social_col].fillna(df[nombre_comercial_col])
df["empresa_key"] = (df["empresa_base"].fillna("SIN_EMPRESA") + " | " + df[pais_col].fillna("SIN_PAIS"))

# Contacto key: prioridad email; si no hay email, usamos nombre+empresa.
df["email_key"] = df[email_corp_col].fillna(df[email_personal_col])
df["contacto_key"] = df["email_key"].fillna(
    (df[interlocutor_col].fillna("SIN_CONTACTO") + " | " + df["empresa_key"])
)

# Flags de contacto (presencia, no cantidad de filas)
df["has_email"] = df["email_key"].notna()
df["has_celular"] = df[celular_col].apply(clean_phone_presence) if celular_col in df.columns else False

# =========================
# 6) Función: resumen por dimensión (País o Industria)
# =========================
def resumen_por_dimension(dim_col, top_n=50):
    temp = df.copy()
    temp[dim_col] = temp[dim_col].fillna("SIN_DATO")

    # A nivel "contacto dentro de la dimensión" nos quedamos con si ese contacto tiene email/celular
    contact_dim = (
        temp.groupby([dim_col, "contacto_key"], dropna=False)
            .agg(
                has_email=("has_email", "max"),
                has_celular=("has_celular", "max"),
            )
            .reset_index()
    )

    # Conteos de contactos únicos y contactos con email/celular (únicos)
    contactos_agg = (
        contact_dim.groupby(dim_col, dropna=False)
            .agg(
                interlocutores_unicos=("contacto_key", "nunique"),
                interlocutores_con_email=("has_email", "sum"),
                interlocutores_con_celular=("has_celular", "sum"),
            )
            .reset_index()
    )

    # Empresas únicas por dimensión (considerando todas las filas)
    empresas_agg = (
        temp.groupby(dim_col, dropna=False)
            .agg(
                empresas_unicas=("empresa_key", pd.Series.nunique),
                registros=("empresa_key", "size"),
                pais_region_ok=("PAIS_REGION_OK", "mean") if dim_col == pais_col else ("PAIS_REGION_OK", "mean"),
            )
            .reset_index()
    )

    out = contactos_agg.merge(empresas_agg, on=dim_col, how="left")

    # Indicadores útiles para lectura rápida
    out["%interlocutores_con_email"] = (out["interlocutores_con_email"] / out["interlocutores_unicos"] * 100).round(2)
    out["%interlocutores_con_celular"] = (out["interlocutores_con_celular"] / out["interlocutores_unicos"] * 100).round(2)

    # Si dim = País, muestra qué tan “OK LATAM+Caribe” es el valor (promedio de bool en filas)
    if dim_col == pais_col:
        out["pais_es_latam_caribe_en_filas_%"] = (out["pais_region_ok"] * 100).round(2)
    out = out.drop(columns=["pais_region_ok"])

    # Orden y top
    out = out.sort_values(["empresas_unicas", "interlocutores_unicos"], ascending=False)

    if top_n is not None and len(out) > top_n:
        top = out.head(top_n).copy()
        otros = pd.DataFrame({
            dim_col: ["OTROS"],
            "interlocutores_unicos": [out.iloc[top_n:]["interlocutores_unicos"].sum()],
            "interlocutores_con_email": [out.iloc[top_n:]["interlocutores_con_email"].sum()],
            "interlocutores_con_celular": [out.iloc[top_n:]["interlocutores_con_celular"].sum()],
            "empresas_unicas": [out.iloc[top_n:]["empresas_unicas"].sum()],
            "registros": [out.iloc[top_n:]["registros"].sum()],
            "%interlocutores_con_email": [np.nan],
            "%interlocutores_con_celular": [np.nan],
        })
        if dim_col == pais_col and "pais_es_latam_caribe_en_filas_%" in out.columns:
            otros["pais_es_latam_caribe_en_filas_%"] = [np.nan]
        out = pd.concat([top, otros], ignore_index=True)

    return out

mapa_pais = resumen_por_dimension(pais_col, top_n=50)
mapa_industria = resumen_por_dimension(industria_col, top_n=50)

# =========================
# 7) Resumen ejecutivo
# =========================
resumen = pd.DataFrame({
    "Metrica": [
        "Registros totales (filas = contactos)",
        "Interlocutores únicos (aprox.)",
        "Empresas únicas (aprox.)",
        "% filas con País informado",
        "% filas con Industria informada",
        "% filas con Email (corp o personal)",
        "% filas con Celular válido",
    ],
    "Valor": [
        len(df),
        df["contacto_key"].nunique(),
        df["empresa_key"].nunique(),
        round(df[pais_col].notna().mean() * 100, 2),
        round(df[industria_col].notna().mean() * 100, 2),
        round(df["has_email"].mean() * 100, 2),
        round(df["has_celular"].mean() * 100, 2),
    ]
})

# =========================
# 8) Exportar entregable
# =========================
salida = "MAPA_DB_IBP.xlsx"
with pd.ExcelWriter(salida, engine="openpyxl") as writer:
    resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
    mapa_pais.to_excel(writer, sheet_name="POR_PAIS", index=False)
    mapa_industria.to_excel(writer, sheet_name="POR_INDUSTRIA", index=False)

print("Listo. Archivo generado:", salida)







# =========================
# 9) Cruce País x Industria
#    - País: TODOS
#    - Industrias: EXACTAMENTE las que quedaron en la hoja POR_INDUSTRIA
# =========================

# 9.1) Lista de industrias "oficiales" (tal cual quedaron en el dataframe mapa_industria)
#      Nota: como mapa_industria ya tiene top_n=50, normalmente serán 50 + "OTROS" = 51.
industrias_oficiales = mapa_industria[industria_col].dropna().astype(str).tolist()
set_industrias_oficiales = set(industrias_oficiales)

# 9.2) Preparar dataframe para el cruce (llenar nulos)
df_cruce = df.copy()
df_cruce[pais_col] = df_cruce[pais_col].fillna("SIN_DATO")
df_cruce[industria_col] = df_cruce[industria_col].fillna("SIN_DATO")

# 9.3) Reglas para que el cruce respete exactamente POR_INDUSTRIA
# - Si "OTROS" existe en POR_INDUSTRIA, todo lo que NO esté en industrias_oficiales se agrupa como "OTROS"
# - Si "OTROS" NO existe, se filtra estrictamente a solo industrias_oficiales
tiene_otros = "OTROS" in set_industrias_oficiales

if tiene_otros:
    df_cruce[industria_col] = df_cruce[industria_col].where(
        df_cruce[industria_col].isin(set_industrias_oficiales),
        other="OTROS"
    )
else:
    df_cruce = df_cruce[df_cruce[industria_col].isin(set_industrias_oficiales)].copy()

# 9.4) Cruce: Empresas únicas por (País, Industria)
cruce_empresas = pd.pivot_table(
    df_cruce,
    index=pais_col,
    columns=industria_col,
    values="empresa_key",
    aggfunc=pd.Series.nunique,
    fill_value=0
)

# 9.5) Cruce: Interlocutores únicos por (País, Industria)
cruce_interlocutores = pd.pivot_table(
    df_cruce,
    index=pais_col,
    columns=industria_col,
    values="contacto_key",
    aggfunc=pd.Series.nunique,
    fill_value=0
)

# 9.6) Forzar columnas EXACTAS y en el MISMO ORDEN que POR_INDUSTRIA
#      (si alguna industria no aparece en el cruce, se crea con ceros)
for ind in industrias_oficiales:
    if ind not in cruce_empresas.columns:
        cruce_empresas[ind] = 0
    if ind not in cruce_interlocutores.columns:
        cruce_interlocutores[ind] = 0

cruce_empresas = cruce_empresas[industrias_oficiales]
cruce_interlocutores = cruce_interlocutores[industrias_oficiales]

# 9.7) (Opcional recomendado) Ordenar países por total para lectura
cruce_empresas = cruce_empresas.loc[cruce_empresas.sum(axis=1).sort_values(ascending=False).index]
cruce_interlocutores = cruce_interlocutores.loc[cruce_interlocutores.sum(axis=1).sort_values(ascending=False).index]

# =========================
# 10) Exportar de nuevo incluyendo las hojas CRUCE_*
#     (En vez de append, re-escribimos el archivo completo para evitar problemas de modo "a")
# =========================
salida = "MAPA_DB_IBP.xlsx"
with pd.ExcelWriter(salida, engine="openpyxl") as writer:
    resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
    mapa_pais.to_excel(writer, sheet_name="POR_PAIS", index=False)
    mapa_industria.to_excel(writer, sheet_name="POR_INDUSTRIA", index=False)
    cruce_empresas.to_excel(writer, sheet_name="CRUCE_EMPRESAS")
    cruce_interlocutores.to_excel(writer, sheet_name="CRUCE_INTERLOCUTORES")

print("Listo. Archivo generado con cruces:", salida)




