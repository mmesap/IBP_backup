# normalizar_industrias.py
# -*- coding: utf-8 -*-

import os
import re
import json
import unicodedata
from datetime import datetime

import pandas as pd

# =========================
# CONFIGURACIÓN
# =========================

# 1) Excel de entrada (tu base con ~43k filas)
INPUT_XLSX = "/Users/elizabethmesa/ibp/ibp/Data/GlobalDataUpdated-15-12-2025.xlsx"

# 2) Nombre de la hoja (None = primera hoja)
SHEET_NAME = "BASE"

# 3) Nombre EXACTO de la columna de industrias en tu Excel
INDUSTRY_COL = "INDUSTRIA"

# 4) Carpeta de salida
OUTPUT_DIR = "/Users/elizabethmesa/ibp/ibp/Normalización Telefonía"

# 5) Archivos de apoyo
# - Diccionario manual persistente (se crea si no existe)
MANUAL_MAP_CSV = os.path.join(OUTPUT_DIR, "industria_manual_map.csv")

# - Pendientes para revisión (se regenera cada corrida)
PENDING_CSV = os.path.join(OUTPUT_DIR, "industrias_pendientes.csv")

# - Mapping final recomendado (manual + aprobadas) (se regenera)
FINAL_MAP_CSV = os.path.join(OUTPUT_DIR, "industria_map_final.csv")

# - Excel final con columna normalizada
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "base_con_industria_normalizada.xlsx")

# 6) Parámetros fuzzy
FUZZY_THRESHOLD = 90  # sugerir solo si >= 90 (recomendado)
MAX_SUGGESTIONS = 5   # cuántas sugerencias mostrar por industria pendiente


# =========================
# UTILIDADES DE NORMALIZACIÓN
# =========================

def normalize_text(s: str) -> str:
    """
    Normaliza texto para comparación:
    - minúsculas
    - sin tildes
    - limpia caracteres raros
    - espacios normalizados
    """
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s).strip().lower()

    # quitar tildes/diacríticos
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))

    # decodificar entidades HTML comunes (&amp; -> &)
    s = s.replace("&amp;", "&")

    # reemplazar separadores raros por espacios
    s = s.replace("_", " ").replace("-", " ").replace("·", " ")

    # mantener letras/números y algunos separadores útiles
    s = re.sub(r"[^a-z0-9\s&/]", " ", s)

    # colapsar espacios
    s = re.sub(r"\s+", " ", s).strip()
    return s


def is_garbage_value(s_clean: str) -> bool:
    """
    Detecta valores que típicamente NO son industria (basura).
    Ajusta según tu caso.
    """
    if not s_clean:
        return True
    garbage = {"0", "unknown", "no registra", "no se encuentra", "sin informacion en linkedin", "var"}
    if s_clean in garbage:
        return True

    # casos donde pegaron descripciones largas tipo "quienes somos..." etc.
    if len(s_clean) >= 120:
        return True

    return False


# =========================
# DICCIONARIO MANUAL (BASE)
# =========================

def ensure_output_dir():
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def load_or_create_manual_map() -> pd.DataFrame:
    """
    Carga un CSV con columnas:
    - industria_clean
    - industria_normalizada

    Si no existe, crea uno con algunas equivalencias típicas y lo guarda.
    """
    ensure_output_dir()

    if os.path.exists(MANUAL_MAP_CSV):
        mm = pd.read_csv(MANUAL_MAP_CSV, dtype=str).fillna("")
        # normalizar columna clave por seguridad
        mm["industria_clean"] = mm["industria_clean"].apply(normalize_text)
        mm = mm[["industria_clean", "industria_normalizada"]].drop_duplicates()
        return mm

    # Plantilla inicial (puedes ampliarla)
    seed_pairs = [
        ("banking", "Banca"),
        ("commercial banking", "Banca"),
        ("investment banking", "Banca de Inversión"),
        ("insurance", "Seguros"),
        ("reinsurance", "Reaseguros"),
        ("telecommunications", "Telecomunicaciones"),
        ("information technology & services", "Servicios y Tecnologías de la Información"),
        ("it services and it consulting", "Servicios y Tecnologías de la Información"),
        ("computer software", "Desarrollo de Software"),
        ("hospital & health care", "Hospitales y Atención Sanitaria"),
        ("health, wellness & fitness", "Sanidad, Bienestar y Ejercicio"),
        ("oil & gas", "Petróleo y Gas"),
        ("utilities", "Servicios Públicos"),
        ("retail", "Venta Minorista"),
        ("wholesale", "Venta al por Mayor"),
        ("higher education", "Educación Superior"),
        ("professional services", "Servicios Profesionales"),
        ("management consulting", "Consultoría de Gestión"),
        ("marketing & advertising", "Marketing y Publicidad"),
        ("logistics & supply chain", "Logística y Cadena de Suministro"),
        ("packaging & containers", "Envases y Embalajes"),
        ("pharmaceuticals", "Productos Farmacéuticos"),
        ("medical devices", "Dispositivos Médicos"),
    ]

    mm = pd.DataFrame(seed_pairs, columns=["industria_clean", "industria_normalizada"])
    mm["industria_clean"] = mm["industria_clean"].apply(normalize_text)
    mm = mm.drop_duplicates()

    mm.to_csv(MANUAL_MAP_CSV, index=False, encoding="utf-8-sig")
    print(f"[OK] Creado diccionario manual inicial: {MANUAL_MAP_CSV}")
    return mm


# =========================
# FUZZY SUGGESTIONS
# =========================

def fuzzy_suggestions(pending_clean: list, targets: list) -> pd.DataFrame:
    """
    Devuelve sugerencias fuzzy para cada industria pendiente.
    Requiere rapidfuzz.
    """
    try:
        from rapidfuzz import process, fuzz
    except ImportError as e:
        raise SystemExit(
            "Falta rapidfuzz. Instala con:\n\n"
            "  pip install rapidfuzz\n"
        ) from e

    rows = []
    for v in pending_clean:
        # retorna lista de tuplas (match, score, idx)
        matches = process.extract(v, targets, scorer=fuzz.token_set_ratio, limit=MAX_SUGGESTIONS)
        # filtrar por threshold para sugerencia “fuerte”
        strong = [m for m in matches if m[1] >= FUZZY_THRESHOLD]

        row = {
            "industria_clean": v,
            "sugerencia_top": strong[0][0] if strong else "",
            "score_top": strong[0][1] if strong else "",
        }

        # columnas extras con top-N (útil para revisión)
        for i, m in enumerate(matches, start=1):
            row[f"cand_{i}"] = m[0]
            row[f"score_{i}"] = m[1]

        rows.append(row)

    return pd.DataFrame(rows)


# =========================
# PROCESO PRINCIPAL
# =========================

def main():
    ensure_output_dir()

    # Leer Excel
    df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET_NAME, dtype=str)
    if INDUSTRY_COL not in df.columns:
        raise SystemExit(
            f"No encuentro la columna '{INDUSTRY_COL}'. Columnas disponibles:\n{list(df.columns)}"
        )

    # Columna original
    df[INDUSTRY_COL] = df[INDUSTRY_COL].fillna("").astype(str)

    # Limpieza para matching
    df["industria_clean"] = df[INDUSTRY_COL].apply(normalize_text)

    # Marcar basura (opcional: la dejamos sin normalizar)
    df["industria_es_basura"] = df["industria_clean"].apply(is_garbage_value)

    # Unicos (para trabajar “inteligente”)
    uniques = (
        df.loc[~df["industria_es_basura"], "industria_clean"]
        .dropna()
        .unique()
        .tolist()
    )

    print(f"[INFO] Filas totales: {len(df):,}")
    print(f"[INFO] Industrias únicas (no basura): {len(uniques):,}")

    # Cargar diccionario manual
    manual_map = load_or_create_manual_map()
    manual_dict = dict(zip(manual_map["industria_clean"], manual_map["industria_normalizada"]))

    # Aplicar mapeo manual
    df["industria_normalizada"] = df["industria_clean"].map(manual_dict)

    # Pendientes (solo las no basura)
    pending_clean = sorted(
        set(df.loc[~df["industria_es_basura"] & df["industria_normalizada"].isna(), "industria_clean"].tolist())
    )

    print(f"[INFO] Pendientes sin resolver (no basura): {len(pending_clean):,}")

    # Targets para fuzzy: usamos los valores normalizados existentes (catálogo)
    catalogo_norm = sorted(set(manual_map["industria_normalizada"].tolist()))
    if catalogo_norm:
        sug = fuzzy_suggestions(pending_clean, catalogo_norm)
        sug.to_csv(PENDING_CSV, index=False, encoding="utf-8-sig")
        print(f"[OK] Generado pendientes con sugerencias: {PENDING_CSV}")
    else:
        # si no hay catálogo, exporta pendientes sin sugerencias
        pd.DataFrame({"industria_clean": pending_clean}).to_csv(PENDING_CSV, index=False, encoding="utf-8-sig")
        print(f"[OK] Generado pendientes (sin sugerencias): {PENDING_CSV}")

    # Exportar mapping final actual (solo lo manual por ahora)
    manual_map.to_csv(FINAL_MAP_CSV, index=False, encoding="utf-8-sig")
    print(f"[OK] Exportado mapping final actual: {FINAL_MAP_CSV}")

    # Exportar Excel con columna normalizada
    # Nota: las filas basura quedan con industria_normalizada vacía (puedes cambiar si quieres)
    df.to_excel(OUTPUT_XLSX, index=False)
    print(f"[OK] Exportado Excel final: {OUTPUT_XLSX}")

    print("\n=== SIGUIENTE PASO RECOMENDADO ===")
    print(f"1) Abre el archivo de pendientes: {PENDING_CSV}")
    print("2) Decide la industria_normalizada correcta para cada industria_clean pendiente.")
    print(f"3) Agrega esas nuevas filas al diccionario manual: {MANUAL_MAP_CSV}")
    print("4) Vuelve a correr este script. Repite hasta que pendientes sea ~0.")


if __name__ == "__main__":
    main()
