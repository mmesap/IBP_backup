import os
import re
import sys
import pandas as pd

# ==========================================
# USO:
# python 02_construir_base_contactos.py 2026-01 GLOBAL
# python 02_construir_base_contactos.py 2026-01 MES
#
# - GLOBAL -> lee Contactabilidad/Data/global/base_global.xlsx
# - MES    -> lee Contactabilidad/out/<MES>/entregas_<MES>_consolidado.xlsx
# ==========================================

if len(sys.argv) != 3:
    print("Uso: python 02_construir_base_contactos.py YYYY-MM (GLOBAL|MES)")
    sys.exit(1)

MES = sys.argv[1]
TIPO = sys.argv[2].upper().strip()

if TIPO not in {"GLOBAL", "MES"}:
    raise ValueError("TIPO debe ser GLOBAL o MES")

if TIPO == "GLOBAL":
    INPUT_PATH = "./Contactabilidad/Data/global/base_global.xlsx"
    OUT_FOLDER = f"./Contactabilidad/out/{MES}/A_historico"
    LABEL = "GLOBAL"
else:
    INPUT_PATH = f"./Contactabilidad/out/{MES}/entregas_{MES}_consolidado.xlsx"
    OUT_FOLDER = f"./Contactabilidad/out/{MES}/B_mes"
    LABEL = f"ENTREGAS_{MES}"

OUT_LONG = f"{OUT_FOLDER}/base_contactos_long_{LABEL}.xlsx"
OUT_CONTACTS = f"{OUT_FOLDER}/base_contactos_{LABEL}.xlsx"

# Teléfonos personales
PHONE_COLS = ["CELULAR", "TELEFONO 1", "TELEFONO 2"]
# Teléfono empresa (se cruza con doble verificación por nombre)
EMP_PHONE_COL = "TELÉFONO DE EMPRESA"

EMAIL_COL = "CORREO ELECTRONICO CORPORATIVO"
NAME_COL  = "INTERLOCUTOR 1"
COMP_COL  = "RAZON SOCIAL"
PAIS_COL  = "PAÍS"  # si tu global a veces trae "PAIS", lo soportamos abajo

TOOL_COLS = ["LUSHA", "SQL", "SH", "APL"]

def norm_text(x):
    if x is None or pd.isna(x):
        return ""
    s = str(x).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def name_key(x):
    """Normaliza nombres para matching blando: minúsculas, sin dobles espacios."""
    s = norm_text(x)
    # opcional: quitar puntuación común (mejora match con contact_name)
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s if s else None

def phone_key(x):
    if x is None or pd.isna(x):
        return None
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.replace("+", "")
    s = re.sub(r"[()\s\-]", "", s)
    s = re.sub(r"\D+", "", s)
    return s if s else None

def build_contact_id(row):
    """
    Preferimos email corporativo si existe.
    Si no, nombre + razón social + país.
    """
    email = norm_text(row.get(EMAIL_COL, ""))
    if email:
        return f"email::{email}"

    name = norm_text(row.get(NAME_COL, ""))
    comp = norm_text(row.get(COMP_COL, ""))
    pais = norm_text(row.get(PAIS_COL, ""))

    return f"key::{name}||{comp}||{pais}"

def infer_tool(row):
    for c in TOOL_COLS:
        if c in row.index:
            v = str(row.get(c, "")).strip().lower()
            if v in {"si", "sí", "1", "true", "x"}:
                return c
    return "SIN_HERRAMIENTA"

def main():
    if not os.path.exists(INPUT_PATH):
        raise FileNotFoundError(f"No existe: {INPUT_PATH}")

    df = pd.read_excel(INPUT_PATH, dtype=str).dropna(how="all")

    # Compatibilidad PAÍS vs PAIS
    if PAIS_COL not in df.columns and "PAIS" in df.columns:
        df[PAIS_COL] = df["PAIS"]

    # Asegurar columnas base existan
    for c in [EMAIL_COL, NAME_COL, COMP_COL, PAIS_COL, "CAMPAÑA"]:
        if c not in df.columns:
            df[c] = ""

    # Asegurar teléfonos personales
    for c in PHONE_COLS:
        if c not in df.columns:
            df[c] = None

    # Asegurar teléfono empresa
    if EMP_PHONE_COL not in df.columns:
        df[EMP_PHONE_COL] = None

    # Construir IDs y llaves de nombre
    df["contact_id"] = df.apply(build_contact_id, axis=1)
    df["HERRAMIENTA"] = df.apply(infer_tool, axis=1)
    df["name_key"] = df[NAME_COL].apply(name_key)

    # Contactos únicos (1 fila por contact_id)
    contacts = df.drop_duplicates(subset=["contact_id"], keep="first").copy()

    # -------- BASE LONG: teléfonos personales (match solo por phone_key) --------
    long_parts = []
    for col in PHONE_COLS:
        tmp = df[["contact_id", "CAMPAÑA", "HERRAMIENTA", "name_key", col]].copy()
        tmp = tmp.rename(columns={col: "raw_phone"})
        tmp["phone_key"] = tmp["raw_phone"].apply(phone_key)
        tmp["phone_type"] = col
        tmp["requires_name_match"] = 0  # no exige doble verificación
        tmp = tmp[tmp["phone_key"].notna()]
        long_parts.append(tmp[["contact_id", "phone_key", "phone_type", "CAMPAÑA", "HERRAMIENTA", "name_key", "requires_name_match"]])

    # -------- BASE LONG: teléfono empresa (requiere doble verificación por nombre) --------
    tmp_emp = df[["contact_id", "CAMPAÑA", "HERRAMIENTA", "name_key", EMP_PHONE_COL]].copy()
    tmp_emp = tmp_emp.rename(columns={EMP_PHONE_COL: "raw_phone"})
    tmp_emp["phone_key"] = tmp_emp["raw_phone"].apply(phone_key)
    tmp_emp["phone_type"] = "TEL_EMPRESA"
    tmp_emp["requires_name_match"] = 1  # <- clave: se valida con nombre en el cruce
    # importante: solo dejamos los que tienen name_key; si no hay nombre, no sirve para doble check
    tmp_emp = tmp_emp[(tmp_emp["phone_key"].notna()) & (tmp_emp["name_key"].notna())]
    long_parts.append(tmp_emp[["contact_id", "phone_key", "phone_type", "CAMPAÑA", "HERRAMIENTA", "name_key", "requires_name_match"]])

    base_long = pd.concat(long_parts, ignore_index=True)

    # Dedupe: por contact_id + phone_key (no queremos inflar)
    base_long = base_long.drop_duplicates(subset=["contact_id", "phone_key"], keep="first")

    os.makedirs(OUT_FOLDER, exist_ok=True)
    contacts.to_excel(OUT_CONTACTS, index=False)
    base_long.to_excel(OUT_LONG, index=False)

    print("\nOK - Base contactos construida (incluye TEL_EMPRESA con doble verificación)")
    print(f"Tipo: {LABEL}")
    print(f"Input: {INPUT_PATH}")
    print(f"Contactos únicos: {contacts['contact_id'].nunique():,}")
    print(f"Filas base_long (contact_id+phone): {len(base_long):,}")
    print(f"Salida contactos: {OUT_CONTACTS}")
    print(f"Salida long: {OUT_LONG}")

    n_emp = int((base_long["phone_type"] == "TEL_EMPRESA").sum())
    print(f"Filas TEL_EMPRESA incluidas: {n_emp:,} (solo si había nombre para validar)")

if __name__ == "__main__":
    main()