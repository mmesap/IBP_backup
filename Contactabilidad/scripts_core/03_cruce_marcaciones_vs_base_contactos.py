# -*- coding: utf-8 -*-

import os
import sys
import re
import unicodedata
import glob
import pandas as pd

# ==========================================
# USO:
# python 03_cruce_marcaciones_vs_base_contactos.py 2026-01 GLOBAL
# python 03_cruce_marcaciones_vs_base_contactos.py 2026-01 MES
# ==========================================

if len(sys.argv) != 3:
    print("Uso: python 03_cruce_marcaciones_vs_base_contactos.py YYYY-MM (GLOBAL|MES)")
    sys.exit(1)

MES = sys.argv[1].strip()
TIPO = sys.argv[2].upper().strip()

if TIPO not in {"GLOBAL", "MES"}:
    raise ValueError("El segundo argumento debe ser GLOBAL o MES")

OUT_MES_DIR = f"./Contactabilidad/out/{MES}"

# Marcaciones: intentamos varias ubicaciones/nombres
MARC_CANDIDATES = [
    f"{OUT_MES_DIR}/A_historico/marcaciones_{MES}_consolidado.xlsx",
    f"{OUT_MES_DIR}/A_historico/marcaciones_{MES}_consolidadas.xlsx",
    f"{OUT_MES_DIR}/marcaciones_{MES}_consolidado.xlsx",
    f"{OUT_MES_DIR}/marcaciones_{MES}_consolidadas.xlsx",
]

def resolve_first_existing(paths):
    for p in paths:
        if os.path.exists(p):
            return p
    return None

MARC_PATH = resolve_first_existing(MARC_CANDIDATES)

if MARC_PATH is None:
    pattern = f"{OUT_MES_DIR}/**/marcaciones_{MES}_*.xlsx"
    hits = sorted(glob.glob(pattern, recursive=True))
    MARC_PATH = hits[0] if hits else None

if MARC_PATH is None:
    raise FileNotFoundError(
        "No se encontró archivo de marcaciones.\n"
        f"Busqué en:\n- " + "\n- ".join(MARC_CANDIDATES) + "\n"
        f"Y también por patrón: {OUT_MES_DIR}/**/marcaciones_{MES}_*.xlsx"
    )

if TIPO == "GLOBAL":
    BASE_LONG_PATH = f"{OUT_MES_DIR}/A_historico/base_contactos_long_GLOBAL.xlsx"
    OUT_FOLDER = f"{OUT_MES_DIR}/A_historico"
else:
    BASE_LONG_PATH = f"{OUT_MES_DIR}/B_mes/base_contactos_long_ENTREGAS_{MES}.xlsx"
    OUT_FOLDER = f"{OUT_MES_DIR}/B_mes"

OUT_PATH = f"{OUT_FOLDER}/cruce_{MES}_{TIPO}_CONTACTOS.xlsx"

# Columnas esperadas
MARC_TEL_COL = "contact_number"
MARC_NAME_COL = "contact_name"

BASE_MATCHED_BY_COL = "matched_by"
BASE_NAME_KEY_COL = "name_key"


# ================= UTILIDADES =================

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

def norm_text(x):
    if x is None or pd.isna(x):
        return None
    s = str(x).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^a-z0-9 ]+", "", s)
    return s if s else None


# ================= MAIN =================

def main():

    if not os.path.exists(BASE_LONG_PATH):
        raise FileNotFoundError(f"No existe base_long: {BASE_LONG_PATH}")

    print("\n==============================")
    print(f"Periodo: {MES}")
    print(f"Tipo base: {TIPO}")
    print(f"Marcaciones: {MARC_PATH}")
    print(f"Base long:   {BASE_LONG_PATH}")
    print("==============================\n")

    marc = pd.read_excel(MARC_PATH, dtype=str)
    base = pd.read_excel(BASE_LONG_PATH, dtype=str)

    # --- Validaciones mínimas
    if MARC_TEL_COL not in marc.columns:
        raise RuntimeError(f"Marcaciones no tiene la columna requerida: {MARC_TEL_COL}")
    if "phone_key" not in base.columns:
        raise RuntimeError("base_long no tiene la columna requerida: phone_key")
    if "contact_id" not in base.columns:
        raise RuntimeError("base_long no tiene la columna requerida: contact_id")

    # --- Normalizar llaves
    marc["telefono_key"] = marc[MARC_TEL_COL].apply(phone_key)
    marc["name_key_marc"] = marc[MARC_NAME_COL].apply(norm_text) if MARC_NAME_COL in marc.columns else None

    # call_id estable
    if "uuid" in marc.columns:
        marc["call_id"] = marc["uuid"]
    elif "cdr_id" in marc.columns:
        marc["call_id"] = marc["cdr_id"]
    else:
        marc["call_id"] = marc.index.astype(str)

    base["phone_key"] = base["phone_key"].apply(phone_key)
    base[BASE_NAME_KEY_COL] = base[BASE_NAME_KEY_COL].apply(norm_text) if BASE_NAME_KEY_COL in base.columns else None

    # --- Merge con sufijos controlados
    jm = marc.merge(
        base,
        left_on="telefono_key",
        right_on="phone_key",
        how="left",
        suffixes=("_marc", "_base"),
    )

    # A partir de aquí, el contact_id correcto (de la base) debe ser:
    # contact_id_base (porque viene de base)
    CONTACT_ID_COL = "contact_id_base" if "contact_id_base" in jm.columns else None
    if CONTACT_ID_COL is None:
        # fallback por si el nombre viniera diferente
        candidates = [c for c in jm.columns if c.endswith("_base") and c.startswith("contact_id")]
        if candidates:
            CONTACT_ID_COL = candidates[0]
        else:
            raise RuntimeError(
                "ERROR: no encontré contact_id de la base después del merge.\n"
                f"Columnas disponibles con 'contact': {[c for c in jm.columns if 'contact' in c.lower()]}"
            )

    # --- Validación TEL_EMPRESA: si el match fue por TEL_EMPRESA, exige nombre igual
    if BASE_MATCHED_BY_COL in jm.columns:
        is_tel_emp = jm[BASE_MATCHED_BY_COL].fillna("").eq("TEL_EMPRESA")
        name_ok = (
            jm["name_key_marc"].notna()
            & jm[BASE_NAME_KEY_COL].notna()
            & (jm["name_key_marc"] == jm[BASE_NAME_KEY_COL])
        )
        invalid_tel_emp = is_tel_emp & (~name_ok)
        jm.loc[invalid_tel_emp, [CONTACT_ID_COL, "CAMPAÑA_STD"]] = None

    # --- Match por llamada
    jm["has_match"] = jm[CONTACT_ID_COL].notna().astype(int)

    per_call = jm.groupby("call_id")[CONTACT_ID_COL].nunique(dropna=True)
    jm = jm.merge(per_call.rename("contact_match_count"), on="call_id", how="left")
    jm["contact_match_count"] = jm["contact_match_count"].fillna(0).astype(int)

    jm["match_type"] = jm["contact_match_count"].apply(
        lambda n: "SIN_MATCH" if n == 0 else ("UNICO" if n == 1 else "MULTIPLE")
    )

    os.makedirs(OUT_FOLDER, exist_ok=True)
    jm.to_excel(OUT_PATH, index=False)

    # --- Resumen consola
    total = len(marc)
    matched = int(jm["has_match"].sum())
    unico = int((jm["match_type"] == "UNICO").sum())
    multi = int((jm["match_type"] == "MULTIPLE").sum())

    print("Cruce generado correctamente.")
    print(f"Marcaciones totales: {total:,}")
    print(f"Con match: {matched:,} ({matched/total:.2%})")
    print(f"Match único: {unico:,}")
    print(f"Match múltiple: {multi:,}")
    print(f"Archivo salida: {OUT_PATH}")

if __name__ == "__main__":
    main()