import os
import pandas as pd
import sys
import re

# =========================
# USO:
# python 01_consolidar_marcaciones.py 2026-01
# python 01_consolidar_marcaciones.py 2026-02_H1
# =========================

if len(sys.argv) != 2:
    print("Uso: python 01_consolidar_marcaciones.py YYYY-MM o YYYY-MM_H1")
    sys.exit(1)

PERIODO = sys.argv[1]

INPUT_PATH = f"./Contactabilidad/Data/marcaciones/{PERIODO}.xlsx"
OUT_FOLDER = f"./Contactabilidad/out/{PERIODO[:7]}"
OUT_PATH = f"{OUT_FOLDER}/marcaciones_{PERIODO}_consolidadas.xlsx"

def normalize_phone(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    # Quitar ".0" por Excel
    if s.endswith(".0"):
        s = s[:-2]
    # quitar +, espacios, guiones, paréntesis
    s = s.replace("+", "")
    s = re.sub(r"[()\s\-]", "", s)
    # dejar solo dígitos
    s = re.sub(r"\D+", "", s)
    return s if s else None

def pick_dedup_key(df: pd.DataFrame):
    """
    Devuelve una lista de columnas para deduplicar.
    Preferimos IDs reales de llamada (cdr_id/uuid).
    """
    for c in ["cdr_id", "uuid", "id", "call_id"]:
        if c in df.columns:
            return [c]
    # fallback: combinación fuerte (si existen)
    candidates = [c for c in ["date", "contact_number", "agent_id", "direction", "status", "talking_time", "ringing_time"] if c in df.columns]
    if len(candidates) >= 2:
        return candidates
    # último recurso: todas las columnas
    return list(df.columns)

def main():
    if not os.path.exists(INPUT_PATH):
        raise FileNotFoundError(f"No existe archivo: {INPUT_PATH}")

    print(f"\nLeyendo marcaciones: {INPUT_PATH}")

    xls = pd.ExcelFile(INPUT_PATH)
    sheets = xls.sheet_names
    print(f"Hojas detectadas: {len(sheets)} -> {sheets}")

    frames = []
    per_sheet_counts = []

    for sh in sheets:
        df = pd.read_excel(INPUT_PATH, sheet_name=sh, dtype=str)
        df["__sheet__"] = sh
        df = df.dropna(how="all")
        frames.append(df)
        per_sheet_counts.append((sh, len(df)))

    print("Filas por hoja:")
    for sh, n in per_sheet_counts:
        print(f" - {sh}: {n:,}")

    marc = pd.concat(frames, ignore_index=True, sort=False)
    marc = marc.dropna(how="all")

    # Normalizar nombres columnas
    marc.columns = [str(c).strip() for c in marc.columns]

    # Validar columna base
    if "contact_number" not in marc.columns:
        raise RuntimeError("No existe columna 'contact_number' en marcaciones")

    # phone_key limpio
    marc["phone_key"] = marc["contact_number"].apply(normalize_phone)

    # Normalizar status
    if "status" in marc.columns:
        marc["status"] = marc["status"].astype(str).str.lower().str.strip()

    # DEDUPLICAR (esto es lo que evita que se te dupliquen a 6500)
    dedup_cols = pick_dedup_key(marc)
    before = len(marc)
    marc = marc.drop_duplicates(subset=dedup_cols, keep="first")
    after = len(marc)

    print(f"\nDeduplicación usando columnas: {dedup_cols}")
    print(f"Filas antes: {before:,} | después: {after:,} | removidas: {before-after:,}")

    # call_id estable (si ya hay cdr_id/uuid lo usamos; si no, index)
    if "cdr_id" in marc.columns:
        marc["call_id"] = marc["cdr_id"].astype(str)
    elif "uuid" in marc.columns:
        marc["call_id"] = marc["uuid"].astype(str)
    else:
        marc["call_id"] = marc.reset_index().index.astype(str)

    os.makedirs(OUT_FOLDER, exist_ok=True)
    marc.to_excel(OUT_PATH, index=False)

    print("\nOK - Marcaciones consolidadas")
    print(f"Periodo: {PERIODO}")
    print(f"Filas finales: {len(marc):,}")
    print(f"Salida: {OUT_PATH}")

if __name__ == "__main__":
    main()