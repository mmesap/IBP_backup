import re
import unicodedata
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd


INPUT_FILE = Path("/Users/elizabethmesa/ibp/Data/Base global/GlobalDataUpdated-30-3-2026.xlsx")
SHEET_NAME = None  # None = primera hoja
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "fase_1_normalizacion_texto"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "GlobalDataUpdated-30-3-2026_fase1.xlsx"
CHANGES_FILE = OUTPUT_DIR / "reporte_cambios_fase1.xlsx"
README_FILE = OUTPUT_DIR / "README_fase1.txt"


def normalize_colname(name: str) -> str:
    if name is None:
        return ""
    text = str(name).strip().lower()
    replacements = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ü": "u", "ñ": "n"
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"\s+", " ", text)
    return text


def build_column_map(columns: List[str]) -> Dict[str, str]:
    return {normalize_colname(col): col for col in columns}


def find_existing_columns(colmap: Dict[str, str], candidates: List[str]) -> List[str]:
    found = []
    for cand in candidates:
        key = normalize_colname(cand)
        if key in colmap:
            found.append(colmap[key])

    unique_found = []
    seen = set()
    for col in found:
        if col not in seen:
            unique_found.append(col)
            seen.add(col)
    return unique_found


def is_null_like(value) -> bool:
    if pd.isna(value):
        return True
    text = str(value).strip()
    return text == "" or text.lower() in {"nan", "none", "<na>"}


def strip_accents(text: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(ch)
    )


def clean_text_value(value):
    if is_null_like(value):
        return pd.NA

    text = str(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\xa0", " ").replace("\t", " ").replace("\n", " ").replace("\r", " ")
    text = "".join(ch for ch in text if unicodedata.category(ch)[0] != "C" or ch == " ")
    text = re.sub(r"\s+", " ", text).strip()
    text = strip_accents(text)
    text = text.upper()
    text = re.sub(r"\s*([,;:/\-–—])\s*", r" \1 ", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"^[\s,;:/\-–—\.]+", "", text)
    text = re.sub(r"[\s,;:/\-–—\.]+$", "", text)
    return text if text else pd.NA


def profile_changes(original: pd.Series, cleaned: pd.Series) -> Tuple[int, int]:
    orig_null = original.isna()
    clean_null = cleaned.isna()
    changed = ((original.astype("string") != cleaned.astype("string")) & ~(orig_null & clean_null)).fillna(False)
    return int(changed.sum()), int(original.notna().sum())


if not INPUT_FILE.exists():
    raise FileNotFoundError(f"No se encontró el archivo fuente: {INPUT_FILE}")

if SHEET_NAME is None:
    df = pd.read_excel(INPUT_FILE)
    sheet_used = "Primera hoja"
else:
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
    sheet_used = SHEET_NAME

df_out = df.copy()
colmap = build_column_map(list(df.columns))

target_candidates = [
    "campaña",
    "campana",
    "nombre comercial empresa",
    "nombre comercial",
    "razon social",
    "razón social",
    "pais",
    "país",
    "departamento / estado",
    "departamento/estado",
    "departamento",
    "estado",
    "ciudad / provincias",
    "ciudad/provincias",
    "ciudad",
    "provincia",
    "industria",
    "sector",
    "descripción compañía",
    "descripcion compañia",
    "descripcion compania",
    "area",
    "área",
    "cargo",
    "interlocutor 1",
    "interlocutor",
    "management level",
    "managementlevel",
    "profiler",
    "status",
    "novedades perfilación",
    "novedades perfilacion",
]

target_columns = find_existing_columns(colmap, target_candidates)

changes_summary = []
samples_by_column = {}

for col in target_columns:
    original = df_out[col].copy()
    cleaned = original.apply(clean_text_value)
    df_out[col] = cleaned

    changed_count, _ = profile_changes(original, cleaned)
    completeness_before = round((original.notna().sum() / len(df_out) * 100), 2) if len(df_out) else 0.0
    completeness_after = round((cleaned.notna().sum() / len(df_out) * 100), 2) if len(df_out) else 0.0

    changes_summary.append({
        "column": col,
        "rows_total": int(len(df_out)),
        "non_null_before": int(original.notna().sum()),
        "non_null_after": int(cleaned.notna().sum()),
        "completeness_before_pct": completeness_before,
        "completeness_after_pct": completeness_after,
        "changed_values": changed_count,
        "pct_changed_over_total": round((changed_count / len(df_out) * 100), 2) if len(df_out) else 0.0,
        "unique_before": int(original.dropna().astype("string").nunique()),
        "unique_after": int(cleaned.dropna().astype("string").nunique()),
    })

    diff_mask = (
        (original.astype("string") != cleaned.astype("string")) &
        ~(original.isna() & cleaned.isna())
    ).fillna(False)

    sample_df = pd.DataFrame({
        "before": original[diff_mask].astype("string"),
        "after": cleaned[diff_mask].astype("string"),
    }).head(200)

    samples_by_column[col] = sample_df

changes_df = pd.DataFrame(changes_summary).sort_values(
    by=["changed_values", "column"], ascending=[False, True]
)

general_summary = pd.DataFrame([
    {"metric": "input_file", "value": str(INPUT_FILE)},
    {"metric": "sheet_used", "value": sheet_used},
    {"metric": "output_file", "value": str(OUTPUT_FILE)},
    {"metric": "rows", "value": int(df_out.shape[0])},
    {"metric": "columns", "value": int(df_out.shape[1])},
    {"metric": "target_columns_normalized", "value": len(target_columns)},
    {"metric": "total_changed_cells", "value": int(changes_df["changed_values"].sum()) if not changes_df.empty else 0},
])

target_columns_df = pd.DataFrame({"normalized_columns": target_columns})

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df_out.to_excel(writer, index=False, sheet_name="data_fase1")

with pd.ExcelWriter(CHANGES_FILE, engine="openpyxl") as writer:
    general_summary.to_excel(writer, sheet_name="general_summary", index=False)
    target_columns_df.to_excel(writer, sheet_name="target_columns", index=False)
    changes_df.to_excel(writer, sheet_name="changes_summary", index=False)

    for col, sample_df in samples_by_column.items():
        sheet_name = re.sub(r"[^A-Za-z0-9_]", "_", col)[:31]
        if sample_df.empty:
            pd.DataFrame({"before": [], "after": []}).to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            sample_df.to_excel(writer, sheet_name=sheet_name, index=False)

with open(README_FILE, "w", encoding="utf-8") as f:
    f.write("FASE 1 - NORMALIZACION ESTRUCTURAL BASICA\n")
    f.write("=" * 50 + "\n\n")
    f.write("Objetivo:\n")
    f.write("Estandarizar columnas textuales generales sin alterar aun correos, telefonos, URLs ni logica de deduplicacion.\n\n")
    f.write("Reglas aplicadas:\n")
    f.write("- Trim de espacios al inicio y final\n")
    f.write("- Colapsar espacios repetidos\n")
    f.write("- Eliminar caracteres invisibles/control\n")
    f.write("- Normalizacion Unicode\n")
    f.write("- Remover tildes\n")
    f.write("- Convertir a mayusculas\n")
    f.write("- Limpiar separadores sobrantes en extremos\n\n")
    f.write("Archivos generados:\n")
    f.write(f"- Base normalizada: {OUTPUT_FILE}\n")
    f.write(f"- Reporte de cambios: {CHANGES_FILE}\n")
    f.write(f"- Resumen de fase: {README_FILE}\n\n")
    f.write("Importante:\n")
    f.write("Esta fase no modifica la base original y no aplica todavia reglas semanticas como quitar SAS, LTDA, INC, ni estandarizar correos, telefonos o LinkedIn.\n")

print("Fase 1 completada correctamente.")
print(f"Base normalizada: {OUTPUT_FILE}")
print(f"Reporte de cambios: {CHANGES_FILE}")
print(f"Columnas normalizadas: {len(target_columns)}")
