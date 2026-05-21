import os
import re
import pandas as pd

# -------- CONFIG --------
MARCACIONES_PATH = "./Contactabilidad/Data/marcaciones/Marcaciones enero.xlsx"
OUTPUT_PATH = "./Contactabilidad/out/marcaciones_enero_consolidadas.xlsx"

PHONE_COL = "contact_number"   # como lo indicaste
# ------------------------

def phone_key(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.replace("+", "")
    s = s.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
    s = re.sub(r"\D+", "", s)
    return s if s != "" else None

def main():
    if not os.path.exists(MARCACIONES_PATH):
        raise FileNotFoundError(f"No existe: {MARCACIONES_PATH}")

    xls = pd.ExcelFile(MARCACIONES_PATH)
    frames = []
    skipped = []

    for sh in xls.sheet_names:
        df = pd.read_excel(MARCACIONES_PATH, sheet_name=sh, dtype=str)

        if df is None or df.empty:
            skipped.append((sh, "hoja vacía"))
            continue

        # limpiar nombres de columnas (solo trim)
        df.columns = [str(c).strip() for c in df.columns]

        if PHONE_COL not in df.columns:
            skipped.append((sh, f"no tiene columna {PHONE_COL}"))
            continue

        # añadir metadata
        df["sheet"] = sh
        df["telefono_key"] = df[PHONE_COL].apply(phone_key)

        # filtrar filas sin teléfono_key
        df = df[df["telefono_key"].notna()].copy()

        frames.append(df)

    if not frames:
        raise RuntimeError(
            "No se pudo consolidar ninguna hoja.\n"
            f"Hojas detectadas: {xls.sheet_names}\n"
            "Revisa que todas tengan la columna contact_number."
        )

    marc = pd.concat(frames, ignore_index=True, sort=False)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    marc.to_excel(OUTPUT_PATH, index=False)

    print("OK. Marcaciones consolidadas (full) generadas.")
    print(f"Salida: {OUTPUT_PATH}")
    print(f"Hojas leídas: {len(frames)} / {len(xls.sheet_names)}")
    print(f"Filas totales (con teléfono): {len(marc):,}")
    print(f"Columnas totales: {len(marc.columns)}")

    if skipped:
        print("\nHojas omitidas (si aplica):")
        for sh, reason in skipped:
            print(f"- {sh}: {reason}")

if __name__ == "__main__":
    main()
