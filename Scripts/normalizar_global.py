import os
import pandas as pd

# ==============================
# RUTAS
# ==============================
INPUT_XLSX = "/Users/elizabethmesa/ibp/Data/GlobalDataUpdated-17-2-2025.xlsx"
OUTPUT_XLSX = os.path.join(
    os.path.dirname(INPUT_XLSX),
    "GlobalDataUpdatedNorm-17-2-2025.xlsx"
)

# ==============================
# CONFIG
# ==============================
PHONE_COLS = ["TELÉFONO DE EMPRESA", "CELULAR", "TELEFONO 1", "TELEFONO 2"]
DATE_COL = "Fecha Carga"

def normalize_phone_value(x):
    """
    Removes spaces, '-', '(' and ')' from phone-like values.
    Keeps NaN/None as is.
    """
    if pd.isna(x):
        return x
    
    s = str(x)
    s = (
        s.replace(" ", "")
         .replace("-", "")
         .replace("(", "")
         .replace(")", "")
    )
    
    return s

def main():
    df = pd.read_excel(INPUT_XLSX, dtype=object)

    # 1) Normalize phone columns
    for col in PHONE_COLS:
        if col in df.columns:
            df[col] = df[col].apply(normalize_phone_value)

    # 2) Replace "Fecha Carga" with "Mes de carga" and "Año de carga"
    if DATE_COL in df.columns:
        dt = pd.to_datetime(df[DATE_COL], errors="coerce")

        df["Mes de carga"] = dt.dt.month
        df["Año de carga"] = dt.dt.year

        df = df.drop(columns=[DATE_COL])

    df.to_excel(OUTPUT_XLSX, index=False)
    print(f"OK -> Archivo creado: {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()