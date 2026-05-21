import os
import re
import pandas as pd

INPUT_PATH = "./Contactabilidad/out/bases_enero_consolidado.xlsx"
OUTPUT_PATH = "./Contactabilidad/out/bases_enero_consolidado_clean.xlsx"

PHONE_COLS = ["CELULAR", "TELEFONO 1", "TELEFONO 2", "TELÉFONO DE EMPRESA"]

def phone_key(x):
    if pd.isna(x) or x is None:
        return None
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.replace("+", "").replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
    s = re.sub(r"\D+", "", s)
    return s if s else None

def main():
    if not os.path.exists(INPUT_PATH):
        raise FileNotFoundError(f"No existe: {INPUT_PATH}")

    df = pd.read_excel(INPUT_PATH, dtype=str)

    # asegurar columnas
    for col in PHONE_COLS:
        if col not in df.columns:
            df[col] = None

    # crear keys
    df["key_celular"] = df["CELULAR"].apply(phone_key)
    df["key_tel1"] = df["TELEFONO 1"].apply(phone_key)
    df["key_tel2"] = df["TELEFONO 2"].apply(phone_key)
    df["key_tel_empresa"] = df["TELÉFONO DE EMPRESA"].apply(phone_key)

    # calcular teléfonos únicos (uniendo las 4 columnas)
    all_keys = pd.concat([
        df["key_celular"],
        df["key_tel1"],
        df["key_tel2"],
        df["key_tel_empresa"],
    ], ignore_index=True).dropna()

    unique_phones = all_keys.nunique()

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    df.to_excel(OUTPUT_PATH, index=False)

    print("OK - January bases cleaned")
    print(f"Rows: {len(df):,}")
    print(f"Unique phones delivered in January: {unique_phones:,}")
    print(f"Output: {OUTPUT_PATH}")

if __name__ == "__main__":
    main()
