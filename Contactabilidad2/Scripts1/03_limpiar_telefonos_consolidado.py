import os
import re
import pandas as pd

INPUT_PATH = "./Contactabilidad/out/entregas_consolidadas.xlsx"
OUTPUT_PATH = "./Contactabilidad/out/entregas_consolidadas_clean.xlsx"

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
    if not os.path.exists(INPUT_PATH):
        raise FileNotFoundError(f"No existe: {INPUT_PATH}")

    df = pd.read_excel(INPUT_PATH, dtype=str)

    # Asegurar columnas de contacto
    for col in ["CELULAR", "TELEFONO 1", "TELEFONO 2"]:
        if col not in df.columns:
            df[col] = None

    # TELÉFONO DE EMPRESA puede venir con o sin tilde según el Excel
    col_tel_emp = None
    if "TELÉFONO DE EMPRESA" in df.columns:
        col_tel_emp = "TELÉFONO DE EMPRESA"
    elif "TELEFONO DE EMPRESA" in df.columns:
        col_tel_emp = "TELEFONO DE EMPRESA"
    else:
        # si no existe, la creamos vacía para que el pipeline siga
        df["TELEFONO DE EMPRESA"] = None
        col_tel_emp = "TELEFONO DE EMPRESA"

    # Crear llaves limpias
    df["key_celular"] = df["CELULAR"].apply(phone_key)
    df["key_tel1"]    = df["TELEFONO 1"].apply(phone_key)
    df["key_tel2"]    = df["TELEFONO 2"].apply(phone_key)
    df["key_tel_empresa"] = df[col_tel_emp].apply(phone_key)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    df.to_excel(OUTPUT_PATH, index=False)

    total = len(df)
    n_cel = df["key_celular"].notna().sum()
    n_t1  = df["key_tel1"].notna().sum()
    n_t2  = df["key_tel2"].notna().sum()
    n_emp = df["key_tel_empresa"].notna().sum()

    print("OK. Consolidado limpio generado (con tel empresa).")
    print(f"Salida: {OUTPUT_PATH}")
    print(f"Filas: {total:,}")
    print(f"Con celular: {n_cel:,} | Con tel1: {n_t1:,} | Con tel2: {n_t2:,} | Con tel empresa: {n_emp:,}")

if __name__ == "__main__":
    main()
