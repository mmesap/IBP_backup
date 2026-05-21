import os
import glob
import pandas as pd

# -------- CONFIG --------
CARPETA_CAMPANAS = "./Contactabilidad/Data/campanas"
SALIDA = "./Contactabilidad/out/entregas_consolidadas.xlsx"
SHEET_NAME = "IMR"
# ------------------------

# columnas que NO deben quedar en el consolidado
EXCLUDE_COLS = {
    "IMR",
    "GESTIÓN",
    "GESTION",
    "STATUS DE GESTION",
    "FECHA - ULTIMA GESTIÓN",
    "FECHA - ULTIMA GESTION",
    "TELEFONO VERIFICADO",
    "CORREO VERIFICADO",
    "CARGO REAL",
    "OBSERVACIONES",
}

def std_colname(c: str) -> str:
    s = str(c).strip()
    s = " ".join(s.split())
    s = s.upper()

    mapping = {
        "MONEDA FACTURACIÓN": "MONEDA FACTURACION",
        "ÁREA": "AREA",
        "CORREO ELECTRÓNICO CORPORATIVO": "CORREO ELECTRONICO CORPORATIVO",
        "CORREO ELECTRÓNICO PERSONAL": "CORREO ELECTRONICO PERSONAL",
        "FACTURACIÓN EMPRESA": "FACTURACION EMPRESA",
        "NÚMERO DE EMPLEADOS": "NUMERO DE EMPLEADOS",
        "EXTENSION  2": "EXTENSION 2",
    }
    return mapping.get(s, s)

def main():
    files = sorted(glob.glob(os.path.join(CARPETA_CAMPANAS, "*.xlsx")))
    if not files:
        raise FileNotFoundError(f"No se encontraron .xlsx en {CARPETA_CAMPANAS}")

    frames = []
    skipped = []

    for fp in files:
        try:
            df = pd.read_excel(fp, sheet_name=SHEET_NAME, dtype=str)
        except ValueError:
            skipped.append((os.path.basename(fp), "no tiene hoja IMR"))
            continue
        except Exception as e:
            skipped.append((os.path.basename(fp), f"error leyendo: {e}"))
            continue

        # normalizar nombres de columnas
        df.columns = [std_colname(c) for c in df.columns]

        # eliminar columnas de gestión/IMR si existen
        cols_to_drop = [c for c in df.columns if c in EXCLUDE_COLS]
        df = df.drop(columns=cols_to_drop, errors="ignore")

        # metadata mínima para trazabilidad
        df["__ARCHIVO_ORIGEN__"] = os.path.basename(fp)

        frames.append(df)

    if not frames:
        raise RuntimeError("No se pudo leer la hoja IMR de ningún archivo.")

    consolidado = pd.concat(frames, ignore_index=True, sort=False)
    consolidado = consolidado.dropna(how="all")

    os.makedirs(os.path.dirname(SALIDA), exist_ok=True)
    consolidado.to_excel(SALIDA, index=False)

    print(f"OK. Consolidado generado: {SALIDA}")
    print(f"Archivos leídos (IMR): {len(frames)} / {len(files)}")
    print(f"Filas totales: {len(consolidado):,}")
    print(f"Columnas totales: {len(consolidado.columns)}")

    if skipped:
        print("\nArchivos omitidos:")
        for name, reason in skipped:
            print(f"- {name}: {reason}")

if __name__ == "__main__":
    main()
