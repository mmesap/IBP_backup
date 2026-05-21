import os
import glob
import pandas as pd
import sys

# =========================
# USO:
# python 00_consolidar_entregas_mes.py 2026-01
# =========================

if len(sys.argv) != 2:
    print("Uso: python 00_consolidar_entregas_mes.py YYYY-MM")
    sys.exit(1)

MES = sys.argv[1]

CARPETA_ENTREGAS_MES = f"./Contactabilidad/Data/entregas/{MES}"
SALIDA = f"./Contactabilidad/out/{MES}/entregas_{MES}_consolidado.xlsx"
SHEET_NAME = "IMR"

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
    if not os.path.exists(CARPETA_ENTREGAS_MES):
        raise FileNotFoundError(f"No existe la carpeta: {CARPETA_ENTREGAS_MES}")

    files = sorted(glob.glob(os.path.join(CARPETA_ENTREGAS_MES, "*.xlsx")))
    if not files:
        raise FileNotFoundError(f"No se encontraron .xlsx en {CARPETA_ENTREGAS_MES}")

    frames = []
    skipped = []

    for fp in files:
        name = os.path.basename(fp)

        try:
            df = pd.read_excel(fp, sheet_name=SHEET_NAME, dtype=str)
        except ValueError:
            skipped.append((name, f"No tiene hoja {SHEET_NAME}"))
            continue
        except Exception as e:
            skipped.append((name, f"Error leyendo: {e}"))
            continue

        df.columns = [std_colname(c) for c in df.columns]

        cols_to_drop = [c for c in df.columns if c in EXCLUDE_COLS]
        df = df.drop(columns=cols_to_drop, errors="ignore")

        df = df.dropna(how="all")

        df["__ARCHIVO_ORIGEN__"] = name
        df["__MES_ENTREGA__"] = MES

        frames.append(df)

    if not frames:
        raise RuntimeError("No se pudo leer la hoja IMR de ningún archivo.")

    consolidado = pd.concat(frames, ignore_index=True, sort=False)
    consolidado = consolidado.dropna(how="all")

    os.makedirs(os.path.dirname(SALIDA), exist_ok=True)
    consolidado.to_excel(SALIDA, index=False)

    print("\nOK - Entregas consolidadas")
    print(f"Mes: {MES}")
    print(f"Archivos leídos: {len(frames)} / {len(files)}")
    print(f"Filas consolidadas: {len(consolidado):,}")
    print(f"Salida: {SALIDA}")

    if skipped:
        print("\nArchivos omitidos:")
        for name, reason in skipped:
            print(f"- {name}: {reason}")

if __name__ == "__main__":
    main()