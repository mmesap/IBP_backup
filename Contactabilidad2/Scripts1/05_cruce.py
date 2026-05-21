import os
import pandas as pd

# -------- CONFIG --------
MARC_PATH = "./Contactabilidad/out/marcaciones_enero_consolidadas.xlsx"
ENT_PATH  = "./Contactabilidad/out/entregas_consolidadas_clean.xlsx"
OUT_PATH  = "./Contactabilidad/out/cruce_marcaciones_vs_entregas.xlsx"

# En marcaciones ya tienes telefono_key; en entregas tienes key_celular/key_tel1/key_tel2
MARC_TEL_COL = "telefono_key"

ENT_KEYS = [
    ("key_celular", "CELULAR"),
    ("key_tel1", "TELEFONO 1"),
    ("key_tel2", "TELEFONO 2"),
    ("key_tel_empresa", "TELEFONO DE EMPRESA"),
]


# Columnas que queremos traer desde entregas cuando el match sea único
ENT_BRING = [
    "CAMPAÑA",
    "FECHA DE CARGUE A LA TELEFONIA",
    "NOMBRE COMERCIAL EMPRESA",
    "RAZON SOCIAL",
    "INTERLOCUTOR 1",
    "CARGO",
    "AREA",
    "MANAGEMENT LEVEL",
    "LUSHA", "SH", "SQL", "APL",
    "# ENTREGA",
    "PROFILER",
    "__ARCHIVO_ORIGEN__",
]
# ------------------------

def main():
    if not os.path.exists(MARC_PATH):
        raise FileNotFoundError(f"No existe: {MARC_PATH}")
    if not os.path.exists(ENT_PATH):
        raise FileNotFoundError(f"No existe: {ENT_PATH}")

    marc = pd.read_excel(MARC_PATH, dtype=str)
    ent  = pd.read_excel(ENT_PATH, dtype=str)

    if MARC_TEL_COL not in marc.columns:
        raise RuntimeError(f"Marcaciones no tiene columna {MARC_TEL_COL}")

    # Construir índice teléfono -> posibles matches (pueden ser múltiples)
    bring_cols = [c for c in ENT_BRING if c in ent.columns]

    idx_parts = []
    for key_col, matched_by in ENT_KEYS:
        if key_col not in ent.columns:
            ent[key_col] = None
        tmp = ent[bring_cols + [key_col]].copy()
        tmp = tmp.rename(columns={key_col: "telefono_key"})
        tmp["matched_by"] = matched_by
        tmp = tmp[tmp["telefono_key"].notna()]
        idx_parts.append(tmp)

    idx = pd.concat(idx_parts, ignore_index=True)

    # Reducir duplicados exactos para no inflar matches
    dedup_subset = ["telefono_key"] + [c for c in ["CAMPAÑA", "RAZON SOCIAL", "INTERLOCUTOR 1"] if c in idx.columns]
    idx = idx.drop_duplicates(subset=dedup_subset)

    # Resumir por teléfono: UNICO vs MULTIPLE
    # Guardamos un "representante" si es único, y agregados si es múltiple
    g = idx.groupby("telefono_key", dropna=True)

    rows = []
    for tel, grp in g:
        out = {"telefono_key": tel, "match_count": len(grp)}
        if len(grp) == 1:
            out["match_type"] = "UNICO"
            r = grp.iloc[0]
            for c in bring_cols:
                out[c] = r.get(c, None)
            out["matched_by"] = r.get("matched_by", None)
        else:
            out["match_type"] = "MULTIPLE"
            if "CAMPAÑA" in grp.columns:
                out["CAMPAÑAS_MATCH"] = " | ".join(sorted(set(grp["CAMPAÑA"].dropna().astype(str))))
            out["MATCHED_BY_COLS"] = " | ".join(sorted(set(grp["matched_by"].dropna().astype(str))))
        rows.append(out)

    summary = pd.DataFrame(rows)

    # Merge final: marcaciones + info de match
    out = marc.merge(summary, left_on=MARC_TEL_COL, right_on="telefono_key", how="left")
    out["match"] = out["match_count"].notna().astype(int)

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    out.to_excel(OUT_PATH, index=False)

    # Mini-resumen útil para tu informe
    total = len(out)
    matched = int(out["match"].sum())
    unico = int((out["match_type"] == "UNICO").sum())
    multi = int((out["match_type"] == "MULTIPLE").sum())

    print("OK. Cruce generado.")
    print(f"Salida: {OUT_PATH}")
    print(f"Marcaciones totales: {total:,}")
    print(f"Con match: {matched:,} ({matched/total:.1%})")
    print(f"Match único: {unico:,} | Match múltiple (ambiguo): {multi:,}")

if __name__ == "__main__":
    main()
