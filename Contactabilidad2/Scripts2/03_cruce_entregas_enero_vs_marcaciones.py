import os
import re
import pandas as pd

# ---------------- CONFIG ----------------
BASES_CLEAN = "./Contactabilidad/out/bases_enero_consolidado_clean.xlsx"
MARCACIONES_XLSX = "./Contactabilidad/Data/marcaciones/Marcaciones enero.xlsx"  # ajusta si está en otra ruta

OUT_XLSX = "./Contactabilidad/out/informe_enero_entregas_vs_marcaciones.xlsx"

PHONE_KEYS_COLS = ["key_celular", "key_tel1", "key_tel2", "key_tel_empresa"]

# En marcaciones, el teléfono está aquí:
MARC_TEL_COL = "contact_number"

# Columnas en marcaciones
MARC_STATUS_COL = "status"
TAG2_COL = "tag 2"     # si viene diferente, lo detectamos
NOTES_COL = "notes"

# Clasificación Tag2
TAG2_VERIFICADO = {
    "volver a contactar",
    "no interes",
    "administrativo",
    "3 sbm",
    "1 mql",
    "maduracion 1",
    "maduracion 2",
    "2 em",
}
TAG2_INVALIDO = {"numero equivocado"}
NOTES_INVALID_PATTERNS = ["fuera de servicio", "num equivocado"]

# --------------------------------------

def norm_text(x):
    if pd.isna(x) or x is None:
        return ""
    s = str(x).strip().lower()
    s = " ".join(s.split())
    s = s.replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u").replace("ñ","n")
    return s

def phone_key(x):
    if pd.isna(x) or x is None:
        return None
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.replace("+","").replace(" ","").replace("-","").replace("(","").replace(")","")
    s = re.sub(r"\D+","",s)
    return s if s else None

def decide_bucket(tag2_raw, notes_raw):
    tag2 = norm_text(tag2_raw)
    notes = norm_text(notes_raw)

    if tag2 in TAG2_INVALIDO:
        return "Número equivocado"
    for pat in NOTES_INVALID_PATTERNS:
        if pat in notes:
            return "Número equivocado"

    if tag2 in TAG2_VERIFICADO:
        return "Efectivo"
    if tag2 == "no conecta":
        return "No conecta"
    if tag2 == "volver a contactar":
        return "Volver a contactar"
    if tag2 == "target no aplica":
        return "Target no aplica"
    if tag2 == "no interes":
        return "No interes"

    if tag2_raw and str(tag2_raw).strip():
        return str(tag2_raw).strip()

    return "Sin tag2"

def read_all_sheets(path):
    xls = pd.ExcelFile(path)
    dfs = []
    for sh in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sh, dtype=str)
        df["__sheet__"] = sh
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

def main():
    if not os.path.exists(BASES_CLEAN):
        raise FileNotFoundError(f"No existe: {BASES_CLEAN}")
    if not os.path.exists(MARCACIONES_XLSX):
        raise FileNotFoundError(f"No existe: {MARCACIONES_XLSX}")

    # 1) Bases enero (clean)
    b = pd.read_excel(BASES_CLEAN, dtype=str)
    b.columns = [str(c).strip() for c in b.columns]

    # Detectar columna campaña
    camp_col = None
    for cand in ["CAMPAÑA", "Campaña", "CAMPANA", "Campana", "campaña", "campana"]:
        if cand in b.columns:
            camp_col = cand
            break
    if camp_col is None:
        raise KeyError(f"No encuentro columna de campaña en bases. Columnas: {list(b.columns)}")

    # asegurar columnas de teléfono
    for col in PHONE_KEYS_COLS:
        if col not in b.columns:
            b[col] = None

    # Explode a nivel teléfono
    phone_rows = []
    for _, r in b.iterrows():
        camp = "" if pd.isna(r.get(camp_col)) else str(r.get(camp_col)).strip()

        keys = []
        for kc in PHONE_KEYS_COLS:
            v = r.get(kc)
            if pd.notna(v) and str(v).strip():
                keys.append(str(v).strip())
        keys = list(set(keys))

        for k in keys:
            phone_rows.append({"phone_key": phone_key(k), "CAMPAÑA": camp})

    base_phones = pd.DataFrame(phone_rows).dropna(subset=["phone_key"]).drop_duplicates()
    delivered_unique = int(base_phones["phone_key"].nunique())

    # 2) Marcaciones enero (todas las hojas)
    m = read_all_sheets(MARCACIONES_XLSX)
    m.columns = [str(c).strip() for c in m.columns]

    # Detectar tag2 si viene con otro nombre
    if TAG2_COL not in m.columns:
        for alt in ["tag2", "tag_2", "Tag 2", "TAG 2", "Tag2", "TAG2"]:
            if alt in m.columns:
                m.rename(columns={alt: TAG2_COL}, inplace=True)
                break

    # Asegurar columnas
    for col in [MARC_TEL_COL, MARC_STATUS_COL, TAG2_COL, NOTES_COL]:
        if col not in m.columns:
            m[col] = ""

    # Normalizar teléfono en marcaciones
    m["phone_key"] = m[MARC_TEL_COL].apply(phone_key)
    m = m.dropna(subset=["phone_key"]).copy()

    # 3) Marcaciones solo para entregas enero
    m_jan = m.merge(base_phones, on="phone_key", how="inner")

    # 4) Métricas globales
    touched_unique = int(m_jan["phone_key"].nunique())
    coverage_pct = round((touched_unique / delivered_unique) * 100, 2) if delivered_unique else 0.0

    total_calls_on_jan_bases = int(len(m_jan))
    answered_calls = int((m_jan[MARC_STATUS_COL].astype(str).str.lower() == "answered").sum())
    missed_calls = int((m_jan[MARC_STATUS_COL].astype(str).str.lower() == "missed").sum())
    contactability_pct = round((answered_calls / total_calls_on_jan_bases) * 100, 2) if total_calls_on_jan_bases else 0.0

    # 5) Tag2 bucket + distribución
    m_jan["tag2_bucket"] = m_jan.apply(lambda r: decide_bucket(r.get(TAG2_COL,""), r.get(NOTES_COL,"")), axis=1)

    tag2_dist = (
        m_jan["tag2_bucket"]
        .value_counts(dropna=False)
        .reset_index()
        .rename(columns={"index":"Tag2", "tag2_bucket":"Marcaciones"})
    )
    tag2_dist["Marcaciones"] = pd.to_numeric(tag2_dist["Marcaciones"], errors="coerce").fillna(0).astype(int)
    tag2_dist["%"] = (tag2_dist["Marcaciones"] / total_calls_on_jan_bases * 100).round(2) if total_calls_on_jan_bases else 0.0

    # 6) Por campaña: cobertura
    delivered_by_camp = base_phones.groupby("CAMPAÑA")["phone_key"].nunique().reset_index().rename(columns={"phone_key":"Telefonos_entregados"})
    touched_by_camp = m_jan.groupby("CAMPAÑA")["phone_key"].nunique().reset_index().rename(columns={"phone_key":"Telefonos_tocados"})

    cov = delivered_by_camp.merge(touched_by_camp, on="CAMPAÑA", how="left").fillna({"Telefonos_tocados":0})
    cov["Telefonos_entregados"] = pd.to_numeric(cov["Telefonos_entregados"], errors="coerce").fillna(0).astype(int)
    cov["Telefonos_tocados"] = pd.to_numeric(cov["Telefonos_tocados"], errors="coerce").fillna(0).astype(int)
    cov["% cobertura"] = cov.apply(lambda r: round((r["Telefonos_tocados"]/r["Telefonos_entregados"])*100,2) if r["Telefonos_entregados"] else 0.0, axis=1)

    # 7) Por campaña: contactabilidad
    bycamp_calls = m_jan.groupby("CAMPAÑA").size().reset_index(name="Marcaciones")
    bycamp_ans = (m_jan[MARC_STATUS_COL].astype(str).str.lower() == "answered").groupby(m_jan["CAMPAÑA"]).sum().reset_index(name="Answered")

    camp_perf = bycamp_calls.merge(bycamp_ans, on="CAMPAÑA", how="left").fillna({"Answered":0})
    camp_perf["Marcaciones"] = pd.to_numeric(camp_perf["Marcaciones"], errors="coerce").fillna(0).astype(int)
    camp_perf["Answered"] = pd.to_numeric(camp_perf["Answered"], errors="coerce").fillna(0).astype(int)
    camp_perf["% contactabilidad"] = camp_perf.apply(lambda r: round((r["Answered"]/r["Marcaciones"])*100,2) if r["Marcaciones"] else 0.0, axis=1)

    # 8) % Tag2 principales por campaña
    pivot = (
        m_jan.pivot_table(index="CAMPAÑA", columns="tag2_bucket", values="phone_key", aggfunc="count", fill_value=0)
        .reset_index()
    )

    for col in ["No conecta", "Número equivocado", "Volver a contactar", "Efectivo"]:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot = pivot.merge(bycamp_calls, on="CAMPAÑA", how="left").fillna({"Marcaciones":0})
    pivot["Marcaciones"] = pd.to_numeric(pivot["Marcaciones"], errors="coerce").fillna(0).astype(int)

    for col in ["No conecta", "Número equivocado", "Volver a contactar", "Efectivo"]:
        pivot[col] = pd.to_numeric(pivot[col], errors="coerce").fillna(0).astype(int)
        pivot[f"% {col}"] = pivot.apply(lambda r: round((r[col]/r["Marcaciones"])*100,2) if r["Marcaciones"] else 0.0, axis=1)

    camp_quality = pivot[["CAMPAÑA","Marcaciones",
                         "% No conecta","% Número equivocado","% Volver a contactar","% Efectivo"]].copy()

    # 9) Resumen ejecutivo
    resumen = pd.DataFrame([
        ["Telefonos únicos entregados (enero)", delivered_unique],
        ["Telefonos únicos tocados en enero (sobre entregas enero)", touched_unique],
        ["% Cobertura (tocados/entregados)", coverage_pct],
        ["Marcaciones sobre entregas enero", total_calls_on_jan_bases],
        ["Answered (sobre entregas enero)", answered_calls],
        ["Missed (sobre entregas enero)", missed_calls],
        ["% Contactabilidad (answered/marcaciones) sobre entregas enero", contactability_pct],
    ], columns=["Métrica","Valor"])

    # 10) Exportar
    os.makedirs(os.path.dirname(OUT_XLSX), exist_ok=True)
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen_entregas_enero", index=False)
        cov.sort_values("% cobertura", ascending=False).to_excel(writer, sheet_name="Cobertura_por_campaña", index=False)
        camp_perf.sort_values("% contactabilidad", ascending=False).to_excel(writer, sheet_name="Contactabilidad_por_campaña", index=False)
        camp_quality.sort_values("Marcaciones", ascending=False).to_excel(writer, sheet_name="Tag2_%_por_campaña", index=False)
        tag2_dist.to_excel(writer, sheet_name="Tag2_dist_entregas_enero", index=False)

    print("OK - Cruce entregas enero vs marcaciones enero")
    print(f"Delivered unique phones (enero): {delivered_unique:,}")
    print(f"Touched unique phones (enero): {touched_unique:,} | Coverage: {coverage_pct}%")
    print(f"Calls on jan deliveries: {total_calls_on_jan_bases:,} | Answered: {answered_calls:,} | Contactability: {contactability_pct}%")
    print(f"Output: {OUT_XLSX}")

if __name__ == "__main__":
    main()
