# -*- coding: utf-8 -*-
"""
05_seguimiento_entregas_vs_marcaciones_mix.py

Seguimiento de ENTREGAS (mes base) vs MARCACIONES (mes base + siguiente H1)
Unidad de análisis: CONTACTO (contact_id_base)
Un contacto se considera "tocado" si cualquiera de sus teléfonos hace match en marcaciones.

USO:
    python 05_seguimiento_entregas_vs_marcaciones_mix.py 2026-01 2026-01 2026-02_H1

Estructura esperada:
    ./Contactabilidad/out/<MES>/marcaciones_<MES>_consolidadas.xlsx
    ./Contactabilidad/out/<MES_BASE>/B_mes/base_contactos_ENTREGAS_<MES_BASE>.xlsx
    ./Contactabilidad/out/<MES_BASE>/B_mes/base_contactos_long_ENTREGAS_<MES_BASE>.xlsx

Outputs:
    ./Contactabilidad/out/<MES_BASE>/C_seguimiento/
        cruce_seguimiento_<MES_BASE>_VS_<P1>+<P2>_CONTACTOS.xlsx
        REPORTE_SEGUIMIENTO_<MES_BASE>_VS_<P1>+<P2>.xlsx
"""

import os
import re
import sys
import pandas as pd

# -------------------------
# Helpers
# -------------------------

def die(msg: str, code: int = 1):
    raise RuntimeError(msg)

def norm_str(x) -> str:
    if x is None:
        return ""
    if isinstance(x, pd.Series):
        x = x.iloc[0] if len(x) else ""
    if pd.isna(x):
        return ""
    return str(x).strip()

def normalize_phone_key(x) -> str:
    """
    Normaliza teléfonos para cruce:
    - quita espacios, guiones, paréntesis, puntos
    - deja solo dígitos y +
    """
    s = norm_str(x)
    if not s:
        return ""
    s = s.replace("(", "").replace(")", "").replace(" ", "").replace("-", "").replace(".", "")
    s = re.sub(r"[^0-9+]", "", s)
    return s

def first_existing_col(df: pd.DataFrame, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

def safe_nunique(s: pd.Series) -> int:
    if s is None or s.empty:
        return 0
    return int(s.dropna().nunique())

# -------------------------
# Main
# -------------------------

def main():
    if len(sys.argv) != 4:
        print("Uso: python 05_seguimiento_entregas_vs_marcaciones_mix.py MES_BASE MARC_P1 MARC_P2")
        print("Ej:  python 05_seguimiento_entregas_vs_marcaciones_mix.py 2026-01 2026-01 2026-02_H1")
        sys.exit(1)

    MES_BASE = sys.argv[1].strip()
    P1 = sys.argv[2].strip()
    P2 = sys.argv[3].strip()

    # -------------------------
    # Rutas
    # -------------------------
    OUT_BASE = f"./Contactabilidad/out/{MES_BASE}"
    OUT_P1 = f"./Contactabilidad/out/{P1}"
    OUT_P2 = f"./Contactabilidad/out/{P2}"

    MARC_P1_PATH = f"{OUT_P1}/marcaciones_{P1}_consolidadas.xlsx"
    MARC_P2_PATH = f"{OUT_P2}/marcaciones_{P2}_consolidadas.xlsx"

    BASE_CONTACTOS_PATH = f"{OUT_BASE}/B_mes/base_contactos_ENTREGAS_{MES_BASE}.xlsx"
    BASE_LONG_PATH = f"{OUT_BASE}/B_mes/base_contactos_long_ENTREGAS_{MES_BASE}.xlsx"

    OUT_FOLDER = f"{OUT_BASE}/C_seguimiento"
    ensure_dir(OUT_FOLDER)

    OUT_CRUCE = f"{OUT_FOLDER}/cruce_seguimiento_{MES_BASE}_VS_{P1}+{P2}_CONTACTOS.xlsx"
    OUT_REPORTE = f"{OUT_FOLDER}/REPORTE_SEGUIMIENTO_{MES_BASE}_VS_{P1}+{P2}.xlsx"

    print("\n==============================")
    print("SEGUIMIENTO ENTREGAS vs MARCACIONES")
    print(f"Base entregas (mes): {MES_BASE}")
    print(f"Marcaciones P1:      {P1} -> {MARC_P1_PATH}")
    print(f"Marcaciones P2:      {P2} -> {MARC_P2_PATH}")
    print(f"Base long:           {BASE_LONG_PATH}")
    print(f"Base contactos:      {BASE_CONTACTOS_PATH}")
    print(f"Salida cruce:        {OUT_CRUCE}")
    print(f"Salida reporte:      {OUT_REPORTE}")
    print("==============================\n")

    # Validaciones
    for fp in [MARC_P1_PATH, MARC_P2_PATH, BASE_CONTACTOS_PATH, BASE_LONG_PATH]:
        if not os.path.exists(fp):
            die(f"No existe archivo requerido: {fp}")

    # -------------------------
    # Leer entradas
    # -------------------------
    m1 = pd.read_excel(MARC_P1_PATH, dtype=str)
    m2 = pd.read_excel(MARC_P2_PATH, dtype=str)
    base_contacts = pd.read_excel(BASE_CONTACTOS_PATH, dtype=str)
    base_long = pd.read_excel(BASE_LONG_PATH, dtype=str)

    # -------------------------
    # Columnas esperadas / flexibles
    # -------------------------
    # Marcaciones: number y status
    MARC_NUMBER_COL = first_existing_col(m1, ["contact_number", "CONTACT_NUMBER", "number", "NUMBER"])
    if not MARC_NUMBER_COL:
        die(f"Marcaciones P1 no tiene columna de número. Busqué contact_number/CONTACT_NUMBER/number/NUMBER. cols={list(m1.columns)}")

    MARC_STATUS_COL = first_existing_col(m1, ["status", "STATUS"])
    if not MARC_STATUS_COL:
        die("Marcaciones no tiene columna 'status'")

    # Identificador de llamada para deduplicar (si hay repetidos)
    CALL_ID_COL = first_existing_col(m1, ["cdr_id", "CDR_ID", "uuid", "UUID", "call_id", "CALL_ID"])
    if not CALL_ID_COL:
        CALL_ID_COL = None

    # Columnas tags
    TAG2_COL = first_existing_col(m1, ["tag 2", "TAG 2", "tags 2", "TAGS 2", "tags estatus 2", "TAGS ESTATUS 2"])

    # IMR = agent_name (CORRECCIÓN)
    IMR_COL = first_existing_col(m1, ["agent_name", "AGENT_NAME", "agent", "AGENT"])
    # NOTA: si ninguna existe, la pestaña IMR_NO_MATCH no se generará.

    # En base_long necesitamos phone_key y contact_id_base
    if "phone_key" not in base_long.columns:
        die(f"Base long no tiene 'phone_key'. cols={list(base_long.columns)}")

    CID_COL_BL = first_existing_col(base_long, ["contact_id_base", "CONTACT_ID_BASE", "contact_id", "CONTACT_ID"])
    if not CID_COL_BL:
        die(f"Base long no tiene columna de id de contacto. Busqué contact_id_base/contact_id. cols={list(base_long.columns)}")

    # En base_contactos necesitamos contact_id_base y HERRAMIENTA/CAMPAÑA
    CID_COL_BC = first_existing_col(base_contacts, ["contact_id_base", "CONTACT_ID_BASE", "contact_id", "CONTACT_ID"])
    if not CID_COL_BC:
        die(f"Base contactos no tiene id. Busqué contact_id_base/contact_id. cols={list(base_contacts.columns)}")

    HERR_COL = first_existing_col(base_contacts, ["HERRAMIENTA", "herramienta"])
    CAMPA_ENT_COL = first_existing_col(base_contacts, ["CAMPAÑA", "CAMPANA", "CAMPAÑA_STD", "CAMPANA_STD"])

    # -------------------------
    # Unir marcaciones P1 + P2
    # -------------------------
    m1["__period__"] = P1
    m2["__period__"] = P2
    marc = pd.concat([m1, m2], ignore_index=True, sort=False)

    # Dedup por id si existe
    if CALL_ID_COL and CALL_ID_COL in marc.columns:
        before = len(marc)
        marc = marc.drop_duplicates(subset=[CALL_ID_COL])
        print(f"Deduplicación: {before:,} -> {len(marc):,} (usando {CALL_ID_COL})")

    # Crear phone_key desde contact_number
    marc["phone_key"] = marc[MARC_NUMBER_COL].apply(normalize_phone_key)
    marc.loc[marc["phone_key"] == "", "phone_key"] = pd.NA

    # -------------------------
    # Cruce: marcaciones vs base_long
    # -------------------------
    bl = base_long[["phone_key", CID_COL_BL]].copy()
    bl = bl.rename(columns={CID_COL_BL: "contact_id_base"})
    bl = bl.dropna(subset=["phone_key"])

    jm = marc.merge(
        bl.drop_duplicates(subset=["phone_key", "contact_id_base"]),
        on="phone_key",
        how="left",
        suffixes=("", "_bl"),
        indicator=True
    )

    jm["has_match"] = jm["contact_id_base"].notna().astype(int)

    # match_type: UNICO vs MULTIPLE por phone_key
    match_counts = (
        jm[jm["has_match"] == 1]
        .groupby("phone_key")["contact_id_base"]
        .nunique()
        .reset_index()
        .rename(columns={"contact_id_base": "match_count"})
    )
    jm = jm.merge(match_counts, on="phone_key", how="left")
    jm["match_type"] = jm["match_count"].apply(
        lambda x: "UNICO" if pd.notna(x) and int(x) == 1 else ("MULTIPLE" if pd.notna(x) and int(x) > 1 else pd.NA)
    )

    # -------------------------
    # Enriquecer con base_contactos: CAMPAÑA / HERRAMIENTA
    # -------------------------
    bc_cols = [CID_COL_BC]
    if CAMPA_ENT_COL: bc_cols.append(CAMPA_ENT_COL)
    if HERR_COL: bc_cols.append(HERR_COL)

    bc = base_contacts[bc_cols].copy()
    bc = bc.rename(columns={CID_COL_BC: "contact_id_base"})
    if CAMPA_ENT_COL and CAMPA_ENT_COL != "CAMPAÑA":
        bc = bc.rename(columns={CAMPA_ENT_COL: "CAMPAÑA"})
    if HERR_COL and HERR_COL != "HERRAMIENTA":
        bc = bc.rename(columns={HERR_COL: "HERRAMIENTA"})

    bc = bc.drop_duplicates(subset=["contact_id_base"])
    jm = jm.merge(bc, on="contact_id_base", how="left")

    # -------------------------
    # Guardar cruce detallado
    # -------------------------
    jm.to_excel(OUT_CRUCE, index=False)

    # -------------------------
    # REPORTE (enfocado a seguimiento entregas MES_BASE)
    # -------------------------
    delivered_contacts = safe_nunique(base_contacts[CID_COL_BC])
    touched_contacts = safe_nunique(jm.loc[jm["has_match"] == 1, "contact_id_base"])
    coverage = round((touched_contacts / delivered_contacts) * 100, 2) if delivered_contacts else 0.0

    calls_total = len(jm)
    calls_on_deliveries = int(jm["has_match"].sum())

    # Answered/Missed SOLO sobre llamadas en entregas base
    jm_match = jm[jm["has_match"] == 1].copy()
    answered_on_del = int((jm_match[MARC_STATUS_COL].str.lower() == "answered").sum())
    missed_on_del = int((jm_match[MARC_STATUS_COL].str.lower() == "missed").sum())
    contactability_on_del = round((answered_on_del / (answered_on_del + missed_on_del)) * 100, 2) if (answered_on_del + missed_on_del) else 0.0

    resumen = pd.DataFrame([{
        "MES_BASE_ENTREGAS": MES_BASE,
        "MARCACIONES_P1": P1,
        "MARCACIONES_P2": P2,
        "MARCACIONES_TOTAL_COMBINADAS": calls_total,
        "MARCACIONES_EN_ENTREGAS_BASE (has_match=1)": calls_on_deliveries,
        "%_MARCACIONES_EN_ENTREGAS_BASE": round((calls_on_deliveries / calls_total) * 100, 2) if calls_total else 0.0,
        "CONTACTOS_ENTREGADOS_BASE": delivered_contacts,
        "CONTACTOS_TOCADOS_BASE": touched_contacts,
        "%_COBERTURA_CONTACTOS (tocados/entregados)": coverage,
        "ANSWERED_EN_ENTREGAS": answered_on_del,
        "MISSED_EN_ENTREGAS": missed_on_del,
        "%_CONTACTABILIDAD_EN_ENTREGAS (answered/(answered+missed))": contactability_on_del
    }])

    # Cobertura por campaña (entregas base)
    if "CAMPAÑA" in bc.columns:
        delivered_by_camp = bc.groupby("CAMPAÑA")["contact_id_base"].nunique().reset_index().rename(columns={"contact_id_base":"Contactos_entregados"})
        touched_by_camp = jm_match.groupby("CAMPAÑA")["contact_id_base"].nunique().reset_index().rename(columns={"contact_id_base":"Contactos_tocados"})
        calls_by_camp = jm_match.groupby("CAMPAÑA").size().reset_index(name="Marcaciones")
        ans_by_camp = jm_match.assign(_ans=(jm_match[MARC_STATUS_COL].str.lower()=="answered").astype(int))\
                             .groupby("CAMPAÑA")["_ans"].sum().reset_index().rename(columns={"_ans":"answered"})
        mis_by_camp = jm_match.assign(_mis=(jm_match[MARC_STATUS_COL].str.lower()=="missed").astype(int))\
                             .groupby("CAMPAÑA")["_mis"].sum().reset_index().rename(columns={"_mis":"missed"})

        camp = delivered_by_camp.merge(touched_by_camp, on="CAMPAÑA", how="left")\
                                .merge(calls_by_camp, on="CAMPAÑA", how="left")\
                                .merge(ans_by_camp, on="CAMPAÑA", how="left")\
                                .merge(mis_by_camp, on="CAMPAÑA", how="left")

        camp = camp.fillna(0)
        camp["%_tocado"] = camp.apply(lambda r: round((r["Contactos_tocados"]/r["Contactos_entregados"])*100,2) if r["Contactos_entregados"] else 0.0, axis=1)
        camp["%_contactabilidad_en_entregas"] = camp.apply(lambda r: round((r["answered"]/(r["answered"]+r["missed"]))*100,2) if (r["answered"]+r["missed"]) else 0.0, axis=1)
        camp = camp.sort_values(by=["%_tocado","Contactos_tocados"], ascending=False)
    else:
        camp = pd.DataFrame()

    # Herramientas (desde entregas base) => solo jm_match
    if "HERRAMIENTA" in bc.columns:
        delivered_by_tool = bc.groupby("HERRAMIENTA")["contact_id_base"].nunique().reset_index().rename(columns={"contact_id_base":"Contactos_entregados"})
        touched_by_tool = jm_match.groupby("HERRAMIENTA")["contact_id_base"].nunique().reset_index().rename(columns={"contact_id_base":"Contactos_tocados"})
        calls_by_tool = jm_match.groupby("HERRAMIENTA").size().reset_index(name="Marcaciones")
        ans_by_tool = jm_match.assign(_ans=(jm_match[MARC_STATUS_COL].str.lower()=="answered").astype(int))\
                              .groupby("HERRAMIENTA")["_ans"].sum().reset_index().rename(columns={"_ans":"answered"})
        mis_by_tool = jm_match.assign(_mis=(jm_match[MARC_STATUS_COL].str.lower()=="missed").astype(int))\
                              .groupby("HERRAMIENTA")["_mis"].sum().reset_index().rename(columns={"_mis":"missed"})

        tools = delivered_by_tool.merge(touched_by_tool, on="HERRAMIENTA", how="left")\
                                 .merge(calls_by_tool, on="HERRAMIENTA", how="left")\
                                 .merge(ans_by_tool, on="HERRAMIENTA", how="left")\
                                 .merge(mis_by_tool, on="HERRAMIENTA", how="left")

        tools = tools.fillna(0)
        tools["%_tocado"] = tools.apply(lambda r: round((r["Contactos_tocados"]/r["Contactos_entregados"])*100,2) if r["Contactos_entregados"] else 0.0, axis=1)
        tools["%_contactabilidad_en_entregas"] = tools.apply(lambda r: round((r["answered"]/(r["answered"]+r["missed"]))*100,2) if (r["answered"]+r["missed"]) else 0.0, axis=1)
        tools = tools.sort_values(by=["%_tocado","Contactos_tocados"], ascending=False)
    else:
        tools = pd.DataFrame()

    # Tags (Tag 2) SOLO sobre jm_match
    if TAG2_COL:
        tag2 = jm_match[TAG2_COL].fillna("SIN_TAG_2").astype(str).str.strip()
        tag2_dist = tag2.value_counts().reset_index()
        tag2_dist.columns = ["tag2", "Marcaciones"]
        tag2_dist["%"] = tag2_dist["Marcaciones"].apply(lambda x: round((x / calls_on_deliveries) * 100, 2) if calls_on_deliveries else 0.0)
    else:
        tag2_dist = pd.DataFrame()

    # IMR (agent_name) por NO MATCH (control de calidad)
    if IMR_COL:
        jm_nomatch = jm[jm["has_match"] == 0].copy()
        imr_nm = jm_nomatch[IMR_COL].fillna("SIN_AGENT_NAME").astype(str).str.strip()
        imr_nomatch = imr_nm.value_counts().reset_index()
        imr_nomatch.columns = ["agent_name (IMR)", "Marcaciones_sin_match"]
    else:
        imr_nomatch = pd.DataFrame()

    # -------------------------
    # Export REPORTE
    # -------------------------
    with pd.ExcelWriter(OUT_REPORTE, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)

        if not camp.empty:
            camp.to_excel(writer, sheet_name="COBERTURA_x_CAMPAÑA", index=False)

        if not tools.empty:
            tools.to_excel(writer, sheet_name="HERRAMIENTAS", index=False)

        if not tag2_dist.empty:
            tag2_dist.to_excel(writer, sheet_name="TAG2", index=False)

        if not imr_nomatch.empty:
            imr_nomatch.to_excel(writer, sheet_name="IMR_NO_MATCH", index=False)

        control = pd.DataFrame([{
            "calls_total_combinadas": calls_total,
            "calls_en_entregas_base": calls_on_deliveries,
            "contactos_entregados_base": delivered_contacts,
            "contactos_tocados_base": touched_contacts,
            "%_cobertura_contactos": coverage,
        }])
        control.to_excel(writer, sheet_name="CONTROL", index=False)

    print("OK - Seguimiento generado")
    print(f"Contactos entregados (base): {delivered_contacts:,}")
    print(f"Contactos tocados (base): {touched_contacts:,} | Cobertura: {coverage:.2f}%")
    print(f"Marcaciones combinadas: {calls_total:,}")
    print(f"Marcaciones EN entregas base: {calls_on_deliveries:,} | Contactabilidad en entregas: {contactability_on_del:.2f}%")
    print(f"Output cruce:   {OUT_CRUCE}")
    print(f"Output reporte: {OUT_REPORTE}")

if __name__ == "__main__":
    main()