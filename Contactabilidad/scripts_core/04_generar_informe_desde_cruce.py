# -*- coding: utf-8 -*-
"""
04_generar_informe_desde_cruce.py

Builds the Excel report from the output produced by:
03_cruce_marcaciones_vs_base_contactos.py

Usage:
  python 04_generar_informe_desde_cruce.py YYYY-MM GLOBAL
  python 04_generar_informe_desde_cruce.py YYYY-MM MES

Outputs:
  ./Contactabilidad/out/YYYY-MM/<A_historico|B_mes>/REPORTE_BD_MARCACIONES_<YYYY-MM>_<GLOBAL|MES>.xlsx
"""

import os
import sys
import re
import pandas as pd


# -----------------------------
# REQUIRED COLUMNS IN CRUCE FILE
# -----------------------------
COL_STATUS = "status"
COL_CAMPANA = "CAMPAÑA"
COL_TAG2 = "tags estatus 1"
COL_CONTACT_ID = "contact_id_base"
COL_HAS_MATCH = "has_match"
COL_MATCH_TYPE = "match_type"

# Optional tool columns (if present in the merge)
TOOL_COLS = ["LUSHA", "SQL", "SH", "APL"]


def ensure_exists(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(f"No existe: {path}")


def norm_text(x):
    if x is None or (isinstance(x, float) and pd.isna(x)) or pd.isna(x):
        return ""
    s = str(x).strip()
    s = " ".join(s.split())
    return s


def normalize_campana(s):
    s = norm_text(s)
    if s == "" or s.lower() in {"nan", "none"}:
        return "(SIN TAG DE CAMPAÑA)"
    return s


def is_truthy_tool(v):
    """Tool flags may come as '1', 'SI', 'X', 'TRUE', etc."""
    if v is None or pd.isna(v):
        return False
    s = str(v).strip().upper()
    return s in {"1", "SI", "SÍ", "X", "TRUE", "VERDADERO", "YES"}


def infer_tool(row):
    """Returns one label among TOOL_COLS or SIN_HERRAMIENTA."""
    present = []
    for c in TOOL_COLS:
        if c in row.index and is_truthy_tool(row[c]):
            present.append(c)
    if not present:
        return "SIN_HERRAMIENTA"
    # If multiple tools marked, keep a combined label for transparency
    return "+".join(present)


def find_cruce_file(out_folder: str, mes: str, tipo: str) -> str:
    """
    We try the conventional name first.
    If not found, we search for any xlsx containing 'cruce' + mes + tipo + 'CONTACTOS'.
    """
    # Most common convention we used:
    # cruce_<YYYY-MM>_<TIPO>_CONTACTOS.xlsx
    cand = os.path.join(out_folder, f"cruce_{mes}_{tipo}_CONTACTOS.xlsx")
    if os.path.exists(cand):
        return cand

    # Fallback search
    for fn in os.listdir(out_folder):
        fnu = fn.upper()
        if fn.endswith(".xlsx") and ("CRUCE" in fnu) and (mes in fn) and (tipo in fnu) and ("CONTACTOS" in fnu):
            return os.path.join(out_folder, fn)

    raise FileNotFoundError(
        f"No encontré archivo de cruce en: {out_folder}. "
        f"Esperaba algo como 'cruce_{mes}_{tipo}_CONTACTOS.xlsx'."
    )


def find_base_contactos_file(out_folder: str, mes: str, tipo: str) -> str:
    """
    Reads base_contactos_*.xlsx to compute delivered contacts + coverage.
    """
    if tipo == "GLOBAL":
        cand = os.path.join(out_folder, "base_contactos_GLOBAL.xlsx")
        if os.path.exists(cand):
            return cand

        # fallback
        for fn in os.listdir(out_folder):
            if fn.lower().startswith("base_contactos_") and "global" in fn.lower() and fn.endswith(".xlsx"):
                return os.path.join(out_folder, fn)

    else:
        # base_contactos_ENTREGAS_YYYY-MM.xlsx (your builder prints that)
        cand = os.path.join(out_folder, f"base_contactos_ENTREGAS_{mes}.xlsx")
        if os.path.exists(cand):
            return cand

        for fn in os.listdir(out_folder):
            if fn.lower().startswith("base_contactos_") and ("entregas" in fn.lower()) and (mes in fn) and fn.endswith(".xlsx"):
                return os.path.join(out_folder, fn)

    raise FileNotFoundError(
        f"No encontré base_contactos en: {out_folder} "
        f"(necesaria para cobertura tocados/entregados)."
    )


def safe_div(a, b):
    return round((a / b) * 100, 2) if b else 0.0


def main():
    if len(sys.argv) != 3:
        print("Uso: python 04_generar_informe_desde_cruce.py YYYY-MM (GLOBAL|MES)")
        sys.exit(1)

    mes = sys.argv[1].strip()
    tipo = sys.argv[2].upper().strip()

    if not re.match(r"^\d{4}-\d{2}$", mes):
        raise ValueError("El primer argumento debe ser YYYY-MM, ej: 2026-01")

    if tipo not in {"GLOBAL", "MES"}:
        raise ValueError("El segundo argumento debe ser GLOBAL o MES")

    # Folder convention
    out_root = f"./Contactabilidad/out/{mes}"
    out_folder = os.path.join(out_root, "A_historico" if tipo == "GLOBAL" else "B_mes")
    ensure_exists(out_folder)

    cruce_path = find_cruce_file(out_folder, mes, tipo)
    base_contactos_path = find_base_contactos_file(out_folder, mes, tipo)

    print("\n==============================")
    print(f"Periodo: {mes}")
    print(f"Tipo base: {tipo}")
    print(f"Cruce:      {cruce_path}")
    print(f"Base cont.: {base_contactos_path}")
    print("==============================\n")

    # Load data
    df = pd.read_excel(cruce_path, dtype=str)
    base = pd.read_excel(base_contactos_path, dtype=str)

    # Validate required columns in cruce
    required = [COL_STATUS, COL_CAMPANA, COL_TAG2, COL_CONTACT_ID, COL_HAS_MATCH, COL_MATCH_TYPE]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise RuntimeError(f"Faltan columnas en el cruce: {missing}")

    # Normalize key fields
    df[COL_STATUS] = df[COL_STATUS].str.lower().str.strip()
    df[COL_CAMPANA] = df[COL_CAMPANA].apply(normalize_campana)
    df[COL_HAS_MATCH] = pd.to_numeric(df[COL_HAS_MATCH], errors="coerce").fillna(0).astype(int)

    # Subsets
    df_match = df[df[COL_HAS_MATCH] == 1].copy()

    # --- GLOBAL KPIs ---
    total_calls = len(df)
    answered = int((df[COL_STATUS] == "answered").sum())
    missed = int((df[COL_STATUS] == "missed").sum())
    contactability = safe_div(answered, (answered + missed))  # answered/(answered+missed)

    calls_with_match = int(df[COL_HAS_MATCH].sum())
    pct_match_calls = safe_div(calls_with_match, total_calls)

    match_unique = int((df[COL_MATCH_TYPE] == "UNICO").sum())
    match_multi = int((df[COL_MATCH_TYPE] == "MULTIPLE").sum())

    # Coverage by CONTACT (delivered vs touched)
    # Delivered contacts = all unique contact_ids in base_contactos file
    # Touched contacts = unique contact_id_base found in matched calls
    delivered_contacts = base["contact_id"].nunique() if "contact_id" in base.columns else base.shape[0]
    touched_contacts = df_match[COL_CONTACT_ID].dropna().nunique()
    coverage_contacts = safe_div(touched_contacts, delivered_contacts)

    # --- STATUS TABLE ---
    status_tbl = (
        df.groupby(COL_STATUS)
        .size()
        .reset_index(name="Conteo")
        .sort_values("Conteo", ascending=False)
    )
    status_tbl["%"] = status_tbl["Conteo"].apply(lambda x: safe_div(x, total_calls))

    # --- TAG2 TABLE (on matched calls) ---
    tag2_tbl = (
        df_match.groupby(COL_TAG2)
        .size()
        .reset_index(name="Conteo")
        .sort_values("Conteo", ascending=False)
    )
    total_tag2 = int(tag2_tbl["Conteo"].sum()) if len(tag2_tbl) else 0
    tag2_tbl["%"] = tag2_tbl["Conteo"].apply(lambda x: safe_div(x, total_tag2))

    # --- PER CAMPAIGN (contacts touched + contactability) ---
    # We compute campaign stats on matched calls (because campaign comes from base side).
    per_campaign = (
        df_match.groupby(COL_CAMPANA)
        .agg(
            Contactos_tocados_unicos=(COL_CONTACT_ID, lambda s: s.dropna().nunique()),
            Marcaciones=(COL_STATUS, "size"),
            answered=(COL_STATUS, lambda s: (s == "answered").sum()),
            missed=(COL_STATUS, lambda s: (s == "missed").sum()),
        )
        .reset_index()
    )
    per_campaign["% contactabilidad"] = per_campaign.apply(
        lambda r: safe_div(int(r["answered"]), int(r["answered"]) + int(r["missed"])), axis=1
    )
    per_campaign = per_campaign.sort_values("Contactos_tocados_unicos", ascending=False)

    # --- COVERAGE BY CAMPAIGN (delivered vs touched) ---
    # Delivered by campaign: from base_contactos (CAMPAÑA column might exist)
    # If not, we can’t compute this section.
    if COL_CAMPANA in base.columns:
        base_cov = base.copy()
        base_cov[COL_CAMPANA] = base_cov[COL_CAMPANA].apply(normalize_campana)

        delivered_by_c = (
            base_cov.groupby(COL_CAMPANA)["contact_id"]
            .nunique()
            .reset_index()
            .rename(columns={"contact_id": "Contactos_entregados"})
        )

        touched_by_c = (
            df_match.groupby(COL_CAMPANA)[COL_CONTACT_ID]
            .nunique()
            .reset_index()
            .rename(columns={COL_CONTACT_ID: "Contactos_tocados"})
        )

        cov_campaign = delivered_by_c.merge(touched_by_c, on=COL_CAMPANA, how="left")
        cov_campaign["Contactos_tocados"] = cov_campaign["Contactos_tocados"].fillna(0).astype(int)
        cov_campaign["%_tocado"] = cov_campaign.apply(
            lambda r: safe_div(int(r["Contactos_tocados"]), int(r["Contactos_entregados"])),
            axis=1
        )
        cov_campaign = cov_campaign.sort_values("%_tocado", ascending=False)
    else:
        cov_campaign = pd.DataFrame(columns=[COL_CAMPANA, "Contactos_entregados", "Contactos_tocados", "%_tocado"])

    # --- TOOLS (if tool columns exist) ---
    if any(c in df_match.columns for c in TOOL_COLS):
        df_match["HERRAMIENTA"] = df_match.apply(infer_tool, axis=1)

        tools_tbl = (
            df_match.groupby("HERRAMIENTA")
            .agg(
                Contactos_tocados_unicos=(COL_CONTACT_ID, lambda s: s.dropna().nunique()),
                Marcaciones=(COL_STATUS, "size"),
                answered=(COL_STATUS, lambda s: (s == "answered").sum()),
                missed=(COL_STATUS, lambda s: (s == "missed").sum()),
            )
            .reset_index()
        )
        tools_tbl["% contactabilidad"] = tools_tbl.apply(
            lambda r: safe_div(int(r["answered"]), int(r["answered"]) + int(r["missed"])), axis=1
        )
        tools_tbl = tools_tbl.sort_values("Marcaciones", ascending=False)
    else:
        tools_tbl = pd.DataFrame(columns=["HERRAMIENTA", "Contactos_tocados_unicos", "Marcaciones", "answered", "missed", "% contactabilidad"])

    # --- SUMMARY SHEET ---
    resumen = pd.DataFrame(
        [
            ["Marcaciones totales", total_calls],
            ["answered", answered],
            ["missed", missed],
            ["% contactabilidad (answered/(answered+missed))", contactability],
            ["Marcaciones con match", calls_with_match],
            ["% match (marcaciones con match / total)", pct_match_calls],
            ["Match único (marcaciones)", match_unique],
            ["Match múltiple (marcaciones)", match_multi],
            ["Contactos entregados (base)", delivered_contacts],
            ["Contactos tocados (≥1 tel marcado)", touched_contacts],
            ["Cobertura contactos (tocados/entregados)", coverage_contacts],
        ],
        columns=["Métrica", "Valor"],
    )

    # Output path
    out_path = os.path.join(out_folder, f"REPORTE_BD_MARCACIONES_{mes}_{tipo}.xlsx")

    # Write Excel
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        status_tbl.to_excel(writer, sheet_name="Status", index=False)
        per_campaign.to_excel(writer, sheet_name="Campañas", index=False)
        cov_campaign.to_excel(writer, sheet_name="Cobertura_Campaña", index=False)
        tools_tbl.to_excel(writer, sheet_name="Herramientas", index=False)
        tag2_tbl.to_excel(writer, sheet_name="Tag2", index=False)

        # (Optional) keep raw matched calls for audit
        df_match.to_excel(writer, sheet_name="Cruce_Matched_Raw", index=False)

    print("OK - Reporte generado desde cruce")
    print(f"Salida: {out_path}")
    print(f"Marcaciones: {total_calls:,} | Answered: {answered:,} | Missed: {missed:,} | Contactabilidad: {contactability:.2f}%")
    print(f"Con match: {calls_with_match:,} ({pct_match_calls:.2f}%) | Único: {match_unique:,} | Múltiple: {match_multi:,}")
    print(f"Contactos entregados: {delivered_contacts:,} | Contactos tocados: {touched_contacts:,} | Cobertura: {coverage_contacts:.2f}%\n")


if __name__ == "__main__":
    main()