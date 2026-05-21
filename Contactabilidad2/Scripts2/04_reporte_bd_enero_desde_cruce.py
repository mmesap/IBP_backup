import os
import re
import pandas as pd
import numpy as np

# =========================================================
# CONFIG (AJUSTA SOLO ESTO)
# =========================================================
MARCACIONES_XLSX = "./Contactabilidad/Data/marcaciones/Marcaciones enero.xlsx"
BASE_ALL_CLEAN_XLSX = "./Contactabilidad/out/entregas_consolidadas_clean.xlsx"
BASE_ENERO_CLEAN_XLSX = "./Contactabilidad/out/bases_enero_consolidado_clean.xlsx"
OUT_REPORT_XLSX = "./Contactabilidad/out/REPORTE_BD_MARCACIONES_ENERO_2026_CONTACTOS_FULL.xlsx"
# =========================================================


# ---------------------------
# Helpers
# ---------------------------

def phone_key(x):
    if x is None or pd.isna(x):
        return None
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.replace("+", "").replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
    s = re.sub(r"\D+", "", s)
    return s if s else None

def norm_col(c: str) -> str:
    return str(c).strip()

def norm_text(x) -> str:
    if x is None or pd.isna(x):
        return ""
    s = str(x).strip().lower()
    s = " ".join(s.split())
    s = s.replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u").replace("ñ","n")
    return s

def safe_pct(num, den):
    return round((num/den)*100, 2) if den else 0.0

def read_all_sheets_xlsx(path: str) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    dfs = []
    for sh in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sh, dtype=str)
        df["__sheet__"] = sh
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

def find_first_existing(cols, candidates):
    for cand in candidates:
        if cand in cols:
            return cand
    return None

def detect_campaign_col(df: pd.DataFrame) -> str:
    cols = list(df.columns)
    for cand in ["CAMPAÑA", "Campaña", "CAMPANA", "Campana", "campaña", "campana"]:
        if cand in cols:
            return cand
    raise KeyError(f"No se encontró columna de campaña. Columnas: {cols}")

def flag_is_true(x) -> bool:
    s = norm_text(x)
    return s in {"si","sí","yes","y","1","true","x","ok","vale"} or s.startswith("si ")

def detect_tool_cols(df: pd.DataFrame):
    """
    Detecta columnas de herramientas (más conservador para evitar basura).
    Ajusta aquí si quieres incluir más.
    """
    tool_candidates = ["LUSHA", "SH", "SQL", "APL", "APOLLO", "SALESQL", "SALESQB", "SALES QB"]
    cols = list(df.columns)
    found = []
    upper_map = {c: str(c).strip().upper() for c in cols}
    targets = set([t.upper() for t in tool_candidates])

    for c in cols:
        if upper_map[c] in targets:
            found.append(c)

    # únicos
    out = []
    seen = set()
    for c in found:
        if c not in seen:
            out.append(c)
            seen.add(c)
    return out


# ---------------------------
# Build base (CONTACTOS)
# ---------------------------

def build_base_contacts_and_long(base_df: pd.DataFrame, campaign_col: str, tool_cols: list):
    """
    base_contacts: 1 fila por contacto (contact_id)
    base_long: 1 fila por (contact_id, phone_key) para cruce por teléfono
    """
    b = base_df.copy().reset_index(drop=True)

    # contact_id interno
    b["contact_id"] = np.arange(len(b), dtype=int)

    # SOLO 3 teléfonos
    phone_cols = ["key_celular", "key_tel1", "key_tel2"]
    for c in phone_cols:
        if c not in b.columns:
            b[c] = None

    base_contacts = b.copy()
    base_contacts["CAMPAÑA_STD"] = base_contacts[campaign_col].fillna("").astype(str).str.strip()

    # Explode a long (por teléfono)
    rows = []
    for _, r in base_contacts.iterrows():
        keys = []
        for kc in phone_cols:
            k = phone_key(r.get(kc, None))
            if k:
                keys.append(k)
        keys = list(set(keys))
        for k in keys:
            row = {
                "contact_id": int(r["contact_id"]),
                "phone_key": k,
                "CAMPAÑA_STD": r["CAMPAÑA_STD"],
            }
            # cargar flags de herramientas a nivel contacto
            for tc in tool_cols:
                row[tc] = r.get(tc, None)
            rows.append(row)

    base_long = pd.DataFrame(rows)

    return base_contacts, base_long


def explode_tools_contacts(df_contacts: pd.DataFrame, tool_cols: list) -> pd.DataFrame:
    """
    Para atribución por herramienta a nivel CONTACTO:
    - Un contacto puede tener múltiples herramientas → se duplica.
    """
    if not tool_cols:
        out = df_contacts[["contact_id"]].copy()
        out["HERRAMIENTA"] = "SIN_HERRAMIENTA"
        return out

    rows = []
    for _, r in df_contacts.iterrows():
        tools = []
        for c in tool_cols:
            if c in df_contacts.columns and flag_is_true(r.get(c, "")):
                tools.append(str(c).strip())
        if not tools:
            tools = ["SIN_HERRAMIENTA"]
        for t in tools:
            rows.append({"contact_id": int(r["contact_id"]), "HERRAMIENTA": t})
    return pd.DataFrame(rows).drop_duplicates()


# ---------------------------
# Report builder (FULL)
# ---------------------------

def report_full_contacts(base_clean_path: str, marcaciones_path: str, label: str):
    # Load base
    b = pd.read_excel(base_clean_path, dtype=str)
    b.columns = [norm_col(c) for c in b.columns]
    campaign_col = detect_campaign_col(b)
    tool_cols = detect_tool_cols(b)

    base_contacts, base_long = build_base_contacts_and_long(b, campaign_col, tool_cols)

    delivered_contacts_total = int(base_contacts["contact_id"].nunique())

    # Load marcaciones (all sheets)
    m = read_all_sheets_xlsx(marcaciones_path)
    m.columns = [norm_col(c) for c in m.columns]

    contact_col = find_first_existing(m.columns, ["contact_number", "contact number", "Contact Number", "CONTACT_NUMBER", "contactNumber"])
    status_col  = find_first_existing(m.columns, ["status", "Status", "STATUS"])
    tag2_col    = find_first_existing(m.columns, ["tag 2", "Tag 2", "TAG 2", "tags 2", "tags 2 ", "tag2", "TAG2", "tag_2"])
    agent_col   = find_first_existing(m.columns, ["agent_name", "agent", "Agent", "AGENT", "usuario", "User", "IMR", "operador", "agente"])

    if contact_col is None or status_col is None:
        raise KeyError(f"Marcaciones debe tener contact_number y status. Encontrado: contact={contact_col}, status={status_col}")

    m["phone_key"] = m[contact_col].apply(phone_key)
    m = m.dropna(subset=["phone_key"]).copy()
    m["call_id"] = np.arange(len(m), dtype=int)

    # Totales globales
    total_calls = int(len(m))
    answered_total = int((m[status_col].astype(str).str.lower() == "answered").sum())
    missed_total = int((m[status_col].astype(str).str.lower() == "missed").sum())
    contactability_total = safe_pct(answered_total, total_calls)

    # Merge por phone_key
    j = m.merge(
        base_long,
        on="phone_key",
        how="left",
        indicator=True,
        suffixes=("_call", "_base"),
    )

    # Si el archivo de marcaciones trae contact_id propio, lo renombramos para no confundir
    if "contact_id_base" in j.columns:
        j = j.rename(columns={"contact_id_base": "contact_id"})
    elif "contact_id_y" in j.columns:
        j = j.rename(columns={"contact_id_y": "contact_id"})

    if "contact_id_call" in j.columns:
        # dejamos por trazabilidad pero NO lo usamos
        pass

    if "contact_id" not in j.columns:
        raise RuntimeError(f"[{label}] No quedó contact_id (base) tras el merge. Columns: {list(j.columns)}")

    j["has_match"] = (j["_merge"] == "both").astype(int)
    j["status_norm"] = j[status_col].astype(str).str.lower()
    j["tag2"] = j[tag2_col].astype(str).fillna("").astype(str) if tag2_col else ""

    # Calls con match (al menos un contacto)
    calls_with_match = int(j.groupby("call_id")["has_match"].max().sum())
    match_pct = safe_pct(calls_with_match, total_calls)

    jm = j[j["has_match"] == 1].copy()

    # Match único vs múltiple (por call_id, conteo de contactos distintos)
    if jm.empty:
        unique_match_calls = 0
        multi_match_calls = 0
        touched_contacts_total = 0
    else:
        per_call_contacts = jm.groupby("call_id")["contact_id"].nunique()
        unique_match_calls = int((per_call_contacts == 1).sum())
        multi_match_calls = int((per_call_contacts > 1).sum())
        touched_contacts_total = int(jm["contact_id"].nunique())

    # Status distribution (todas las llamadas, como antes)
    status_counts = (
        m[status_col].astype(str).str.lower()
        .value_counts(dropna=False)
        .reset_index()
    )
    status_counts.columns = ["Status", "Conteo"]
    status_counts["%"] = (status_counts["Conteo"] / total_calls * 100).round(2)

    # Status x Tag2 TOP200 (solo matched)
    if jm.empty:
        st_tag_top = pd.DataFrame(columns=["status_norm", "tag2", "Conteo", "%"])
    else:
        st_tag = (
            jm.groupby(["status_norm", "tag2"])
            .size()
            .reset_index(name="Conteo")
            .sort_values("Conteo", ascending=False)
        )
        st_tag["%"] = (st_tag["Conteo"] / len(jm) * 100).round(2)
        st_tag_top = st_tag.head(200)

    # ==============================
    # Por CAMPAÑA (FULL)
    # ==============================
    if jm.empty:
        camp_full = pd.DataFrame(columns=["CAMPAÑA","Contactos_tocados_unicos","%_sobre_contactos_tocados","answered","missed","Marcaciones","% contactabilidad"])
        cov_camp = pd.DataFrame(columns=["CAMPAÑA","Contactos_entregados","Contactos_tocados","%_tocado"])
        camp_tag2 = pd.DataFrame(columns=["CAMPAÑA","tag2","Conteo","Total_camp","%"])
    else:
        jm["CAMPAÑA"] = jm["CAMPAÑA_STD"].fillna("").astype(str)

        touched_by_camp = jm.groupby("CAMPAÑA")["contact_id"].nunique().reset_index(name="Contactos_tocados_unicos")
        touched_total_contacts = int(jm["contact_id"].nunique())
        touched_by_camp["%_sobre_contactos_tocados"] = np.where(
            touched_total_contacts > 0,
            (touched_by_camp["Contactos_tocados_unicos"] / touched_total_contacts) * 100,
            0.0
        )
        touched_by_camp["%_sobre_contactos_tocados"] = touched_by_camp["%_sobre_contactos_tocados"].round(2)

        # Contactabilidad por campaña: answered/(answered+missed) sobre matches atribuidos a esa campaña
        camp_status = (
            jm.groupby(["CAMPAÑA", "status_norm"])
            .size()
            .reset_index(name="Conteo")
        )
        camp_pivot = camp_status.pivot_table(index="CAMPAÑA", columns="status_norm", values="Conteo", aggfunc="sum", fill_value=0).reset_index()
        if "answered" not in camp_pivot.columns: camp_pivot["answered"] = 0
        if "missed" not in camp_pivot.columns: camp_pivot["missed"] = 0
        camp_pivot["Marcaciones"] = camp_pivot["answered"] + camp_pivot["missed"]
        camp_pivot["% contactabilidad"] = np.where(
            camp_pivot["Marcaciones"] > 0,
            (camp_pivot["answered"] / camp_pivot["Marcaciones"]) * 100,
            0.0
        )
        camp_pivot["% contactabilidad"] = camp_pivot["% contactabilidad"].round(2)

        camp_full = touched_by_camp.merge(camp_pivot, on="CAMPAÑA", how="left").fillna(0)
        camp_full = camp_full.sort_values("Contactos_tocados_unicos", ascending=False)

        # Coverage campaña por CONTACTOS (entregados vs tocados)
        base_contacts["CAMPAÑA"] = base_contacts["CAMPAÑA_STD"]
        delivered_by_camp = base_contacts.groupby("CAMPAÑA")["contact_id"].nunique().reset_index(name="Contactos_entregados")
        touched_by_camp2 = jm.groupby("CAMPAÑA")["contact_id"].nunique().reset_index(name="Contactos_tocados")
        cov_camp = delivered_by_camp.merge(touched_by_camp2, on="CAMPAÑA", how="left").fillna({"Contactos_tocados": 0})
        cov_camp["%_tocado"] = cov_camp.apply(lambda r: safe_pct(int(r["Contactos_tocados"]), int(r["Contactos_entregados"])), axis=1)
        cov_camp = cov_camp.sort_values("%_tocado", ascending=False)

        # Tag2 por campaña (% dentro de campaña) sobre matches
        camp_tag2 = (
            jm.groupby(["CAMPAÑA", "tag2"])
            .size()
            .reset_index(name="Conteo")
        )
        totals_by_camp = camp_tag2.groupby("CAMPAÑA")["Conteo"].sum().reset_index(name="Total_camp")
        camp_tag2 = camp_tag2.merge(totals_by_camp, on="CAMPAÑA", how="left")
        camp_tag2["%"] = np.where(camp_tag2["Total_camp"] > 0, (camp_tag2["Conteo"] / camp_tag2["Total_camp"]) * 100, 0.0)
        camp_tag2["%"] = camp_tag2["%"].round(2)
        camp_tag2 = camp_tag2.sort_values(["CAMPAÑA", "Conteo"], ascending=[True, False])

    # ==============================
    # Por HERRAMIENTA (FULL)
    # ==============================
    # Delivered por herramienta (contact-level)
    base_tool_contacts = explode_tools_contacts(base_contacts[["contact_id"] + tool_cols] if tool_cols else base_contacts[["contact_id"]], tool_cols)
    deliv_tool = base_tool_contacts.groupby("HERRAMIENTA")["contact_id"].nunique().reset_index(name="Contactos_entregados")

    if jm.empty:
        herr_full = pd.DataFrame(columns=["HERRAMIENTA","Contactos_tocados_unicos","%_sobre_contactos_tocados","answered","missed","Marcaciones","% contactabilidad"])
        cov_tool = pd.DataFrame(columns=["HERRAMIENTA","Contactos_entregados","Contactos_tocados","%_tocado"])
    else:
        # tocados por herramienta: asignación multi-touch (si contacto tiene varias herramientas cuenta en todas)
        touched_contacts = jm[["contact_id"]].drop_duplicates()
        touched_tool_contacts = touched_contacts.merge(base_tool_contacts, on="contact_id", how="left")
        touched_tool_contacts["HERRAMIENTA"] = touched_tool_contacts["HERRAMIENTA"].fillna("SIN_HERRAMIENTA")
        touched_by_tool = touched_tool_contacts.groupby("HERRAMIENTA")["contact_id"].nunique().reset_index(name="Contactos_tocados_unicos")

        touched_total_contacts_tool = int(touched_contacts["contact_id"].nunique())
        touched_by_tool["%_sobre_contactos_tocados"] = np.where(
            touched_total_contacts_tool > 0,
            (touched_by_tool["Contactos_tocados_unicos"] / touched_total_contacts_tool) * 100,
            0.0
        )
        touched_by_tool["%_sobre_contactos_tocados"] = touched_by_tool["%_sobre_contactos_tocados"].round(2)

        # Contactabilidad por herramienta: replicamos matches por herramienta del contacto (atribución multi-touch)
        jm_tool = jm.merge(base_tool_contacts, on="contact_id", how="left")
        jm_tool["HERRAMIENTA"] = jm_tool["HERRAMIENTA"].fillna("SIN_HERRAMIENTA")

        herr_status = (
            jm_tool.groupby(["HERRAMIENTA", "status_norm"])
            .size()
            .reset_index(name="Conteo")
        )
        herr_pivot = herr_status.pivot_table(index="HERRAMIENTA", columns="status_norm", values="Conteo", aggfunc="sum", fill_value=0).reset_index()
        if "answered" not in herr_pivot.columns: herr_pivot["answered"] = 0
        if "missed" not in herr_pivot.columns: herr_pivot["missed"] = 0
        herr_pivot["Marcaciones"] = herr_pivot["answered"] + herr_pivot["missed"]
        herr_pivot["% contactabilidad"] = np.where(herr_pivot["Marcaciones"] > 0, (herr_pivot["answered"] / herr_pivot["Marcaciones"]) * 100, 0.0)
        herr_pivot["% contactabilidad"] = herr_pivot["% contactabilidad"].round(2)

        herr_full = touched_by_tool.merge(herr_pivot, on="HERRAMIENTA", how="left").fillna(0)
        herr_full = herr_full.sort_values("Contactos_tocados_unicos", ascending=False)

        cov_tool = deliv_tool.merge(
            touched_by_tool.rename(columns={"Contactos_tocados_unicos": "Contactos_tocados"}),
            on="HERRAMIENTA",
            how="left"
        ).fillna({"Contactos_tocados": 0})
        cov_tool["%_tocado"] = cov_tool.apply(lambda r: safe_pct(int(r["Contactos_tocados"]), int(r["Contactos_entregados"])), axis=1)
        cov_tool = cov_tool.sort_values("%_tocado", ascending=False)

    # ==============================
    # Agent breakdown (igual que antes, call-based)
    # ==============================
    if agent_col:
        m_agent = m.copy()
        m_agent["agent"] = m_agent[agent_col].astype(str).fillna("").str.strip()
        m_agent["status_norm"] = m_agent[status_col].astype(str).str.lower()
        m_agent["tag2"] = m_agent[tag2_col].astype(str).fillna("").astype(str) if tag2_col else ""

        sin_tag = (m_agent["tag2"].astype(str).str.strip() == "").groupby(m_agent["agent"]).sum().reset_index(name="Sin_tag2")
        volver = (m_agent["tag2"].astype(str).str.lower().str.strip() == "volver a contactar").groupby(m_agent["agent"]).sum().reset_index(name="Volver_a_contactar")
        nocon = (m_agent["tag2"].astype(str).str.lower().str.strip() == "no conecta").groupby(m_agent["agent"]).sum().reset_index(name="No_conecta")
        total_ag = m_agent.groupby("agent").size().reset_index(name="Marcaciones")
        agent_summary = total_ag.merge(sin_tag, on="agent", how="left").merge(volver, on="agent", how="left").merge(nocon, on="agent", how="left").fillna(0)
        agent_summary = agent_summary.sort_values("Marcaciones", ascending=False)

        ag = (
            m_agent.groupby(["agent", "status_norm"])
            .size()
            .reset_index(name="Conteo")
        )
    else:
        agent_summary = pd.DataFrame(columns=["agent","Marcaciones","Sin_tag2","Volver_a_contactar","No_conecta"])
        ag = pd.DataFrame(columns=["agent","status_norm","Conteo"])

    # ==============================
    # Resumen (FULL)
    # ==============================
    resumen = pd.DataFrame([
        ["Periodo", "Enero 2026"],
        ["Marcaciones totales", total_calls],
        ["Answered", answered_total],
        ["Missed", missed_total],
        ["% contactabilidad (answered/marcaciones)", contactability_total],
        ["Marcaciones con match (al menos 1)", calls_with_match],
        ["% match (marcaciones con match / total)", match_pct],
        ["Match único (marcaciones)", unique_match_calls],
        ["Match múltiple (marcaciones)", multi_match_calls],
        ["Contactos únicos entregados (base)", delivered_contacts_total],
        ["Contactos únicos tocados (si algún teléfono match)", touched_contacts_total],
        ["% cobertura contactos (tocados/entregados)", safe_pct(touched_contacts_total, delivered_contacts_total)],
    ], columns=["Métrica", "Valor"])

    sheets = {
        f"{label}_Resumen": resumen,
        f"{label}_Status": status_counts,
        f"{label}_Status_x_Tag_TOP200": st_tag_top,

        f"{label}_Por_Campaña": camp_full,
        f"{label}_Coverage_Campaña": cov_camp,
        f"{label}_Campaña_x_Tag2": camp_tag2,

        f"{label}_Por_Herramienta": herr_full,
        f"{label}_Coverage_Herramienta": cov_tool,

        f"{label}_Agent_x_Status": ag,
        f"{label}_Agent_Summary": agent_summary,
    }

    return sheets


def main():
    for p in [MARCACIONES_XLSX, BASE_ALL_CLEAN_XLSX]:
        if not os.path.exists(p):
            raise FileNotFoundError(f"No existe: {p}")

    os.makedirs(os.path.dirname(OUT_REPORT_XLSX), exist_ok=True)

    sheets_all = report_full_contacts(BASE_ALL_CLEAN_XLSX, MARCACIONES_XLSX, label="ALL")

    sheets_enero = {}
    if BASE_ENERO_CLEAN_XLSX and os.path.exists(BASE_ENERO_CLEAN_XLSX):
        sheets_enero = report_full_contacts(BASE_ENERO_CLEAN_XLSX, MARCACIONES_XLSX, label="ENERO")
    else:
        print("No se encontró BASE_ENERO_CLEAN_XLSX, se omitirá el corte ENERO.")

    with pd.ExcelWriter(OUT_REPORT_XLSX, engine="openpyxl") as writer:
        for name, df in {**sheets_all, **sheets_enero}.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)

    print("OK - Reporte FULL generado (CONTACTOS)")
    print(f"Output: {OUT_REPORT_XLSX}")


if __name__ == "__main__":
    main()
