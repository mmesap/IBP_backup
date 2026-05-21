import os
import pandas as pd

CRUCE_PATH = "./Contactabilidad/out/cruce_marcaciones_vs_entregas.xlsx"
OUT_PATH = "./Contactabilidad/out/Informe_A_Marcaciones_Enero_Campaña_Tags.xlsx"

STATUS_COL = "status"
TAG2_COL = "tag 2"
ANSWERED_VALUE = "answered"

def safe_col(df, col):
    if col not in df.columns:
        df[col] = None

def parse_campaigns(row):
    """
    - UNICO: usa CAMPAÑA con peso 1
    - MULTIPLE: usa CAMPAÑAS_MATCH separadas por | con peso 1/N
    """
    if str(row.get("match", "")).strip() != "1":
        return [], 0.0

    mtype = row.get("match_type", None)

    if mtype == "UNICO":
        camp = row.get("CAMPAÑA", None)
        if pd.isna(camp) or str(camp).strip() == "":
            return [], 0.0
        return [str(camp).strip()], 1.0

    if mtype == "MULTIPLE":
        camps = row.get("CAMPAÑAS_MATCH", None)
        if pd.isna(camps) or str(camps).strip() == "":
            return [], 0.0
        parts = [c.strip() for c in str(camps).split("|") if c.strip() != ""]
        if not parts:
            return [], 0.0
        return parts, 1.0 / len(parts)

    return [], 0.0

def main():
    if not os.path.exists(CRUCE_PATH):
        raise FileNotFoundError(f"No existe: {CRUCE_PATH}")

    df = pd.read_excel(CRUCE_PATH, dtype=str)

    for c in ["match", "match_type", "CAMPAÑA", "CAMPAÑAS_MATCH"]:
        safe_col(df, c)
    safe_col(df, STATUS_COL)
    safe_col(df, TAG2_COL)

    rows = []
    for _, r in df.iterrows():
        camps, w = parse_campaigns(r)
        if not camps:
            continue

        status = str(r.get(STATUS_COL, "")).strip().lower()
        tag2 = r.get(TAG2_COL, None)
        tag2 = "SIN_TAG_2" if (pd.isna(tag2) or str(tag2).strip() == "") else str(tag2).strip()

        for c in camps:
            rows.append({
                "CAMPAÑA": c,
                "w": w,
                "answered_w": w if status == ANSWERED_VALUE else 0.0,
                "tag 2": tag2
            })

    camp_long = pd.DataFrame(rows)

    if camp_long.empty:
        camp_contact = pd.DataFrame(columns=["CAMPAÑA", "marcaciones_match_ponderadas", "answered_ponderado", "%_answered"])
        camp_tag2 = pd.DataFrame(columns=["CAMPAÑA", "tag 2", "conteo_ponderado", "%_dentro_campaña"])
    else:
        camp_contact = (
            camp_long.groupby("CAMPAÑA", dropna=False)
            .agg(
                marcaciones_match_ponderadas=("w", "sum"),
                answered_ponderado=("answered_w", "sum"),
            )
            .reset_index()
        )
        camp_contact["%_answered"] = (
            (camp_contact["answered_ponderado"] / camp_contact["marcaciones_match_ponderadas"].replace({0: pd.NA})) * 100
        ).round(2)
        camp_contact = camp_contact.sort_values("%_answered", ascending=False, na_position="last")

        camp_tag2 = (
            camp_long.groupby(["CAMPAÑA", "tag 2"], dropna=False)["w"]
            .sum()
            .reset_index(name="conteo_ponderado")
        )
        totals = camp_tag2.groupby("CAMPAÑA")["conteo_ponderado"].sum().reset_index(name="total_camp")
        camp_tag2 = camp_tag2.merge(totals, on="CAMPAÑA", how="left")
        camp_tag2["%_dentro_campaña"] = (
            (camp_tag2["conteo_ponderado"] / camp_tag2["total_camp"].replace({0: pd.NA})) * 100
        ).round(2)
        camp_tag2 = camp_tag2.drop(columns=["total_camp"]).sort_values(["CAMPAÑA", "conteo_ponderado"], ascending=[True, False])

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with pd.ExcelWriter(OUT_PATH, engine="openpyxl") as writer:
        camp_contact.to_excel(writer, sheet_name="Contactabilidad_x_Campaña", index=False)
        camp_tag2.to_excel(writer, sheet_name="Campaña_x_Tag2", index=False)

    print("OK. Informe A generado.")
    print(f"Salida: {OUT_PATH}")

if __name__ == "__main__":
    main()
