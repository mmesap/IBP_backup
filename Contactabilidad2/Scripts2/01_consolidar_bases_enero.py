import os
import pandas as pd

# ---------------------------
# Config
# ---------------------------
INPUT_DIR = "./Contactabilidad/Data/campanas enero"  # folder with January campaign files
SHEET_NAME = "IMR"
OUTPUT_PATH = "./Contactabilidad/out/bases_enero_consolidado.xlsx"

# Columns you care about (keep all if you want; for now we keep everything)
# You can uncomment to keep only a subset later.
# KEEP_COLS = [...]

def main():
    if not os.path.isdir(INPUT_DIR):
        raise FileNotFoundError(f"Input folder not found: {INPUT_DIR}")

    files = [
        f for f in os.listdir(INPUT_DIR)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    ]

    if not files:
        raise FileNotFoundError(f"No .xlsx files found in: {INPUT_DIR}")

    dfs = []
    skipped = []
    for f in sorted(files):
        path = os.path.join(INPUT_DIR, f)
        try:
            # Read IMR sheet explicitly
            df = pd.read_excel(path, sheet_name=SHEET_NAME, dtype=str)

            # Add provenance columns (useful for debugging/audit)
            df["__source_file__"] = f
            df["__source_sheet__"] = SHEET_NAME

            # If you want to keep only certain cols later:
            # if KEEP_COLS:
            #     missing = [c for c in KEEP_COLS if c not in df.columns]
            #     for c in missing:
            #         df[c] = ""
            #     df = df[KEEP_COLS + ["__source_file__", "__source_sheet__"]]

            dfs.append(df)

        except Exception as e:
            skipped.append((f, str(e)))

    if not dfs:
        raise RuntimeError(
            "No files could be read successfully from IMR sheets. "
            "Check if the sheet name is exactly 'IMR' in the files."
        )

    out = pd.concat(dfs, ignore_index=True)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    out.to_excel(OUTPUT_PATH, index=False)

    print("OK - January bases consolidated")
    print(f"Input folder: {INPUT_DIR}")
    print(f"Files found: {len(files)} | Files read: {len(dfs)} | Files skipped: {len(skipped)}")
    print(f"Rows consolidated: {len(out):,}")
    print(f"Output: {OUTPUT_PATH}")

    if skipped:
        print("\nSkipped files (file -> reason):")
        for f, reason in skipped[:20]:
            print(f"- {f} -> {reason}")
        if len(skipped) > 20:
            print(f"... and {len(skipped) - 20} more")

if __name__ == "__main__":
    main()
