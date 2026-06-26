"""
Excel Data Cleaner & Analyzer
---------------------------------
Messy Excel/CSV ko clean karta hai:
- Duplicate rows hatata hai
- Column names aur text formatting fix karta hai
- Dates aur numbers ko proper format deta hai
- Missing/invalid data flag karta hai
- Summary charts generate karta hai
- Clean file + report Excel mein save karta hai

Usage:
    python excel_cleaner.py input.xlsx
    python excel_cleaner.py data.csv
"""

import sys
import os
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # GUI ke bina charts save karne ke liye
import matplotlib.pyplot as plt


def load_file(path):
    """Excel ya CSV file load karta hai."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(path)
    elif ext == ".csv":
        return pd.read_csv(path)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Use .xlsx, .xls ya .csv")


def clean_column_names(df):
    """Column names ko clean karta hai: lowercase, no spaces, no extra symbols."""
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"[^\w]+", "_", regex=True)
        .str.strip("_")
    )
    return df


def clean_text_columns(df):
    """Text columns ko trim aur title-case karta hai."""
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()
        # 'nan' string ko wapas missing banate hain
        df[col] = df[col].replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})
    return df


def fix_data_types(df):
    """Numbers aur dates ko detect karke proper type deta hai."""
    for col in df.columns:
        # Pehle number try karo
        converted = pd.to_numeric(df[col], errors="coerce")
        if converted.notna().sum() >= len(df) * 0.6:  # 60%+ valid numbers
            df[col] = converted
            continue
        # Phir date try karo
        try:
            dates = pd.to_datetime(df[col], errors="coerce")
            if dates.notna().sum() >= len(df) * 0.6:
                df[col] = dates
        except Exception:
            pass
    return df


def remove_duplicates(df):
    """Duplicate rows hatata hai aur count return karta hai."""
    before = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    removed = before - len(df)
    return df, removed


def build_quality_report(df, removed_dupes):
    """Har column ke liye missing values aur data quality summary."""
    report = []
    for col in df.columns:
        missing = int(df[col].isna().sum())
        missing_pct = round(missing / len(df) * 100, 1) if len(df) else 0
        report.append({
            "column": col,
            "data_type": str(df[col].dtype),
            "missing_values": missing,
            "missing_percent": missing_pct,
            "unique_values": int(df[col].nunique(dropna=True)),
        })
    report_df = pd.DataFrame(report)
    report_df.attrs["duplicates_removed"] = removed_dupes
    return report_df


def generate_charts(df, output_prefix):
    """Numeric columns ke liye charts banata hai."""
    chart_files = []
    numeric_cols = df.select_dtypes(include="number").columns[:3]  # pehle 3 numeric
    for col in numeric_cols:
        plt.figure(figsize=(7, 4))
        df[col].dropna().plot(kind="hist", bins=20, color="#4F46E5", edgecolor="white")
        plt.title(f"Distribution of {col}", fontsize=13, fontweight="bold")
        plt.xlabel(col)
        plt.ylabel("Count")
        plt.tight_layout()
        fname = f"{output_prefix}_chart_{col}.png"
        plt.savefig(fname, dpi=120)
        plt.close()
        chart_files.append(fname)
    return chart_files


def save_results(clean_df, report_df, output_path):
    """Clean data aur quality report ko ek Excel file mein save karta hai."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        clean_df.to_excel(writer, sheet_name="Clean Data", index=False)
        report_df.to_excel(writer, sheet_name="Quality Report", index=False)


def main():
    if len(sys.argv) < 2:
        print("Usage: python excel_cleaner.py <file.xlsx | file.csv>")
        sys.exit(1)

    input_path = sys.argv[1]
    if not os.path.exists(input_path):
        print(f"File nahi mili: {input_path}")
        sys.exit(1)

    print(f"📂 Loading: {input_path}")
    df = load_file(input_path)
    print(f"   Original rows: {len(df)}, columns: {len(df.columns)}")

    df = clean_column_names(df)
    df = clean_text_columns(df)
    df = fix_data_types(df)
    df, removed = remove_duplicates(df)
    print(f"🧹 Duplicates removed: {removed}")

    report_df = build_quality_report(df, removed)

    base = os.path.splitext(input_path)[0]
    output_path = f"{base}_cleaned.xlsx"
    save_results(df, report_df, output_path)
    print(f"✅ Clean file saved: {output_path}")

    charts = generate_charts(df, base)
    if charts:
        print(f"📊 Charts generated: {', '.join(charts)}")

    print("\n✨ Done! Spreadsheet clean aur analysis-ready hai.")


if __name__ == "__main__":
    main()
