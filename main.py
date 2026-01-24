import os
import pandas as pd
import numpy as np


# KONFIGURATION
na_values = ["--", ""]
company_name = "E.ON"
output_path = r"C:/Users/peich/Documents/TH-Köln/Vorlesung/Semester 5/ABA/Prüfung/Tabellen/ESG_Table_Bereinigt_Transformiert.xlsx"
report_file = r"C:/Users/peich/Documents/TH-Köln/Vorlesung/Semester 5/ABA/Prüfung/Tabellen/analyse_report.txt"


# EXCEL DATEIEN LADEN
e = pd.read_excel(r"C:/Users/peich/Documents/TH-Köln/Vorlesung/Semester 5/ABA/Prüfung/Tabellen/E.ON - Environment.xlsx",
                  sheet_name="Environment", header=1, na_values=na_values)
s = pd.read_excel(r"C:/Users/peich/Documents/TH-Köln/Vorlesung/Semester 5/ABA/Prüfung/Tabellen/E.ON - Social.xlsx",
                  sheet_name="Social", header=1, na_values=na_values)
g = pd.read_excel(r"C:/Users/peich/Documents/TH-Köln/Vorlesung/Semester 5/ABA/Prüfung/Tabellen/E.ON - Governance.xlsx",
                  sheet_name="Governance", header=1, na_values=na_values)
k = pd.read_excel(r"C:/Users/peich/Documents/TH-Köln/Vorlesung/Semester 5/ABA/Prüfung/Tabellen/E.ON - Controversies.xlsx",
                  sheet_name="Controversies", header=1, na_values=na_values)

dfs_raw = [e, s, g, k]
categories = ["Environmental", "Social", "Governance", "Kontroversen"]


# SPALTEN BEREINIGEN & JAHRE UMBENENNEN
for df in dfs_raw:
    # erste Spalte entfernen & Attribute benennen
    df.drop(df.columns[0], axis=1, inplace=True)
    df.rename(columns={df.columns[0]: "Attribute"}, inplace=True)
    df["Attribute"] = df["Attribute"].astype(str).str.strip()
    df["Attribute"] = df["Attribute"].replace(["", "nan", "NaN", "None", "NONE", "Null", "NULL", " "], np.nan)

    # Jahres-Spalten: 2023 bis 2002
    year_columns = {f"Unnamed: {i}": year for i, year in zip(range(2, 24), range(2023, 2001, -1))}
    df.rename(columns=year_columns, inplace=True)

    # Sicherstellen, dass Spaltennamen Integer sind
    df.columns = ["Attribute"] + [int(col) if col != "Attribute" else col for col in df.columns[1:]]


# ENT-PIVOTIEREN & KATEGORIE HINZUFÜGEN
def melt_with_category(df, category, company):
    df_long = df.melt(id_vars=["Attribute"], var_name="Jahr", value_name="Wert")
    df_long["Jahr"] = df_long["Jahr"].astype(int)  # Jahr als Integer
    df_long["Kategorie"] = category
    df_long["Unternehmen"] = company
    return df_long

dfs_long = [melt_with_category(df, cat, company_name) for df, cat in zip(dfs_raw, categories)]
df_all = pd.concat(dfs_long, ignore_index=True)


# WERTE TRANSFORMIEREN
df_all["Wert_raw"] = df_all["Wert"].astype(str).str.strip()
df_all["Wert_num"] = df_all["Wert_raw"].str.replace("%", "", regex=False).apply(pd.to_numeric, errors="coerce")
val_norm = df_all["Wert_raw"].str.upper()
df_all["Wert_bool"] = val_norm.map({"TRUE": 1, "FALSE": 0})
mask_text = df_all["Wert_num"].isna() & df_all["Wert_bool"].isna()
df_all["Wert_text"] = np.where(mask_text, df_all["Wert"], np.nan)
df_all["Wert_rest"] = np.where(df_all[["Wert_num", "Wert_bool", "Wert_text"]].notna().any(axis=1),
                               np.nan, df_all["Wert_raw"])
df_all["Wert_Fehlend"] = np.where(df_all["Wert"].isna() | (df_all["Wert"].astype(str).str.strip() == ""), 1, 0)


# AUSREIßER ERKENNEN
q1 = df_all["Wert_num"].quantile(0.25)
q3 = df_all["Wert_num"].quantile(0.75)
iqr = q3 - q1
df_all["Outlier"] = ~df_all["Wert_num"].between(q1 - 1.5*iqr, q3 + 1.5*iqr)
outliers = df_all[df_all["Outlier"] == True]


# Filtern der bestimmten Attribute
scope3_attributes = {
    # Gesamt
    "CO2 Equivalent Emissions Indirect, Scope 3",

    # Upstream
    "Upstream scope 3 emissions Purchased goods and services",
    "Upstream scope 3 emissions Capital goods",
    "Upstream scope 3 emissions Fuel- and Energy-related Activities",
    "Upstream scope 3 emissions Transportation and Distribution",
    "Upstream scope 3 emissions Waste Generated in Operations",
    "Upstream scope 3 emissions Business Travel",
    "Upstream scope 3 emissions Employee Commuting",
    "Upstream scope 3 emissions Leased Assets",
    "Upstream scope 3 emissions Other",


    # Downstream
    "Downstream scope 3 emissions Transportation and Distribution",
    "Downstream scope 3 emissions Processing of Sold Products",
    "Downstream scope 3 emissions Use of Sold Products",
    "Downstream scope 3 emissions End-of-life Treatment of Sold Products",
    "Downstream scope 3 emissions Leased Assets",
    "Downstream scope 3 emissions Franchises",
    "Downstream scope 3 emissions Investments",
    "Downstream scope 3 emissions Other",


    # Intensität & Trends
    "GHG Emissions Scope 1,2,3 YoY Chg Tonnes",
    "Target Variance, GHG Emissions Intensity Scope 1,2,3 YoY Chg Pct",
    "GHG Emissions Scope 1,2,3 Paris Agreement Aligned",
    "GHG Emissions Scope 1,2,3 - 3 year CAGR",
    "GHG Emissions Scope 1,2,3 - 5 year CAGR",
    "GHG Emissions Intensity Scope 1,2,3 - 3 year CAGR",
    "GHG Emissions Intensity Scope 1,2,3 - 5 year CAGR",
    "GHG Emissions Intensity Indirect, Scope 3 YoY Percent Chg",
    "CO2 Equivalent Emissions Indirect Scope 3 YoY",
    "CO2 Equivalent Emissions Indirect Scope 3 CAGR 3 year",
    "CO2 Equivalent Emissions Indirect Scope 3 CAGR 5 year",
    "Emission Reduction Target Percentage",
    "Reduction Target, GHG Emissions Scope 1,2,3 Tonnes",

    #Scope 1 & Scope 2
    "CO2 Equivalent Emissions Direct, Scope 1",
    "CO2 Equivalent Emissions Indirect, Scope 2",


}

df_scope3 = df_all[df_all["Attribute"].isin(scope3_attributes)].copy()

df_scope3["ESG_Fokus"] = "Scope 3"


# Optional: gefilterte Datei separat speichern
scope3_output_path = r"C:/Users/peich/Documents/TH-Köln/Vorlesung/Semester 5/ABA/Prüfung/Tabellen/ESG_Scope3_Dashboard.xlsx"
df_scope3.to_excel(scope3_output_path, index=False)

print(f"✅ Scope-3 Dashboard Table saved to: {scope3_output_path}")


# ERGEBNISSE SPEICHERN
df_all.to_excel(output_path, index=False)

# Analyse-Report erstellen
total_rows = len(df_all)
missing_count = df_all["Wert_Fehlend"].sum()
present_count = (df_all["Wert_Fehlend"] == 0).sum()
missing_ratio = df_all["Wert_Fehlend"].mean() * 100
num_count = df_all["Wert_num"].count()
bool_count = df_all["Wert_bool"].count()
text_count = df_all["Wert_text"].count()

with open(report_file, "w", encoding="utf-8") as f:
    f.write("***** Erste Analyse des ESG-Datensatzes *****\n\n")
    f.write(f"Gesamtanzahl Zeilen: {total_rows}\n")
    f.write(f"Fehlende Werte: {missing_count}\n")
    f.write(f"Vorhandene Werte: {present_count}\n")
    f.write(f"Anteil Fehlender Werte: {missing_ratio:.2f}%\n")
    f.write(f"Anzahl Zahlenwerte: {num_count}\n")
    f.write(f"Anzahl Bool Werte: {bool_count}\n")
    f.write(f"Anzahl Text Werte: {text_count}\n\n")

print(f"\n✅ Bereinigte und transformierte Datei: {output_path}")
print(f"✅ Analyse-Report: {report_file}")

if os.path.exists(output_path):
    os.startfile(output_path)
