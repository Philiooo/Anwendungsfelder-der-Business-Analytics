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
target_attributes = {
    # Energie
    "Total Energy Use / Million in Revenue $": "Energie",
    "Energy Use Total": "Energie",
    "Electricity Purchased": "Energie",
    "Electricity Produced": "Energie",
    "Renewable Energy Use Ratio": "Energie",
    "Total Renewable Energy": "Energie",
    "Renewable Energy Purchased": "Energie",
    "Renewable Energy Produced": "Energie",
    "Renewable Energy Use": "Energie",

    # Wasser
    "Total Water Use / Million in Revenue $": "Wasser",
    "Water Withdrawal Total": "Wasser",
    "Water Recycled": "Wasser",
    "Water Incidents": "Wasser",
    "Water Use Target": "Wasser",
    "Water Withdrawal in Stressed Regions": "Wasser",
    "Water Discharged": "Wasser",

    # Abfall
    "Accidental Spills To Revenues USD in million": "Abfall",
    "Accidental Spills": "Abfall",
    "Waste Recycling Ratio": "Abfall",
    "Hazardous Waste": "Abfall",
    "Waste Total": "Abfall",
    "Total Waste / Million in Revenue $": "Abfall",
    "Waste Reduction Initiatives": "Abfall",
    "Waste Recycled To Total Waste": "Abfall",

    # Emissionen
    "Internal Carbon Pricing": "Emissionen",
    "Emission Reduction Target Percentage": "Emissionen",
    "Emission Reduction Target Year": "Emissionen",
    "Estimated CO2 Equivalents Emission Total": "Emissionen",
    "Total CO2 Emissions / Million in Revenue $": "Emissionen",
    "CO2 Equivalent Emissions Total": "Emissionen",
    "CO2 Equivalent Emissions Direct, Scope 1": "Emissionen",
    "CO2 Equivalent Emissions Indirect, Scope 2": "Emissionen",
    "CO2 Equivalent Emissions Indirect, Scope 3": "Emissionen",

    # Target
    "Targets Resource Use": "Target",
    "Resource Reduction Targets": "Target",
    "Targets Water Efficiency": "Target",
    "Targets Energy Efficiency": "Target",
    "Water Use Target": "Target",
    "Water Stressed Targets": "Target",
    "Targets Emissions": "Target",
    "Emission Reduction Target Percentage": "Target",
    "Emission Reduction Target Year": "Target",
    "Emissions Target Type": "Target",
    "Emissions Target Annual Reduction": "Target",
    "Reduction Target, GHG Emissions Scope 1,2": "Target",
    "Reduction Target, GHG Emissions Scope 1,2,3": "Target",
    "Reduction Target, GHG Emissions Intensity Scope 1,2": "Target",
    "Reduction Target, GHG Emissions Intensity Scope 1,2,3": "Target",
    "Targets Waste": "Target",
    "Targets Pollution": "Target",
    "Biodiversity Targets": "Target",

    # Policy
    "Policy Resource Efficiency": "Policy",
    "Resource Reduction Policy": "Policy",
    "Policy Water Efficiency": "Policy",
    "Policy Energy Efficiency": "Policy",
    "Policy Sustainable Packaging": "Policy",
    "Policy Environmental Supply Chain": "Policy",
    "Policy Environmental Supply Chain Management": "Policy",
    "Policy Pollution": "Policy",
    "Policy Waste": "Policy",
    "Policy Emissions": "Policy",
    "Policy Nuclear Safety": "Policy",
}

def map_attribute(attr):
    val = target_attributes.get(attr, None)
    if val in ["Energie", "Wasser", "Abfall", "Emissionen", "Target", "Policy"]:
        return val, "Operativ"
    else:
        return None, "Operativ"

df_all[["Kategorie", "Datentyp"]] = df_all["Attribute"].apply(lambda x: pd.Series(map_attribute(x)))

df_final_filtered = df_all[df_all["Attribute"].isin(target_attributes)]

# Optional: gefilterte Datei separat speichern
filtered_output_path = r"C:/Users/peich/Documents/TH-Köln/Vorlesung/Semester 5/ABA/Prüfung/Tabellen/ESG_Filtered_Analysis.xlsx"
df_final_filtered.to_excel(filtered_output_path, index=False)
print(f"✅ Filtered Data saved to: {filtered_output_path}")


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
