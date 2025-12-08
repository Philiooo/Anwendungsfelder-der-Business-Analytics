import os
import pandas as pd
import numpy as np


# leere Werte definieren
na_values = ["--", ""]

#Tabellen einladen
e = pd.read_excel("C:/Users\peich\Documents\TH-Köln\Vorlesung\Semester 5\ABA\Prüfung\Tabellen\E.ON - Environment.xlsx", sheet_name="Environment", header=1, na_values=na_values)
s = pd.read_excel("C:/Users\peich\Documents\TH-Köln\Vorlesung\Semester 5\ABA\Prüfung\Tabellen\E.ON - Social.xlsx", sheet_name="Social", header=1, na_values=na_values)
g = pd.read_excel("C:/Users\peich\Documents\TH-Köln\Vorlesung\Semester 5\ABA\Prüfung\Tabellen\E.ON - Governance.xlsx", sheet_name="Governance", header=1, na_values=na_values)
k = pd.read_excel("C:/Users\peich\Documents\TH-Köln\Vorlesung\Semester 5\ABA\Prüfung\Tabellen\E.ON - Controversies.xlsx", sheet_name="Controversies", header=1, na_values=na_values)

# Erste Spalte entfernen
for df in [e, s, g, k]:
    df.drop(df.columns[0], axis=1, inplace=True)
# Erste Spalte bennen
for df in [e, s, g, k]:
    df.rename(columns={df.columns[0]: "Attribute"}, inplace=True)
# Leerzeichen in Attributnamen entfernen
for df in [e, s, g, k]:
    df["Attribute"] = df["Attribute"].astype(str).str.strip()

#WICHTIG: Fake-NANs in echte NaN umwandeln
    df["Attribute"] = df["Attribute"].replace(
        ["", "nan", "NaN", "None", "NONE", "Null", "NULL", " "],
        np.nan
    )




#NEUE FUNKTION: Jahr über NaN erkennen
def melt_with_category(df, cat, company):
    df = df.copy()

    # Jahr aus NaN-Zeilen in Spalte 2 übernehmen
    df["Jahr"] = np.where(df["Attribute"].isna(), df.iloc[:, 1], np.nan)

    # Jahr nach unten füllen
    df["Jahr"] = df["Jahr"].ffill()

    # Jahr-Zeilen entfernen
    #df = df[df["Attribute"].notna()]

    # Entpivotieren: alle Spalten außer Attribute + Jahr
    value_cols = [c for c in df.columns if c not in ["Attribute", "Jahr"]]
    df_long = df.melt(
        id_vars=["Attribute", "Jahr"],
        value_vars=value_cols,
        var_name="Variable",
        value_name="Wert"
    )

    # Kategorie + Unternehmen ergänzen
    df_long["Kategorie"] = cat
    df_long["Unternehmen"] = company

    return df_long


#Tabelle entpivotieren
#def melt_with_category(df, cat, company):
#    return (
#        df.melt(id_vars=["Attribute"], var_name="Jahr", value_name="Wert")
#        .assign(Kategorie=cat, Unternehmen=company)
#)

#Kategorien und Unternehmensnamen ergänzen
e_long = melt_with_category(e, "Environmental", "E.ON")
s_long = melt_with_category(s, "Social", "E.ON")
g_long = melt_with_category(g, "Governance", "E.ON")
k_long = melt_with_category(k, "Kontroversen", "E.ON")
df_all = pd.concat([e_long, s_long, g_long, k_long], ignore_index=True)

# "Werte-Spalte" duplizieren
df_all["Wert_raw"] = df_all["Wert"].astype(str).str.strip()
# Numerik-Spalte erzeugen (Prozentzeichen entfernen, Zahlen extrahieren)
df_all["Wert_num"] = (
    df_all["Wert_raw"]
    .str.replace("%", "", regex=False)
    .apply(pd.to_numeric, errors="coerce")
)
# Bool-Spalte erzeugen für "WAHR" und "FALSCH" Werte
# Normalisieren (Großbuchstaben, Leerzeichen weg)
val_norm = df_all["Wert_raw"].astype(str).str.strip().str.upper()
# Neue Spalte: 1 für TRUE, 0 für FALSE, sonst NaN
df_all["Wert_bool"] = val_norm.map({"TRUE": 1, "FALSE": 0})

# Textspalte (alles, was nicht numerisch ist, nicht WAHR/FALSCH ist und nicht NaN ist)
mask_text = (
    df_all["Wert_num"].isna() &  # keine Zahl
    df_all["Wert_bool"].isna()    # keine Bool
)
df_all["Wert_text"] = np.where(mask_text, df_all["Wert"], np.nan)

# Prüfspalte, sind noch Werte nicht übernommen worden aus der Spalte "Werte_Raw"
df_all["Wert_rest"] = np.where(
    df_all[["Wert_bool", "Wert_num", "Wert_text"]].notna().any(axis=1),
    np.nan,
    df_all["Wert_raw"]
)

# Neue Spalte "Wert_fehlend": 1 wenn Wert fehlt; Wenn ein Wert vorhanden ist 0
df_all["Wert_Fehlend"] = np.where(
    df_all["Wert"].isna() | (df_all["Wert"].astype(str).str.strip() == ""),
    1,
    0
)

output_path = "C:/Users\peich\Documents\TH-Köln\Vorlesung\Semester 5\ABA\Prüfung\Tabellen\ESG_Table_Bereinigt_Transformiert.xlsx"
df_all.to_excel(output_path, index=False)
print(f"\n✅ Bereinigte und transformierte Datei gespeichert unter: {output_path}")


if os.path.exists(output_path):
    os.startfile(output_path)
else:
    print("Datei nicht gefunden. Bitte Pfad prüfen.")


# Gesamtanzahl Zeilen
total_rows = len(df_all)
print(f"Total Rows: {total_rows}")
#Fehlende Werte
# Anzahl fehlender Werte (1 = fehlt)
missing_count = df_all["Wert_Fehlend"].sum()
print(f"Missing Count: {missing_count}")
# Anzahl vorhandener Werte (0 = vorhanden)
present_count = (df_all["Wert_Fehlend"] == 0).sum()
print(f"Present Count: {present_count}")
# Anteil fehlender Werte in Prozent
missing_ratio = df_all["Wert_Fehlend"].mean() * 100
print(f"Missing Ratio: {missing_ratio}")
#Anzahl Werte
# Anzahl numerischer Werte
num_count = df_all["Wert_num"].count()
print(f"num_count: {num_count}")
# Anzahl boolescher Werte
bool_count = df_all["Wert_bool"].count()
print(f"bool_count: {bool_count}")
# Anzahl Textwerte
text_count = df_all["Wert_text"].count()
print(f"text_count: {text_count}")

q1 = df_all["Wert_num"].quantile(0.25)
q3 = df_all["Wert_num"].quantile(0.75)
iqr = q3 - q1
df_all["Outlier"] = ~df_all["Wert_num"].between(q1 - 1.5*iqr, q3 + 1.5*iqr)
outliers = df_all[df_all["Outlier"] == True]
print(outliers)



with open("/Users/Desktop/analyse_report.txt", "w", encoding="utf-8") as f:
    f.write("***** Erste Analyse des ESG-Datensatzes *****\n\n")
    f.write(f"Gesamtanzahl Zeilen: {total_rows}\n")
    f.write(f"Fehlende Werte: {missing_count}\n")
    f.write(f"Vorhandene Werte: {present_count}\n")
    f.write(f"Anteil Fehlender Werte: {missing_ratio}\n")
    f.write(f"Anzahl Zahlenwerte: {num_count}\n")
    f.write(f"Anzahl Bool Werte: {bool_count}\n")
    f.write(f"Anzahl Text Werte: {text_count}\n\n")


