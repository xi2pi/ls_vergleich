
import pandas as pd

# Lade die Excel-Dateien
xlsx1_path = 'ls_all.xlsx'
xlsx2_path = 'ls_all_2025.xlsx'
df1 = pd.read_excel(xlsx1_path)
df2 = pd.read_excel(xlsx2_path)

# Sicherstellen, dass die Spalte "aktenzeichen" vorhanden ist
if 'aktenzeichen' not in df1.columns or 'aktenzeichen' not in df2.columns:
    raise ValueError("Die Spalte 'aktenzeichen' muss in beiden Dateien vorhanden sein.")

# Vergleiche die Spalte "aktenzeichen"
set1 = set(df1['aktenzeichen'])
set2 = set(df2['aktenzeichen'])

# Finde Unterschiede
nur_in_df1 = set1 - set2
nur_in_df2 = set2 - set1

# Extrahiere die entsprechenden Zeilen
unterschiede_df1 = df1[df1['aktenzeichen'].isin(nur_in_df1)]
unterschiede_df2 = df2[df2['aktenzeichen'].isin(nur_in_df2)]

# Kombiniere die Unterschiede
unterschiede = pd.concat([unterschiede_df1, unterschiede_df2], ignore_index=True)

# Speichere die Unterschiede in eine neue Excel-Datei
unterschiede.to_excel('unterschiede.xlsx', index=False)

print("Vergleich abgeschlossen. Unterschiede wurden in 'unterschiede.xlsx' gespeichert.")
