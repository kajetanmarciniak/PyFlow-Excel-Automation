# -*- coding: utf-8 -*-
import pandas as pd
from pathlib import Path
from difflib import get_close_matches

# --- KONFIGURACJA ŚCIEŻEK ---
locations_to_check = [
    Path.home() / "OneDrive" / "Desktop" / "PLIKI",
    Path.home() / "Desktop" / "PLIKI",
    Path(__file__).parent / "PLIKI"
]

pliki_folder = next((loc for loc in locations_to_check if loc.exists()), None)

if not pliki_folder:
    print("❌ ERROR: Folder 'PLIKI' not found!")
    exit()

# Pobierz pliki (Excel i CSV)
files = list(pliki_folder.glob("*.xlsx")) + list(pliki_folder.glob("*.xls")) + list(pliki_folder.glob("*.csv"))
files = [f for f in files if not f.name.startswith("Combined_")]

if len(files) < 2:
    print(f"❌ ERROR: Need at least 2 files. Found: {len(files)}")
    exit()

# Słowa kluczowe dla detekcji kolumn
id_patterns = ['id_users', 'id_user', 'id', 'userid', 'user', 'uzytkownik', 'pracownik']
val_patterns = ['wyplata', 'salary', 'wynagrodzenie', 'amount', 'kasa', 'payment']

all_extracted_data = []

print(f"🚀 Starting scan of {len(files)} files...\n")

for file in files:
    try:
        if file.suffix.lower() == '.csv':
            df = pd.read_csv(file, sep=None, engine='python', encoding='utf-8-sig')
        else:
            df = pd.read_excel(file)
        
        print(f"📄 Processing: {file.name}")
        found_in_file = 0

        # SKANER KOLUMN (Szuka par ID + Wypłata w całym arkuszu)
        # Iterujemy po wszystkich kolumnach od lewej do prawej
        for i in range(len(df.columns) - 1): 
            col_name = str(df.columns[i]).lower().strip()
            next_col_name = str(df.columns[i+1]).lower().strip()

            # Czy obecna kolumna to ID?
            is_id = any(p in col_name for p in id_patterns)
            # Czy następna kolumna to Wypłata?
            is_val = any(p in next_col_name for p in val_patterns)

            if is_id and is_val:
                # Wyciągamy parę kolumn
                pair = df.iloc[:, [i, i+1]].copy()
                pair.columns = ['ID_USERS', 'Wyplata']
                
                # Czyścimy dane (usuwamy puste wiersze w tej parze)
                pair = pair.dropna(subset=['ID_USERS'])
                
                if not pair.empty:
                    all_extracted_data.append(pair)
                    found_in_file += len(pair)
                    print(f"   ✅ Found block at columns [{df.columns[i]}, {df.columns[i+1]}] ({len(pair)} rows)")

        if found_in_file == 0:
            print(f"   ⚠️ Warning: No valid data blocks found in {file.name}")

    except Exception as e:
        print(f"   ❌ Error reading {file.name}: {e}")

# --- ŁĄCZENIE I ANALIZA ---
if not all_extracted_data:
    print("\n❌ CRITICAL: No data found in any of the files!")
    exit()

combined_df = pd.concat(all_extracted_data, ignore_index=True)

# Konwersja wypłaty na liczbę (obsługa przecinków i kropek)
combined_df['Wyplata'] = combined_df['Wyplata'].astype(str).str.replace(',', '.').str.strip()
combined_df['Wyplata'] = pd.to_numeric(combined_df['Wyplata'], errors='coerce')

# Usuwamy błędne dane i grupujemy
combined_df = combined_df.dropna(subset=['Wyplata'])
combined_df['ID_USERS'] = combined_df['ID_USERS'].astype(str).str.strip()

print(f"\n📊 Grouping data for {combined_df['ID_USERS'].nunique()} unique users...")
final_report = combined_df.groupby('ID_USERS', as_index=False)['Wyplata'].sum()

# --- ZAPIS WYNIKÓW ---
output_path = pliki_folder / f"Combined_Report_{len(files)}_files.xlsx"
final_report.to_excel(output_path, index=False)

print(f"\n✅ SUCCESS!")
print(f"📁 Saved to: {output_path}")
print(f"💰 Grand Total Salary: {final_report['Wyplata'].sum():,.2f}")