# -*- coding: utf-8 -*-
import pandas as pd
from pathlib import Path
import os

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

# --- OBSŁUGA PLIKÓW ---
files = list(pliki_folder.glob("*.xlsx")) + list(pliki_folder.glob("*.xls")) + list(pliki_folder.glob("*.csv"))
files = [f for f in files if not f.name.startswith("Combined_")]

if len(files) < 2:
    print(f"❌ ERROR: Need at least 2 files. Found: {len(files)}")
    exit()

# Słowa kluczowe (wzorce)
id_patterns = ['id_users', 'id_user', 'id', 'userid', 'user', 'uzytkownik', 'pracownik', 'kajtek', 'name', 'osoba']
val_patterns = ['wyplata', 'salary', 'wynagrodzenie', 'amount', 'kasa', 'payment', 'pieniadze', 'total', 'kwota']

all_extracted_data = []

print(f"🚀 Starting scan of {len(files)} files...\n")

for file in files:
    try:
        if str(file).endswith('.csv'):
            df = pd.read_csv(file, sep=None, engine='python', encoding='utf-8-sig')
        else:
            df = pd.read_excel(file)
        
        print(f"📄 Processing: {file.name}")

        # --- INTELIGENTNE WYKRYWANIE KOLUMN ---
        id_col = None
        val_col = None

        # 1. Próbujemy dopasować po nazwach (twój wzorzec)
        id_col = next((c for c in df.columns if any(p in str(c).lower() for p in id_patterns)), None)
        val_col = next((c for c in df.columns if any(p in str(c).lower() for p in val_patterns)), None)

        # 2. Jeśli nie znaleźliśmy po nazwach, bierzemy 1-szą i 2-gą ZAPISANĄ kolumnę (omijamy puste/Unnamed)
        if not id_col or not val_col:
            valid_columns = [c for c in df.columns if "Unnamed" not in str(c)]
            if len(valid_columns) >= 2:
                id_col = valid_columns[0]  # Pierwsza zapisana
                val_col = valid_columns[1] # Druga zapisana
                print(f"   ℹ️ Auto-detected by position: [{id_col}] and [{val_col}]")

        # 3. Pobieranie danych jeśli znaleźliśmy parę
        if id_col and val_col:
            pair = df[[id_col, val_col]].copy()
            pair.columns = ['ID_USERS', 'Wyplata']
            
            # Czyszczenie danych
            pair = pair.dropna(subset=['ID_USERS'])
            
            if not pair.empty:
                all_extracted_data.append(pair)
                print(f"   ✅ Successfully extracted {len(pair)} rows.")
        else:
            print(f"   ⚠️ Warning: Could not find valid columns in {file.name}")

    except Exception as e:
        print(f"   ❌ Error reading {file.name}: {e}")

# --- ŁĄCZENIE I ANALIZA ---
if not all_extracted_data:
    print("\n❌ CRITICAL: No data found in any of the files!")
    exit()

combined_df = pd.concat(all_extracted_data, ignore_index=True)

# Konwersja wypłaty (usuwanie symboli walut, spacji, zamiana przecinków na kropki)
combined_df['Wyplata'] = combined_df['Wyplata'].astype(str).str.replace(',', '.').str.replace(r'[^\d.]', '', regex=True)
combined_df['Wyplata'] = pd.to_numeric(combined_df['Wyplata'], errors='coerce')

combined_df = combined_df.dropna(subset=['Wyplata'])
combined_df['ID_USERS'] = combined_df['ID_USERS'].astype(str).str.strip().str.upper()

print(f"\n📊 Grouping data for {combined_df['ID_USERS'].nunique()} unique users...")
final_report = combined_df.groupby('ID_USERS', as_index=False)['Wyplata'].sum()

# --- ZAPIS WYNIKÓW DO .XLSX ---
output_path = pliki_folder / f"Combined_Report_{len(files)}_files.xlsx"

try:
    final_report.to_excel(output_path, index=False, engine='openpyxl')
    print(f"\n✅ SUCCESS!")
    print(f"📁 Saved to: {output_path}")
    print(f"💰 Grand Total Salary: {final_report['Wyplata'].sum():,.2f}")
except Exception as e:
    print(f"❌ Error saving Excel file: {e}")
