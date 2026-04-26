# -*- coding: utf-8 -*-
"""
Created on Sat Apr 25 13:04:38 2026

@author: bocci
"""


import pandas as pd
import json, time
from pathlib import Path
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError

SLEEP_SEC = 1.1 # Nominatim: max 1 req/sec
CHECKPOINT = 50 # salva ogni N righe
USER_AGENT = 'gis_portfolio_geocoder_v7/1.0'

# Lista completa di tutti i capoluoghi di provincia italiani
CAPOLUOGHI = {
    'agrigento', 'alessandria', 'ancona', 'aosta', 'arezzo', 'ascoli piceno', 'asti', 
    'avellino', 'bari', 'barletta', 'andria', 'trani', 'belluno', 'benevento', 'bergamo', 
    'biella', 'bologna', 'bolzano', 'brescia', 'brindisi', 'cagliari', 'caltanissetta', 
    'campobasso', 'carbonia', 'caserta', 'catania', 'catanzaro', 'chieti', 'como', 
    'cosenza', 'cremona', 'crotone', 'cuneo', 'enna', 'fermo', 'ferrara', 'firenze', 
    'foggia', 'forlì', 'cesena', 'frosinone', 'genova', 'gorizia', 'grosseto', 'imperia', 
    'isernia', "l'aquila", 'la spezia', 'latina', 'lecce', 'lecco', 'livorno', 'lodi', 
    'lucca', 'macerata', 'mantova', 'massa', 'carrara', 'matera', 'messina', 'milano', 
    'modena', 'monza', 'napoli', 'novara', 'nuoro', 'oristano', 'padova', 'palermo', 
    'parma', 'pavia', 'perugia', 'pesaro', 'urbino', 'pescara', 'piacenza', 'pisa', 
    'pistoia', 'pordenone', 'potenza', 'prato', 'ragusa', 'ravenna', 'reggio calabria', 
    'reggio emilia', 'rieti', 'rimini', 'roma', 'rovigo', 'salerno', 'sassari', 'savona', 
    'siena', 'siracusa', 'sondrio', 'taranto', 'teramo', 'terni', 'torino', 'trapani', 
    'trento', 'treviso', 'trieste', 'udine', 'varese', 'venezia', 'verbania', 'vercelli', 
    'verona', 'vibo valentia', 'vicenza', 'viterbo'
}

def geocode_address(geocoder, address: str, retries=3):
    for attempt in range(retries):
        try:
            loc = geocoder.geocode(address, timeout=10, addressdetails=True)
            if loc:
                addr_info = loc.raw.get('address', {})
                country_code = addr_info.get('country_code', '').lower()
                
                citta_raw = addr_info.get('city', addr_info.get('town', addr_info.get('village', addr_info.get('county', ''))))
                nome_citta = citta_raw.title() if citta_raw else "Non determinata"
                postcode = addr_info.get('postcode', '')
                city_name_lower = citta_raw.lower() if citta_raw else ""
                is_city_tag = 'city' in addr_info
                
                # Classificazione
                if country_code != 'it':
                    categoria = 'Estero'
                elif is_city_tag or city_name_lower in CAPOLUOGHI:
                    categoria = 'Capoluogo/Grande Città'
                else:
                    categoria = 'Provincia/Piccolo Comune'
                    
                return (loc.latitude, loc.longitude, categoria, nome_citta, postcode)
            
            return (None, None, None, None, None)
        except GeocoderTimedOut:
            time.sleep(2 ** attempt)
        except GeocoderServiceError as e:
            print(f' Errore servizio: {e}')
            return (None, None, None, None, None)
    return (None, None, None, None, None)

def colora_matrix(df_export, df_completo):
    """Crea una matrice di colori della stessa forma del dataframe esportato"""
    styles = pd.DataFrame('', index=df_export.index, columns=df_export.columns)
    
    for idx in df_export.index:
        # Recuperiamo la riga originale con i dati nascosti
        riga = df_completo.loc[idx]
        cat = riga.get('Categoria_Geografica')
        bg = ''
        
        # Colore base riga
        if cat == 'Estero':
            bg = 'background-color: #FF9999' # Rosso chiaro
        elif cat == 'Capoluogo/Grande Città':
            bg = 'background-color: #90EE90' # Verde chiaro
        elif cat == 'Provincia/Piccolo Comune':
            bg = 'background-color: #FFFFE0' # Giallo chiaro
            
        styles.loc[idx, :] = bg
        
        # Colore specifico celle riparate
        imputed_str = str(riga.get('_imputed', ''))
        if imputed_str and imputed_str != 'nan':
            for col in imputed_str.split(','):
                if col and col in styles.columns:
                    styles.loc[idx, col] = 'background-color: #87CEFA' # Azzurro
                    
    return styles

def run_batch(input_path: str, col_id: str, col_addr: str, col_city: str, col_cap: str):
    input_path = input_path.strip('"').strip("'")
    
    df = pd.read_csv(input_path, dtype=str) if input_path.endswith('.csv') else pd.read_excel(input_path, dtype=str)
        
    geocoder = Nominatim(user_agent=USER_AGENT)
    results, failed = [], []
    out_stem = Path(input_path).stem
    out_dir = Path(input_path).parent

    for i, row in df.iterrows():
        addr_parts = [str(row[col_addr])]
        if col_cap and col_cap in df.columns and pd.notna(row[col_cap]):
            addr_parts.append(str(row[col_cap]))
        if col_city and col_city in df.columns and pd.notna(row[col_city]):
            addr_parts.append(str(row[col_city]))

        full_address = ", ".join(addr_parts)
        lat, lon, categoria, nome_citta, postcode = geocode_address(geocoder, full_address)
        
        row_dict = row.to_dict()
        output_row = {}
        imputed_cols = [] # Tracker per le celle azzurre
        
        if col_id and col_id in df.columns:
            output_row[col_id] = row_dict.pop(col_id)
            
        output_row.update(row_dict)
        
        if lat is not None:
            # --- LOGICA DI RIEMPIMENTO DATI MANCANTI ---
            if col_cap and col_cap in df.columns:
                val_cap = str(output_row.get(col_cap, '')).strip().lower()
                if val_cap in ('', 'nan', 'none') and postcode:
                    output_row[col_cap] = postcode
                    imputed_cols.append(col_cap)
                    print(f'   [FIX] Aggiunto CAP ({postcode}) in {full_address}')
            
            if col_city and col_city in df.columns:
                val_city = str(output_row.get(col_city, '')).strip().lower()
                if val_city in ('', 'nan', 'none') and nome_citta != "Non determinata":
                    output_row[col_city] = nome_citta
                    imputed_cols.append(col_city)
                    print(f'   [FIX] Aggiunta Città ({nome_citta}) in {full_address}')
            
            # --------------------------------------------
            
            output_row.update({
                'Indirizzo_Cercato': full_address, 
                'Città_Trovata': nome_citta,
                'lat': lat, 'lon': lon,
                'Categoria_Geografica': categoria,
                '_imputed': ",".join(imputed_cols) # Colonna nascosta per i colori
            })
            results.append(output_row)
        else:
            output_row.update({'Indirizzo_Cercato': full_address, 'errore': 'non trovato'})
            failed.append(output_row)
            print(f' [FAIL] {full_address}')
            
        time.sleep(SLEEP_SEC)

    # --- SALVATAGGIO IN UN UNICO FILE EXCEL CON DUE FOGLI ---
    out_excel = out_dir / f'{out_stem}_risultati_finali.xlsx'
    
    with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
        df_completo = pd.DataFrame(results)
        
        if not df_completo.empty:
            # Rimuoviamo la colonna nascosta per non sporcare il file Excel finale
            df_export = df_completo.drop(columns=['_imputed'])
            
            # ---> NUOVO: Salvataggio in CSV dei successi <---
            df_export.to_csv(out_dir / f'{out_stem}_geocodificati.csv', index=False, encoding='utf-8-sig')
            print(f'\n✅ File CSV "Geocodificati" creato.')
            
            # Applichiamo i colori usando la matrice
            df_styled = df_export.style.apply(lambda x: colora_matrix(df_export, df_completo), axis=None)
            df_styled.to_excel(writer, sheet_name='Geocodificati', index=False)
            print(f'\n✅ Foglio "Geocodificati" creato.')
        
        df_fail = pd.DataFrame(failed)
        if not df_fail.empty:
            
            # ---> NUOVO: Salvataggio in CSV dei fallimenti <---
            df_fail.to_csv(out_dir / f'{out_stem}_falliti.csv', index=False, encoding='utf-8-sig')
            print(f'⚠️ File CSV "Falliti" creato.')
            
            df_fail.to_excel(writer, sheet_name='Falliti', index=False)
            print(f'⚠️ Foglio "Falliti" aggiunto al file.')

    if results:
        features = [
            {
                'type':'Feature',
                'geometry': {'type': 'Point', 'coordinates': [r['lon'], r['lat']]},
                'properties': {k: v for k, v in r.items() if k not in ('lat', 'lon', 'Categoria_Geografica', '_imputed')}
            } for r in results
        ]
        with open(out_dir / f'{out_stem}_geocoded.geojson', 'w', encoding='utf-8') as f:
            # Aggiunto indent=4 per formattare il file in verticale e renderlo leggibile
            json.dump({'type':'FeatureCollection', 'features': features}, f, ensure_ascii=False, indent=4)

    print(f'\n📊 Elaborazione completata: {out_excel.name}')

if __name__ == '__main__':
    print("=== GEOCODER GIS BATCH (V7: Imputation & Color Matrix) ===")
    print("\n---------------- LEGENDA COLORI ----------------")
    print("🟩 VERDE    -> Capoluogo o Grande Città Italiana")
    print("🟨 GIALLO   -> Provincia o Piccolo Comune Italiano")
    print("🟥 ROSSO    -> Estero (Internazionale)")
    print("🟦 AZZURRO  -> Dato mancante riparato dallo script")
    print("------------------------------------------------\n")
    
    path = input('File CSV/XLSX: ').strip().strip('"').strip("'")
    print("\n--- IMPOSTAZIONI COLONNE ---")
    c_id = input('0. Colonna ID Univoco (opzionale): ').strip()
    c_addr = input('1. Colonna Indirizzo: ').strip()
    c_cap = input('2. Colonna CAP (opzionale): ').strip()
    c_city = input('3. Colonna Città (opzionale): ').strip()
    
    print("\nAvvio processo...\n")
    run_batch(path, c_id, c_addr, c_city, c_cap)
