import sqlite3
import pandas as pd
import os

DB_FILE = "frota_10cicom.db"

# Procura o arquivo Excel na sua pasta
arquivo_encontrado = None
for f in os.listdir('.'):
    if f.endswith('.xlsx') and not f.startswith('~$'):
        arquivo_encontrado = f
        break

if not arquivo_encontrado:
    print("❌ Nenhum arquivo Excel (.xlsx) encontrado na pasta!")
else:
    print(f"📄 Lendo arquivo encontrado: {arquivo_encontrado}")
    raw_df = pd.read_excel(arquivo_encontrado, header=None)
    vtr_rows = {4: "25-1001", 5: "25-1111", 6: "25-1329", 7: "25-1353", 8: "25-1394"}
    date_cols = [c for c in range(len(raw_df.columns)) if pd.notna(raw_df.iloc[2, c])]
    
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    
    inseridos = 0
    for col in date_cols:
        d_val = raw_df.iloc[2, col]
        data_str = d_val.strftime("%Y-%m-%d") if isinstance(d_val, pd.Timestamp) else str(d_val)[:10]
        for r, prefixo in vtr_rows.items():
            km = raw_df.iloc[r, col]
            motorista = raw_df.iloc[r, col+1] if col+1 < len(raw_df.columns) else ""
            qtde = raw_df.iloc[r, col+2] if col+2 < len(raw_df.columns) else ""
            horario = raw_df.iloc[r, col+3] if col+3 < len(raw_df.columns) else ""
            
            if pd.notna(km) and str(km).strip() not in ['', 'nan', 'N/I', 'km']:
                try:
                    km_num = float(str(km).replace(',', '.').strip())
                    c.execute("""
                        INSERT OR IGNORE INTO abastecimentos (data, prefixo, motorista, km_atual, litros_liberados, horario)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (data_str, prefixo, str(motorista).strip(), km_num, str(qtde).strip(), str(horario).strip()))
                    if c.rowcount > 0:
                        inseridos += 1
                except ValueError:
                    pass
                    
    conn.commit()
    conn.close()
    print(f"✅ SUCESSO! {inseridos} registros foram carregados diretamente para o banco de dados.")
    