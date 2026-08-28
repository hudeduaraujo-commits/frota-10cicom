import pandas as pd
import sqlite3
import datetime

def importar_planilha_excel(caminho_excel="KMS DE ABASTECIMENTOS VTRs 10 (1).xlsx"):
    raw_df = pd.read_excel(caminho_excel, sheet_name=0, header=None)
    conn = sqlite3.connect("frota_10cicom.db")
    cursor = conn.cursor()

    # 1. Atualiza Cadastro de Viaturas com as Placas
    frota = [
        ("25-1001", "S10", "Diesel", "TRX-6I85", "Ativa", 30000, 40000),
        ("25-1111", "S10", "Diesel", "TRX-4B85", "Ativa", 50000, 60000),
        ("25-1329", "Spin", "Gasolina", "TRZ-7E17", "Ativa", 0, 10000),
        ("25-1353", "Spin", "Gasolina", "TSC-2D46", "Emprestada", 0, 10000),
        ("25-1394", "Spin", "Gasolina", "UTS-3J57", "Ativa", 0, 10000)
    ]
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS viaturas (
        prefixo TEXT PRIMARY KEY,
        modelo TEXT,
        combustivel TEXT,
        placa TEXT,
        status TEXT,
        km_ultima_revisao REAL,
        km_proxima_revisao REAL
    )
    """)
    cursor.executemany("INSERT OR REPLACE INTO viaturas VALUES (?, ?, ?, ?, ?, ?, ?)", frota)

    # 2. Estrutura a Tabela de Abastecimentos
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS abastecimentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data TEXT,
        prefixo TEXT,
        motorista TEXT,
        km_atual REAL,
        litros_liberados TEXT,
        horario TEXT
    )
    """)

    # Mapeamento das linhas do Excel
    vtr_rows = {
        4: "25-1001",
        5: "25-1111",
        6: "25-1329",
        7: "25-1353",
        8: "25-1394"
    }

    # As colunas de datas iniciam na coluna 3 e saltam de 5 em 5
    date_cols = [c for c in range(len(raw_df.columns)) if pd.notna(raw_df.iloc[2, c])]

    abastecimentos = []
    for c in date_cols:
        d_val = raw_df.iloc[2, c]
        data_str = d_val.strftime("%Y-%m-%d") if isinstance(d_val, (pd.Timestamp, datetime.datetime)) else str(d_val)[:10]

        for r, prefixo in vtr_rows.items():
            km = raw_df.iloc[r, c]
            motorista = raw_df.iloc[r, c+1] if c+1 < len(raw_df.columns) else ""
            qtde = raw_df.iloc[r, c+2] if c+2 < len(raw_df.columns) else ""
            horario = raw_df.iloc[r, c+3] if c+3 < len(raw_df.columns) else ""

            if pd.notna(km) and str(km).strip() not in ['', 'nan', 'N/I']:
                abastecimentos.append((
                    data_str,
                    prefixo,
                    str(motorista).strip(),
                    float(str(km).replace(',', '.')),
                    str(qtde).strip(),
                    str(horario).strip()
                ))

    cursor.executemany("""
        INSERT INTO abastecimentos (data, prefixo, motorista, km_atual, litros_liberados, horario)
        VALUES (?, ?, ?, ?, ?, ?)
    """, abastecimentos)

    conn.commit()
    conn.close()
    print(f"Sucesso: {len(abastecimentos)} abastecimentos importados da planilha!")

if __name__ == "__main__":
    importar_planilha_excel()