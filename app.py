import streamlit as st
import pandas as pd
import sqlite3
import datetime
import docx
import os
import io

# Configuração da Página
st.set_page_config(page_title="Gestão de Frota - 10ª CICOM", layout="wide", page_icon="🚔")

DB_FILE = "frota_10cicom.db"

# --- CONTROLE DE ACESSO (LOGIN) ---
SENHA_PADRAO = "10cicom"  # Altere aqui se desejar outra senha

def check_login():
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False

    if not st.session_state.autenticado:
        st.markdown("<h2 style='text-align: center;'>🚔 Gestão de Frota - 10ª CICOM</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center;'>Acesso restrito ao efetivo autorizado</p>", unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1, 1.2, 1])
        with col2:
            with st.form("form_login"):
                senha = st.text_input("Digite a Senha de Acesso", type="password")
                entrar = st.form_submit_button("Entrar no Sistema", use_container_width=True)
                if entrar:
                    if senha == SENHA_PADRAO:
                        st.session_state.autenticado = True
                        st.rerun()
                    else:
                        st.error("Senha incorreta. Tente novamente.")
        return False
    return True

# --- BANCO DE DADOS ---
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
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
    c.execute("""
    CREATE TABLE IF NOT EXISTS abastecimentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data TEXT,
        prefixo TEXT,
        motorista TEXT,
        km_atual REAL,
        litros_liberados TEXT,
        horario TEXT,
        UNIQUE(data, prefixo, km_atual, horario)
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS ocorrencias (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data_hora TEXT,
        prefixo TEXT,
        tipo TEXT,
        descricao TEXT,
        motorista TEXT
    )
    """)
    
    # Frota padrão
    frota = [
        ("25-1001", "S10", "Diesel", "TRX-6I85", "Ativa", 30000, 40000),
        ("25-1111", "S10", "Diesel", "TRX-4B85", "Ativa", 50000, 60000),
        ("25-1329", "Spin", "Gasolina", "TRZ-7E17", "Ativa", 0, 10000),
        ("25-1353", "Spin", "Gasolina", "TSC-2D46", "Emprestada", 0, 10000),
        ("25-1394", "Spin", "Gasolina", "UTS-3J57", "Ativa", 0, 10000)
    ]
    c.executemany("INSERT OR IGNORE INTO viaturas VALUES (?, ?, ?, ?, ?, ?, ?)", frota)
    
    # Remove duplicidades se houver
    c.execute("""
        DELETE FROM abastecimentos 
        WHERE id NOT IN (
            SELECT MIN(id) 
            FROM abastecimentos 
            GROUP BY data, prefixo, km_atual, horario
        )
    """)
    conn.commit()
    conn.close()

def importar_excel_direto(caminho_ou_arquivo, conn):
    raw_df = pd.read_excel(caminho_ou_arquivo, header=None)
    vtr_rows = {4: "25-1001", 5: "25-1111", 6: "25-1329", 7: "25-1353", 8: "25-1394"}
    date_cols = [c for c in range(len(raw_df.columns)) if pd.notna(raw_df.iloc[2, c])]
    novos_inseridos = 0
    cursor = conn.cursor()
    
    for c in date_cols:
        d_val = raw_df.iloc[2, c]
        data_str = d_val.strftime("%Y-%m-%d") if isinstance(d_val, pd.Timestamp) else str(d_val)[:10]
        for r, prefixo in vtr_rows.items():
            km = raw_df.iloc[r, c]
            motorista = raw_df.iloc[r, c+1] if c+1 < len(raw_df.columns) else ""
            qtde = raw_df.iloc[r, c+2] if c+2 < len(raw_df.columns) else ""
            horario = raw_df.iloc[r, c+3] if c+3 < len(raw_df.columns) else ""
            
            if pd.notna(km) and str(km).strip() not in ['', 'nan', 'N/I', 'km']:
                try:
                    km_num = float(str(km).replace(',', '.').strip())
                    cursor.execute("""
                        INSERT OR IGNORE INTO abastecimentos (data, prefixo, motorista, km_atual, litros_liberados, horario)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (data_str, prefixo, str(motorista).strip(), km_num, str(qtde).strip(), str(horario).strip()))
                    if cursor.rowcount > 0:
                        novos_inseridos += 1
                except ValueError:
                    pass
    conn.commit()
    return novos_inseridos

def extrair_texto_word(uploaded_file):
    doc = docx.Document(uploaded_file)
    texto_completo = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n".join(texto_completo)

def converter_df_para_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatorio')
    return output.getvalue()

init_db()

# Verifica se o usuário está autenticado
if not check_login():
    st.stop()

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("🚔 10ª CICOM")
    st.write("Sistema Integrado de Frota")
    menu = st.radio("Navegação", [
        "📊 Painel de Controle (Dashboard)",
        "⛽ Controle de Abastecimento", 
        "🛠️ Status da Frota & Revisões", 
        "📋 Livro de Ocorrências / Avarias", 
        "🔎 Consulta & Exportação Excel", 
        "📥 Importar Planilha Excel"
    ])
    st.write("---")
    if st.button("🚪 Sair do Sistema"):
        st.session_state.autenticado = False
        st.rerun()

conn = sqlite3.connect(DB_FILE)

# 1. PAINEL DE CONTROLE (DASHBOARD)
if menu == "📊 Painel de Controle (Dashboard)":
    st.subheader("📊 Indicadores Gerais de Consumo e Frota")
    
    df_abast = pd.read_sql_query("SELECT * FROM abastecimentos", conn)
    df_vtr = pd.read_sql_query("SELECT * FROM viaturas", conn)
    
    if not df_abast.empty:
        # Extrai os litros como número
        df_abast['litros_num'] = df_abast['litros_liberados'].str.replace('L', '').str.replace('l', '').astype(float, errors='ignore')
        df_abast['litros_num'] = pd.to_numeric(df_abast['litros_num'], errors='coerce').fillna(0)
        
        total_litros = df_abast['litros_num'].sum()
        total_abast = len(df_abast)
        vtrs_ativas = len(df_vtr[df_vtr['status'] == 'Ativa'])
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Abastecimentos", f"{total_abast} registros")
        col2.metric("Total de Combustível Liberado", f"{total_litros:.0f} Litros")
        col3.metric("Viaturas Ativas", f"{vtrs_ativas} viaturas")
        
        st.write("---")
        
        c_graf1, c_graf2 = st.columns(2)
        with c_graf1:
            st.write("##### ⛽ Consumo Total de Litros por Viatura")
            consumo_vtr = df_abast.groupby('prefixo')['litros_num'].sum()
            st.bar_chart(consumo_vtr)
            
        with c_graf2:
            st.write("##### 📈 Quantidade de Abastecimentos por Viatura")
            qtd_vtr = df_abast['prefixo'].value_counts()
            st.bar_chart(qtd_vtr)
    else:
        st.info("Nenhum dado de abastecimento cadastrado para gerar o painel.")

# 2. ABASTECIMENTO
elif menu == "⛽ Controle de Abastecimento":
    st.subheader("⛽ Lançamento e Liberação Diária de Combustível")
    
    col1, col2 = st.columns(2)
    with col1:
        prefixo = st.selectbox("Selecione a Viatura", ["25-1001", "25-1111", "25-1329", "25-1353", "25-1394"])
        motorista = st.text_input("Nome do Motorista (Farol)")
        km_atual = st.number_input("Quilometragem Atual (KM)", min_value=0.0, step=1.0)
    with col2:
        autonomia = st.number_input("Autonomia Restante informada no painel (km)", min_value=0.0, step=1.0)
        hora = st.time_input("Horário do Abastecimento", datetime.datetime.now().time())
        data = st.date_input("Data", datetime.date.today())

    liberado = False
    litros = "0L"
    msg = ""

    if prefixo == "25-1353":
        msg = "⚠️ Viatura 25-1353 está emprestada/cedida para outra unidade. Abastecimento bloqueado."
    elif prefixo in ["25-1001", "25-1111"]:
        if 0 < autonomia <= 70:
            liberado = True
            litros = "60L"
            msg = "✅ COTA LIBERADA: 60 Litros de Diesel (Autonomia: <= 70km)"
        elif 70 < autonomia <= 180:
            liberado = True
            litros = "50L"
            msg = "✅ COTA LIBERADA: 50 Litros de Diesel (Autonomia: até 180km)"
        else:
            msg = f"❌ Bloqueado: Autonomia ({autonomia}km) acima do limite (>180km)."
    else:
        if 0 < autonomia <= 70:
            liberado = True
            litros = "40L"
            msg = "✅ COTA LIBERADA: 40 Litros de Gasolina (Autonomia: <= 70km)"
        elif 95 <= autonomia <= 115:
            liberado = True
            litros = "30L"
            msg = "✅ COTA LIBERADA: 30 Litros de Gasolina (Autonomia entre 95 e 115km)"
        elif 115 < autonomia <= 125:
            liberado = True
            litros = "35L"
            msg = "✅ COTA LIBERADA: 35 Litros de Gasolina (Autonomia entre 115 e 125km)"
        else:
            msg = f"❌ Bloqueado: Autonomia ({autonomia}km) fora das faixas permitidas."

    if autonomia > 0:
        if liberado:
            st.success(msg)
        else:
            st.error(msg)

    if st.button("Confirmar Lançamento", use_container_width=True):
        if not motorista:
            st.warning("Informe o nome do motorista.")
        elif not liberado:
            st.error("Não é possível registrar abastecimento bloqueado pela regra de cota.")
        else:
            c = conn.cursor()
            c.execute("INSERT OR IGNORE INTO abastecimentos (data, prefixo, motorista, km_atual, litros_liberados, horario) VALUES (?, ?, ?, ?, ?, ?)",
                      (data.strftime("%Y-%m-%d"), prefixo, motorista, km_atual, litros, hora.strftime("%H:%M")))
            conn.commit()
            st.success("Abastecimento registrado com sucesso!")

# 3. STATUS & REVISÕES
elif menu == "🛠️ Status da Frota & Revisões":
    st.subheader("🛠️ Monitoramento de Manutenção Preventiva")
    df_vtr = pd.read_sql_query("SELECT * FROM viaturas", conn)
    
    ultimos_kms = []
    restantes = []
    alertas = []
    for _, row in df_vtr.iterrows():
        pref = row['prefixo']
        res = conn.execute("SELECT km_atual FROM abastecimentos WHERE prefixo = ? ORDER BY km_atual DESC LIMIT 1", (pref,)).fetchone()
        km_atual = res[0] if res else 0.0
        prox_rev = row['km_proxima_revisao']
        km_restante = prox_rev - km_atual
        
        ultimos_kms.append(km_atual)
        restantes.append(km_restante)
        
        if row['status'] == 'Emprestada':
            alertas.append("Emprestada")
        elif km_restante <= 500:
            alertas.append("🚨 CRÍTICO: Revisão Urgente")
        elif km_restante <= 1500:
            alertas.append("⚠️ ATENÇÃO: Revisão Próxima")
        else:
            alertas.append("✅ Regular")

    df_vtr['KM Atual'] = ultimos_kms
    df_vtr['KM Restante p/ Revisão'] = restantes
    df_vtr['Status Revisão'] = alertas
    
    st.dataframe(df_vtr[['prefixo', 'modelo', 'placa', 'status', 'KM Atual', 'km_proxima_revisao', 'KM Restante p/ Revisão', 'Status Revisão']], use_container_width=True)

# 4. OCORRÊNCIAS / AVARIAS
elif menu == "📋 Livro de Ocorrências / Avarias":
    st.subheader("📋 Registro de Alterações, Arranhões e Sinistros")
    tab1, tab2 = st.tabs(["✍️ Digitação Manual", "📄 Importar do Word (.docx)"])
    
    with tab1:
        with st.form("form_ocorrencia_manual"):
            col1, col2 = st.columns(2)
            with col1:
                prefixo_oc = st.selectbox("Viatura", ["25-1001", "25-1111", "25-1329", "25-1353", "25-1394"], key="pref_m")
                tipo_oc = st.selectbox("Tipo de Ocorrência", ["Arranhão/Avaria Leve", "Sinistro/Colisão", "Problema Mecânico/Elétrico", "Cautela/Empréstimo", "Outro"], key="tipo_m")
                motorista_oc = st.text_input("Policial / Motorista Envolvido", key="mot_m")
            with col2:
                data_oc = st.date_input("Data do Ocorrido", datetime.date.today(), key="data_m")
                descricao = st.text_area("Descrição Detalhada", placeholder="Descreva o arranhão, sinistro ou alteração...", key="desc_m")
                
            if st.form_submit_button("Salvar Registro"):
                if not descricao:
                    st.warning("Preencha a descrição da alteração.")
                else:
                    c = conn.cursor()
                    c.execute("INSERT INTO ocorrencias (data_hora, prefixo, tipo, descricao, motorista) VALUES (?, ?, ?, ?, ?)",
                              (data_oc.strftime("%Y-%m-%d"), prefixo_oc, tipo_oc, descricao, motorista_oc))
                    conn.commit()
                    st.success("Ocorrência registrada com sucesso!")

    with tab2:
        st.write("Faça o upload do relatório salvo em formato Word (.docx):")
        arquivo_word = st.file_uploader("Selecione o arquivo Word", type=["docx"])
        
        if arquivo_word is not None:
            texto_extraido = extrair_texto_word(arquivo_word)
            st.text_area("Texto Lido do Arquivo:", value=texto_extraido, height=150)
            
            col_w1, col_w2 = st.columns(2)
            with col_w1:
                pref_w = st.selectbox("Viatura do Documento", ["25-1001", "25-1111", "25-1329", "25-1353", "25-1394"], key="pref_w")
                tipo_w = st.selectbox("Tipo", ["Avaria/Livro de Alterações", "Sinistro", "Manutenção", "Outro"], key="tipo_w")
            with col_w2:
                mot_w = st.text_input("Policial / Responsável", key="mot_w")
                data_w = st.date_input("Data", datetime.date.today(), key="data_w")
                
            if st.button("Gravar Histórico do Word no Sistema"):
                c = conn.cursor()
                c.execute("INSERT INTO ocorrencias (data_hora, prefixo, tipo, descricao, motorista) VALUES (?, ?, ?, ?, ?)",
                          (data_w.strftime("%Y-%m-%d"), pref_w, tipo_w, texto_extraido, mot_w))
                conn.commit()
                st.success("Conteúdo do Word gravado com sucesso no livro de alterações!")

    st.write("---")
    st.write("### Ocorrências Registradas")
    df_oc = pd.read_sql_query("SELECT data_hora as Data, prefixo as Viatura, tipo as Tipo, motorista as Envolvido, descricao as Detalhes FROM ocorrencias ORDER BY id DESC", conn)
    st.dataframe(df_oc, use_container_width=True)

# 5. CONSULTA & EXPORTAÇÃO EXCEL
elif menu == "🔎 Consulta & Exportação Excel":
    st.subheader("🔎 Busca Avançada e Exportação de Relatórios")
    
    df_abast = pd.read_sql_query("SELECT data as Data, prefixo as Viatura, motorista as Motorista, km_atual as [KM Odômetro], litros_liberados as [Qtd Liberada], horario as Horário FROM abastecimentos ORDER BY data DESC, id DESC", conn)
    
    if not df_abast.empty:
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            filtro_vtr = st.selectbox("Filtrar por Viatura", ["Todas"] + list(df_abast['Viatura'].unique()))
        with col_f2:
            filtro_motorista = st.text_input("Buscar por Motorista")
        with col_f3:
            filtro_data = st.date_input("Filtrar por Data Específica", value=None)
            
        df_filtrado = df_abast.copy()
        if filtro_vtr != "Todas":
            df_filtrado = df_filtrado[df_filtrado['Viatura'] == filtro_vtr]
        if filtro_motorista:
            df_filtrado = df_filtrado[df_filtrado['Motorista'].str.contains(filtro_motorista, case=False, na=False)]
        if filtro_data is not None:
            df_filtrado = df_filtrado[df_filtrado['Data'] == filtro_data.strftime("%Y-%m-%d")]
            
        st.write(f"**Registros encontrados:** {len(df_filtrado)}")
        st.dataframe(df_filtrado, use_container_width=True)
        
        # Botão para baixar planilha Excel
        excel_bytes = converter_df_para_excel(df_filtrado)
        st.download_button(
            label="📥 Baixar estes dados em Planilha Excel (.xlsx)",
            data=excel_bytes,
            file_name="Relatorio_Abastecimento_10CICOM.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Nenhum dado encontrado para consulta.")

# 6. IMPORTAR PLANILHA EXCEL
elif menu == "📥 Importar Planilha Excel":
    st.subheader("📥 Importar Dados de Planilha de Abastecimento (.xlsx)")
    st.write("Faça o envio de uma planilha no modelo da 10ª CICOM para carregar os registros no sistema:")
    arquivo_excel = st.file_uploader("Selecione o arquivo Excel (.xlsx)", type=["xlsx"])
    
    if arquivo_excel is not None:
        if st.button("Processar e Carregar Dados"):
            novos = importar_excel_direto(arquivo_excel, conn)
            if novos > 0:
                st.success(f"Foram importados {novos} novos registros de abastecimento!")
            else:
                st.info("Todos os registros desta planilha já constavam no sistema (nenhuma duplicidade foi inserida).")

conn.close()
