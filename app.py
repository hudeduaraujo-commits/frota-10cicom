import streamlit as st
import pandas as pd
import sqlite3
import os
import glob
import re
from datetime import datetime
from zoneinfo import ZoneInfo
import io

# -----------------------------------------------------------------------------
# CONFIGURAÇÃO GERAL E FUSO OFICIAL DE MANAUS (UTC-4)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Gestão de Frota - 10ª CICOM",
    page_icon="🚔",
    layout="wide",
    initial_sidebar_state="expanded"
)

FUSO_MANAUS = ZoneInfo("America/Manaus")

def agora_manaus():
    return datetime.now(FUSO_MANAUS)

NOME_BANCO = "frota_10cicom.db"
ARQUIVO_EXCEL = "kms_abastecimento_10cicom.xlsx"

# -----------------------------------------------------------------------------
# BANCO DE DADOS SQLITE & SCHEMA DEFENSIVO
# -----------------------------------------------------------------------------
def get_conexao():
    return sqlite3.connect(NOME_BANCO, check_same_thread=False)

def inicializar_banco():
    conn = get_conexao()
    cursor = conn.cursor()
    
    # Tabela de Viaturas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS viaturas (
            prefixo TEXT PRIMARY KEY,
            modelo TEXT,
            placa TEXT,
            combustivel TEXT,
            km_revisao_base INTEGER,
            intervalo_revisao INTEGER DEFAULT 10000,
            status TEXT DEFAULT 'Operacional'
        )
    """)
    
    # Tabela de Abastecimentos e Quilometragem
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS abastecimentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            prefixo TEXT,
            data TEXT,
            horario TEXT,
            km_atual INTEGER,
            litros REAL DEFAULT 0.0,
            motorista TEXT,
            placa TEXT,
            observacao TEXT,
            origem TEXT DEFAULT 'PLANILHA',
            FOREIGN KEY (prefixo) REFERENCES viaturas(prefixo)
        )
    """)

    # Tabela de Avarias e Ocorrências
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS ocorrencias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            prefixo TEXT,
            data TEXT,
            tipo TEXT,
            descricao TEXT,
            status TEXT DEFAULT 'Pendente',
            registrado_por TEXT,
            FOREIGN KEY (prefixo) REFERENCES viaturas(prefixo)
        )
    """)
    
    # Auto-migração para garantir compatibilidade de colunas
    cursor.execute("PRAGMA table_info(abastecimentos)")
    colunas_existentes = [c[1] for c in cursor.fetchall()]
    for col_nome, col_tipo in [
        ("horario", "TEXT"),
        ("km_atual", "INTEGER DEFAULT 0"),
        ("litros", "REAL DEFAULT 0.0"),
        ("motorista", "TEXT"),
        ("placa", "TEXT"),
        ("observacao", "TEXT"),
        ("origem", "TEXT DEFAULT 'PLANILHA'")
    ]:
        if col_nome not in colunas_existentes:
            cursor.execute(f"ALTER TABLE abastecimentos ADD COLUMN {col_nome} {col_tipo}")

    # Carga de Viaturas Ativas da 10ª CICOM
    cursor.execute("SELECT COUNT(*) FROM viaturas")
    if cursor.fetchone()[0] == 0:
        viaturas_iniciais = [
            ("25-1001", "Chevrolet S10", "TRX-6I85", "Diesel", 40000, 10000, "Operacional"),
            ("25-1111", "Chevrolet S10", "TRX-4B85", "Diesel", 40000, 10000, "Operacional"),
            ("25-1329", "Chevrolet Spin", "TRZ-7E17", "Gasolina", 50000, 10000, "Operacional"),
            ("25-1353", "Chevrolet Spin", "TSC-2D46", "Gasolina", 30000, 10000, "Operacional"),
            ("25-1394", "Chevrolet Spin", "UTS-3J57", "Gasolina", 20000, 10000, "Operacional")
        ]
        cursor.executemany("""
            INSERT OR IGNORE INTO viaturas (prefixo, modelo, placa, combustivel, km_revisao_base, intervalo_revisao, status)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, viaturas_iniciais)
        
    conn.commit()
    conn.close()

inicializar_banco()

# -----------------------------------------------------------------------------
# PARSER ROBUSTO PARA MATRIZ DE AGOSTO E SETEMBRO (143 COLUNAS)
# -----------------------------------------------------------------------------
def normalizar_qualquer_data(val, default_year="2026"):
    """Converte qualquer formato de data (Excel serial, texto, DD/MM, ISO) para YYYY-MM-DD."""
    if pd.isna(val) or val is None:
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%Y-%m-%d")
        
    # Trata data serial numérica do Excel (ex: 46266)
    try:
        f_val = float(val)
        if 40000 < f_val < 60000:
            dt = pd.to_datetime(f_val, unit='D', origin='1899-12-30')
            return dt.strftime("%Y-%m-%d")
    except:
        pass
        
    s = str(val).strip()
    
    # Formato ISO: YYYY-MM-DD
    m_iso = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
    if m_iso:
        return f"{m_iso.group(1)}-{m_iso.group(2).zfill(2)}-{m_iso.group(3).zfill(2)}"
        
    # Formato BR completo: DD/MM/YYYY
    m_br = re.search(r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', s)
    if m_br:
        return f"{m_br.group(3)}-{m_br.group(2).zfill(2)}-{m_br.group(1).zfill(2)}"
        
    # Formato curto: DD/MM (ex: 01/09 ou 12/08)
    m_curto = re.search(r'^(\d{1,2})[-/](\d{1,2})$', s)
    if m_curto:
        dia = m_curto.group(1).zfill(2)
        mes = m_curto.group(2).zfill(2)
        return f"{default_year}-{mes}-{dia}"
        
    # Formato textual: 10-Aug, 1-Set, 04-Sep
    meses_map = {
        "jan": "01", "fev": "02", "feb": "02", "mar": "03", "abr": "04", "apr": "04",
        "mai": "05", "may": "05", "jun": "06", "jul": "07", "ago": "08", "aug": "08",
        "set": "09", "sep": "09", "out": "10", "oct": "10", "nov": "11", "dez": "12", "dec": "12"
    }
    m_txt = re.search(r'(\d{1,2})[-/ ]([A-Za-z]{3})', s)
    if m_txt:
        dia = m_txt.group(1).zfill(2)
        mes = meses_map.get(m_txt.group(2).lower(), "08")
        return f"{default_year}-{mes}-{dia}"
        
    return None

def resolver_data_da_coluna(df_raw, col_km):
    """Localiza a data do bloco de KM buscando nas linhas 1 e 2 e colunas vizinhas."""
    # Tenta na linha 2 (onde estão as datas diárias)
    for offset in [0, -1, 1, -2, 2, -3, 3, -4]:
        c = col_km + offset
        if 0 <= c < len(df_raw.columns):
            val2 = df_raw.iloc[2, c]
            d_fmt = normalizar_qualquer_data(val2)
            if d_fmt:
                return d_fmt
            # Tenta na linha 1 caso esteja mesclado acima
            val1 = df_raw.iloc[1, c]
            d_fmt1 = normalizar_qualquer_data(val1)
            if d_fmt1:
                return d_fmt1
    return "2026-08-10"

def extrair_registros_10cicom(df_raw):
    records = []
    
    # 1. Mapeamento das viaturas e placas fixas (colunas 1 e 2 a partir da linha 4)
    placa_map = {}
    for r in range(4, len(df_raw)):
        v_num = str(df_raw.iloc[r, 1]).strip()
        p_val = str(df_raw.iloc[r, 2]).strip()
        if v_num.isdigit():
            placa_map[v_num] = p_val if p_val.upper() not in ['NAN', 'NONE'] else ''

    # 2. Identifica TODAS as colunas que representam KM na linha 3 (cobre agosto e setembro)
    km_cols = []
    for c in range(len(df_raw.columns)):
        val_header = str(df_raw.iloc[3, c]).strip().lower()
        if val_header == 'km' or val_header.startswith('km'):
            km_cols.append(c)

    # 3. Varre todos os blocos diários
    for c_km in km_cols:
        data_registro = resolver_data_da_coluna(df_raw, c_km)

        c_vtr = c_km - 1
        c_farol = c_km + 1
        c_qtde = c_km + 2
        c_hora = c_km + 3

        for r in range(4, len(df_raw)):
            vtr_raw = str(df_raw.iloc[r, c_vtr]).strip()
            if not vtr_raw.isdigit():
                vtr_raw = str(df_raw.iloc[r, 1]).strip()
            if not vtr_raw.isdigit():
                continue

            prefixo = f"25-{vtr_raw}"
            placa = placa_map.get(vtr_raw, "")

            km_raw = str(df_raw.iloc[r, c_km]).strip()
            km_clean = km_raw.replace('.', '').replace(',', '').replace(' ', '')

            # Identifica leitura de odômetro válida
            if km_clean.isdigit() and int(km_clean) > 100:
                km_val = int(km_clean)

                motorista = "Turno Serviço"
                if c_farol < len(df_raw.columns):
                    m_val = str(df_raw.iloc[r, c_farol]).strip()
                    if m_val.upper() not in ['NAN', 'NONE', 'N/A', 'N/I', '']:
                        motorista = m_val

                litros = 0.0
                if c_qtde < len(df_raw.columns):
                    q_val = str(df_raw.iloc[r, c_qtde]).strip()
                    l_nums = re.findall(r'\d+', q_val)
                    if l_nums:
                        litros = float(l_nums[0])

                horario = "07:00"
                if c_hora < len(df_raw.columns):
                    h_val = str(df_raw.iloc[r, c_hora]).strip()
                    if h_val.upper() not in ['NAN', 'NONE', 'N/A', 'N/I', ''] and ':' in h_val:
                        horario = h_val[:5]

                records.append({
                    "prefixo": prefixo,
                    "placa": placa,
                    "data": data_registro,
                    "km_atual": km_val,
                    "motorista": motorista,
                    "litros": litros,
                    "horario": horario
                })

    return records

def salvar_no_banco(registros, limpar_antes=False):
    conn = get_conexao()
    cursor = conn.cursor()
    
    if limpar_antes:
        cursor.execute("DELETE FROM abastecimentos WHERE origem = 'PLANILHA'")
        conn.commit()

    novos, duplicados = 0, 0
    for reg in registros:
        cursor.execute("""
            SELECT COUNT(*) FROM abastecimentos 
            WHERE prefixo = ? AND data = ? AND km_atual = ?
        """, (reg["prefixo"], reg["data"], reg["km_atual"]))

        if cursor.fetchone()[0] == 0:
            cursor.execute("""
                INSERT INTO abastecimentos (prefixo, data, horario, km_atual, litros, motorista, placa, origem)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'PLANILHA')
            """, (reg["prefixo"], reg["data"], reg["horario"], reg["km_atual"], reg["litros"], reg["motorista"], reg["placa"]))
            novos += 1
        else:
            duplicados += 1

    conn.commit()
    conn.close()
    return novos, duplicados

def processar_planilha_local(limpar_antes=False):
    arqs = glob.glob(f"{ARQUIVO_EXCEL}*") + glob.glob("*abastecimento*10*.*")
    if not arqs:
        return None, "Arquivo kms_abastecimento_10cicom.xlsx não encontrado."
    arq = arqs[0]
    try:
        xl = pd.ExcelFile(arq)
        aba_alvo = next((s for s in xl.sheet_names if "KMS" in s.upper() or "ABASTECIMENTO" in s.upper()), xl.sheet_names[0])
        df_raw = xl.parse(aba_alvo, header=None)
        registros = extrair_registros_10cicom(df_raw)
        if not registros:
            return None, "Nenhum registro extraído da planilha."
        novos, duplicados = salvar_no_banco(registros, limpar_antes=limpar_antes)
        return (arq, aba_alvo, len(registros), novos, duplicados), None
    except Exception as e:
        return None, str(e)

# -----------------------------------------------------------------------------
# PARSER DE MENSAGENS DO WHATSAPP
# -----------------------------------------------------------------------------
def extrair_dados_whatsapp(texto):
    dados = {
        "prefixo": None,
        "km": 0,
        "litros": 0.0,
        "motorista": "",
        "horario": agora_manaus().strftime("%H:%M"),
        "observacao": ""
    }
    
    # 1. Prefixo
    vtr_match = re.search(r'(?:VTR|VIATURA|PREFIXO)[:\s]*([A-Za-z0-9\-\*]+)', texto, re.IGNORECASE)
    if vtr_match:
        nums = re.findall(r'\d{4}', vtr_match.group(1))
        if nums:
            dados["prefixo"] = f"25-{nums[0]}"
    if not dados["prefixo"]:
        for p in ["1001", "1111", "1329", "1353", "1394"]:
            if p in texto:
                dados["prefixo"] = f"25-{p}"
                break
                
    # 2. KM
    km_match = re.search(r'(?:KM|ODOMETRO|QUILOMETRAGEM)[:\s]*([0-9\.\,]+)', texto, re.IGNORECASE)
    if km_match:
        km_clean = km_match.group(1).replace('.', '').replace(',', '')
        if km_clean.isdigit():
            dados["km"] = int(km_clean)
            
    # 3. Litros
    l_match = re.search(r'(?:LITROS|LT|LTS|QTDE|QUANTIDADE)[:\s]*([0-9\.\,]+)', texto, re.IGNORECASE)
    if l_match:
        try:
            dados["litros"] = float(l_match.group(1).replace(',', '.'))
        except:
            pass
            
    # 4. Motorista
    mot_match = re.search(r'(?:MOTORISTA|CONDUTOR|POLICIAL|FAROL)[:\s]*([^\n\r]+)', texto, re.IGNORECASE)
    if mot_match:
        dados["motorista"] = mot_match.group(1).strip()
        
    # 5. Horário
    h_match = re.search(r'(?:HORA|HORARIO|HORÁRIO)[:\s]*([0-9]{1,2}[:hH][0-9]{2})', texto, re.IGNORECASE)
    if h_match:
        h = h_match.group(1).replace('h', ':').replace('H', ':')
        if len(h) == 4 and h[1] == ':':
            h = '0' + h
        dados["horario"] = h[:5]
        
    # 6. Observação / Avaria
    obs_match = re.search(r'(?:OBS|AVARIA|ALTERACAO|ALTERAÇÃO|OBSERVACAO|OBSERVAÇÃO)[:\s]*([^\n\r]+)', texto, re.IGNORECASE)
    if obs_match:
        dados["observacao"] = obs_match.group(1).strip()
        
    return dados

# Carga automática se o banco estiver vazio
conn_c = get_conexao()
qtd = conn_c.cursor().execute("SELECT COUNT(*) FROM abastecimentos").fetchone()[0]
conn_c.close()
if qtd == 0:
    processar_planilha_local()

# -----------------------------------------------------------------------------
# CONSULTAS AUXILIARES
# -----------------------------------------------------------------------------
def obter_viaturas():
    conn = get_conexao()
    df = pd.read_sql_query("SELECT * FROM viaturas ORDER BY prefixo ASC", conn)
    conn.close()
    return df

def obter_historico(prefixo=None):
    conn = get_conexao()
    if prefixo and prefixo != "Todas":
        q = "SELECT id, prefixo, data, horario, km_atual, litros, motorista, placa, observacao, origem FROM abastecimentos WHERE prefixo = ? ORDER BY data DESC, km_atual DESC, id DESC"
        df = pd.read_sql_query(q, conn, params=(prefixo,))
    else:
        q = "SELECT id, prefixo, data, horario, km_atual, litros, motorista, placa, observacao, origem FROM abastecimentos ORDER BY data DESC, km_atual DESC, id DESC"
        df = pd.read_sql_query(q, conn)
    conn.close()
    return df

def obter_ocorrencias(status_filtro="Todas"):
    conn = get_conexao()
    if status_filtro != "Todas":
        q = "SELECT * FROM ocorrencias WHERE status = ? ORDER BY id DESC"
        df = pd.read_sql_query(q, conn, params=(status_filtro,))
    else:
        q = "SELECT * FROM ocorrencias ORDER BY id DESC"
        df = pd.read_sql_query(q, conn)
    conn.close()
    return df

# -----------------------------------------------------------------------------
# INTERFACE DO USUÁRIO (STREAMLIT)
# -----------------------------------------------------------------------------
st.title("🚔 Frota Operacional — 10ª CICOM")
st.caption(f"Horário Oficial de Manaus: {agora_manaus().strftime('%d/%m/%Y %H:%M:%S')} (Fuso UTC-4)")

df_todos = obter_historico()

# BARRA LATERAL COM BOTÕES DE SINCRONIZAÇÃO COMPLETA
with st.sidebar:
    st.header("⚙️ Painel Operacional")
    filtro_vtr = st.selectbox("Filtrar Viatura:", ["Todas"] + obter_viaturas()['prefixo'].tolist())
    st.write("---")
    st.subheader("📥 Sincronização")
    
    if st.button("🔄 Recarregar Novos Dados", use_container_width=True):
        res, erro = processar_planilha_local(limpar_antes=False)
        if erro:
            st.error(f"Erro: {erro}")
        else:
            arq, aba, total, novos, dups = res
            st.success(f"Sincronizado! {total} registros mapeados (+{novos} novos).")
            st.rerun()

    if st.button("⚡ Limpar e Reimportar Tudo", type="primary", use_container_width=True, help="Limpa o banco e reimporta todos os dias de Agosto e Setembro da planilha"):
        res, erro = processar_planilha_local(limpar_antes=True)
        if erro:
            st.error(f"Erro: {erro}")
        else:
            arq, aba, total, novos, dups = res
            st.success(f"✅ Banco recarregado: {novos} lançamentos de Agosto e Setembro importados!")
            st.rerun()

# ABAS DO SISTEMA
tab_dash, tab_whatsapp, tab_avarias, tab_hist, tab_lanca, tab_manut = st.tabs([
    "📊 Visão Geral & Odômetros",
    "📲 Ler Texto do WhatsApp",
    "⚠️ Avarias & Ocorrências",
    "📑 Histórico Cronológico",
    "⛽ Lançamento Manual",
    "🛠️ Revisões Preventivas"
])

# 1. DASHBOARD
with tab_dash:
    df_vtrs = obter_viaturas()
    st.subheader("Odômetro Mais Recente da Frota (Agosto e Setembro)")
    cols = st.columns(len(df_vtrs))
    
    for idx, vtr in df_vtrs.iterrows():
        p = vtr['prefixo']
        sub = df_todos[df_todos['prefixo'] == p]
        km_ultimo = sub['km_atual'].iloc[0] if not sub.empty else 0
        litros_total = sub['litros'].sum() if not sub.empty else 0
        data_recente = sub['data'].iloc[0] if not sub.empty else "Sem registro"
        
        with cols[idx]:
            st.metric(
                label=f"VTR {p}",
                value=f"{km_ultimo:,} km".replace(",", "."),
                delta=f"{litros_total:.0f} L totais"
            )
            st.caption(f"**{vtr['modelo']}** | Placa: `{vtr['placa']}`\nÚltimo reg: `{data_recente}`")

    st.write("---")
    st.subheader("Últimos Lançamentos Registrados")
    st.dataframe(df_todos.head(15), use_container_width=True)

# 2. LEITURA DE TEXTO DO WHATSAPP (ENTRADA DE SERVIÇO / ABASTECIMENTO)
with tab_whatsapp:
    st.subheader("📲 Leitura Automática de Texto do WhatsApp")
    st.markdown("""
    Cole abaixo a mensagem enviada pelo motorista no grupo de WhatsApp da 10ª CICOM. 
    O sistema identificará a **Viatura, Odômetro, Litragem, Motorista, Horário** e eventuais **Avarias**.
    """)
    
    txt_colado = st.text_area(
        "Cole a mensagem do WhatsApp aqui:",
        height=140,
        placeholder="Exemplo:\n*ASSUNÇÃO DE SERVIÇO*\nVTR: 1001\nKM: 36.200\nLitros: 50L\nMotorista: CB PM SILVA\nHorário: 08:30\nObs: Farol esquerdo queimado"
    )
    
    if st.button("🔍 Ler e Extrair Dados da Mensagem", type="primary", use_container_width=True):
        if txt_colado.strip():
            st.session_state['dados_zap'] = extrair_dados_whatsapp(txt_colado)
            st.success("✅ Dados extraídos com sucesso! Confira e confirme abaixo:")
        else:
            st.warning("⚠️ Cole uma mensagem antes de clicar no botão.")
            
    if 'dados_zap' in st.session_state:
        dz = st.session_state['dados_zap']
        df_vtrs = obter_viaturas()
        lista_vtrs = df_vtrs['prefixo'].tolist()
        
        idx_vtr = 0
        if dz['prefixo'] in lista_vtrs:
            idx_vtr = lista_vtrs.index(dz['prefixo'])
            
        with st.form("form_confirma_zap"):
            st.markdown("##### Dados Extraídos para Validação:")
            c1, c2, c3 = st.columns(3)
            with c1:
                vtr_sel = st.selectbox("Viatura:", lista_vtrs, index=idx_vtr)
                data_sel = st.date_input("Data do Registro:", value=agora_manaus().date())
            with c2:
                km_val = st.number_input("Odômetro (KM):", value=dz['km'], step=1)
                litros_val = st.number_input("Litros Abastecidos:", value=float(dz['litros']), step=0.1, format="%.2f")
            with c3:
                motorista_val = st.text_input("Motorista:", value=dz['motorista'] or "Turno Serviço")
                hora_val = st.text_input("Horário (Manaus):", value=dz['horario'])
                
            obs_val = st.text_input("Observação / Avaria Reportada:", value=dz['observacao'])
            
            salvar_zap = st.form_submit_button("💾 Confirmar e Gravar no Banco de Dados", use_container_width=True)
            if salvar_zap:
                conn = get_conexao()
                cur = conn.cursor()
                cur.execute("""
                    INSERT INTO abastecimentos (prefixo, data, horario, km_atual, litros, motorista, observacao, origem)
                    VALUES (?, ?, ?, ?, ?, ?, ?, 'WHATSAPP')
                """, (vtr_sel, data_sel.strftime('%Y-%m-%d'), hora_val, int(km_val), float(litros_val), motorista_val, obs_val))
                
                if obs_val.strip():
                    cur.execute("""
                        INSERT INTO ocorrencias (prefixo, data, tipo, descricao, status, registrado_por)
                        VALUES (?, ?, 'Avaria / Alteração', ?, 'Pendente', ?)
                    """, (vtr_sel, data_sel.strftime('%Y-%m-%d'), obs_val, motorista_val))
                    
                conn.commit()
                conn.close()
                st.success(f"✅ Lançamento da VTR {vtr_sel} gravado com sucesso!")
                del st.session_state['dados_zap']
                st.rerun()

# 3. AVARIAS & OCORRÊNCIAS
with tab_avarias:
    st.subheader("⚠️ Controle de Avarias e Ocorrências das Viaturas")
    
    col_cad, col_lista = st.columns([1, 1.6])
    
    with col_cad:
        st.markdown("#### 📝 Registrar Nova Avaria")
        df_vtrs = obter_viaturas()
        with st.form("form_avaria", clear_on_submit=True):
            av_vtr = st.selectbox("Viatura:", df_vtrs['prefixo'].tolist())
            av_data = st.date_input("Data da Ocorrência:", value=agora_manaus().date())
            av_tipo = st.selectbox("Tipo de Alteração:", [
                "Mecânica / Motor",
                "Elétrica / Farol / Sirene",
                "Pneus / Alinhamento",
                "Funilaria / Pintura / Batida",
                "Ar-condicionado",
                "Documentação / Cartão",
                "Outros"
            ])
            av_desc = st.text_area("Descrição do Problema:", placeholder="Ex: Pneu dianteiro direito furado ou farol queimado.")
            av_policial = st.text_input("Policial / Condutor Relatante:", placeholder="Ex: SGT PM CARDOSO")
            
            if st.form_submit_button("🚨 Salvar Registro de Avaria", use_container_width=True):
                if av_desc.strip():
                    conn = get_conexao()
                    cur = conn.cursor()
                    cur.execute("""
                        INSERT INTO ocorrencias (prefixo, data, tipo, descricao, status, registrado_por)
                        VALUES (?, ?, ?, ?, 'Pendente', ?)
                    """, (av_vtr, av_data.strftime('%Y-%m-%d'), av_tipo, av_desc.strip(), av_policial.strip()))
                    conn.commit()
                    conn.close()
                    st.success(f"Avaria da VTR {av_vtr} registrada com sucesso!")
                    st.rerun()
                else:
                    st.warning("Descreva o problema antes de salvar.")

    with col_lista:
        st.markdown("#### 📋 Avarias Registradas")
        filtro_status_avaria = st.radio("Exibir:", ["Todas", "Pendente", "Solucionado"], horizontal=True)
        df_ocorr = obter_ocorrencias(filtro_status_avaria)
        
        if df_ocorr.empty:
            st.info("Nenhuma avaria cadastrada neste status.")
        else:
            for _, row in df_ocorr.iterrows():
                id_oc = row['id']
                status_icon = "🔴" if row['status'] == "Pendente" else "🟢"
                with st.expander(f"{status_icon} VTR {row['prefixo']} | {row['tipo']} ({row['data']})"):
                    st.write(f"**Descrição:** {row['descricao']}")
                    st.write(f"**Relatado por:** {row['registrado_por']}")
                    st.write(f"**Status Atual:** {row['status']}")
                    
                    if row['status'] == 'Pendente':
                        if st.button(f"✅ Marcar como Solucionado / Reparado", key=f"btn_res_{id_oc}"):
                            conn = get_conexao()
                            cur = conn.cursor()
                            cur.execute("UPDATE ocorrencias SET status = 'Solucionado' WHERE id = ?", (id_oc,))
                            conn.commit()
                            conn.close()
                            st.success("Ocorrência atualizada para SOLUCIONADO!")
                            st.rerun()

# 4. HISTÓRICO COMPLETO (AGORÁ COM TODAS AS DATAS E ORDENAÇÃO EXATA)
with tab_hist:
    st.subheader(f"Histórico Completo de Abastecimentos ({filtro_vtr})")
    df_filtrado = obter_historico(filtro_vtr)
    
    if not df_filtrado.empty:
        total_regs = len(df_filtrado)
        data_min = df_filtrado['data'].min()
        data_max = df_filtrado['data'].max()
        st.caption(f"Exibindo **{total_regs}** lançamentos de **{data_min}** até **{data_max}**.")
        
    st.dataframe(df_filtrado, use_container_width=True)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name='Abastecimentos')
        
    st.download_button(
        label="📥 Baixar Histórico em Excel (.xlsx)",
        data=buffer.getvalue(),
        file_name=f"frota_10cicom_{agora_manaus().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# 5. LANÇAMENTO MANUAL
with tab_lanca:
    st.subheader("Novo Lançamento Manual (Manaus UTC-4)")
    with st.form("form_manual", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            sel_vtr = st.selectbox("Viatura:", obter_viaturas()['prefixo'].tolist())
            dt_lanc = st.date_input("Data:", value=agora_manaus().date())
        with c2:
            hr_lanc = st.time_input("Horário:", value=agora_manaus().time())
            mot_lanc = st.text_input("Policial / Motorista:", placeholder="Ex: SD RAMON")
        with c3:
            km_lanc = st.number_input("Odômetro (KM):", min_value=0, step=1)
            lt_lanc = st.number_input("Litros Abastecidos:", min_value=0.0, step=0.1, format="%.2f")

        obs_lanc = st.text_input("Observação / Alteração:")

        if st.form_submit_button("💾 Salvar Registro", use_container_width=True):
            conn = get_conexao()
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO abastecimentos (prefixo, data, horario, km_atual, litros, motorista, observacao, origem)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'MANUAL')
            """, (sel_vtr, dt_lanc.strftime('%Y-%m-%d'), hr_lanc.strftime('%H:%M'), int(km_lanc), float(lt_lanc), mot_lanc, obs_lanc))
            conn.commit()
            conn.close()
            st.success(f"VTR {sel_vtr} gravada com sucesso!")
            st.rerun()

# 6. MANUTENÇÃO PREVENTIVA
with tab_manut:
  st.subheader('🛠️ Controle de Revisões Preventivas (Intervalo: 10.000 KM)')

  # 1. Garantir correção das Spins 1329 e 1394 no banco de dados (Base = 0 para a 1ª revisão aos 10.000 km)
  conn_rev = get_conexao()
  cur_rev = conn_rev.cursor()
  cur_rev.execute(
      "UPDATE viaturas SET km_revisao_base = 0, intervalo_revisao = 10000 WHERE"
      " prefixo IN ('25-1329', '25-1394') AND km_revisao_base > 10000"
  )
  conn_rev.commit()
  conn_rev.close()

  df_vtrs = obter_viaturas()
  df_todos = obter_historico()

  # 2. Tabela de Monitoramento da Frota
  lista_m = []
  for _, vtr in df_vtrs.iterrows():
    p = vtr['prefixo']
    sub = df_todos[df_todos['prefixo'] == p]
    km_atual = (
        sub['km_atual'].iloc[0] if not sub.empty else vtr['km_revisao_base']
    )
    base = vtr['km_revisao_base']
    intervalo = (
        vtr['intervalo_revisao'] if vtr['intervalo_revisao'] > 0 else 10000
    )

    # Cálculo da próxima revisão
    if km_atual < intervalo and base == 0:
      prox = 10000  # 1ª Revisão de Fábrica
    else:
      prox = (
          base + (((km_atual - base) // intervalo) + 1) * intervalo
          if km_atual >= base
          else base
      )

    restante = prox - km_atual

    if restante <= 0:
      status = '🔴 VENCIDA / URGENTE'
    elif restante <= 1000:
      status = '🟡 ALERTA (< 1.000 km)'
    else:
      status = '🟢 REGULAR'

    lista_m.append({
        'Viatura': p,
        'Modelo': vtr['modelo'],
        'KM Atual': f'{km_atual:,}'.replace(',', '.'),
        'Última Revisão (Base)': (
            f'{base:,} km'.replace(',', '.') if base > 0 else 'Zero KM (Nova)'
        ),
        'Próxima Revisão': f'{prox:,}'.replace(',', '.'),
        'Faltam': f'{restante:,} km'.replace(',', '.'),
        'Situação': status,
    })

  st.dataframe(pd.DataFrame(lista_m), use_container_width=True)

  st.write('---')

  # 3. Formulário para Atualizar Parâmetros de Revisão Manualmente
  st.markdown('#### ⚙️ Atualizar Dados de Revisão da Viatura')
  st.caption(
      'Use para registrar uma revisão que acabou de ser realizada ou ajustar a'
      ' quilometragem base.'
  )

  with st.form('form_ajuste_revisao'):
    col_v, col_base, col_int = st.columns(3)
    with col_v:
      vtr_escolhida = st.selectbox(
          'Selecione a Viatura:', df_vtrs['prefixo'].tolist()
      )
    with col_base:
      km_base_novo = st.number_input(
          'KM da Última Revisão Feita (ou 0 se ainda não fez a 1ª):',
          min_value=0,
          step=1000,
          value=0,
      )
    with col_int:
      int_novo = st.number_input(
          'Intervalo de Manutenção (Padrão: 10.000):',
          min_value=1000,
          step=1000,
          value=10000,
      )

    if st.form_submit_button(
        '💾 Atualizar Parâmetros da Viatura', use_container_width=True
    ):
      conn_up = get_conexao()
      cur_up = conn_up.cursor()
      cur_up.execute(
          """
                UPDATE viaturas 
                SET km_revisao_base = ?, intervalo_revisao = ?
                WHERE prefixo = ?
            """,
          (int(km_base_novo), int(int_novo), vtr_escolhida),
      )
      conn_up.commit()
      conn_up.close()
      st.success(
          f'✅ Parâmetros da VTR {vtr_escolhida} atualizados com sucesso!'
      )
      st.rerun()
      
