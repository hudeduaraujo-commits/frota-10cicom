import streamlit as st
import pandas as pd
import sqlite3
import os
import glob
import re
from datetime import datetime
from zoneinfo import ZoneInfo
import io
from PIL import Image

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
    
    # 1. Tabela de Viaturas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS viaturas (
            prefixo TEXT PRIMARY KEY,
            modelo TEXT,
            placa TEXT,
            combustivel TEXT,
            km_revisao_base INTEGER DEFAULT 0,
            intervalo_revisao INTEGER DEFAULT 10000,
            status TEXT DEFAULT 'Operacional'
        )
    """)
    
    # 2. Tabela de Abastecimentos
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
            origem TEXT DEFAULT 'MANUAL',
            FOREIGN KEY (prefixo) REFERENCES viaturas(prefixo)
        )
    """)

    # 3. Tabela de Assunções de Serviço
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS assuncoes_servico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            prefixo TEXT,
            data TEXT,
            horario TEXT,
            turno TEXT,
            cmt_guarnicao TEXT,
            motorista TEXT,
            km_inicial INTEGER,
            observacao TEXT,
            origem TEXT DEFAULT 'MANUAL',
            FOREIGN KEY (prefixo) REFERENCES viaturas(prefixo)
        )
    """)

    # 4. Tabela de Avarias e Ocorrências
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

    # 5. Tabela de Histórico de Revisões Realizadas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historico_revisoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            prefixo TEXT,
            data TEXT,
            km_revisao INTEGER,
            tipo_revisao TEXT,
            oficina TEXT,
            responsavel TEXT,
            observacoes TEXT,
            FOREIGN KEY (prefixo) REFERENCES viaturas(prefixo)
        )
    """)
    
    # Carga inicial de Viaturas Ativas
    cursor.execute("SELECT COUNT(*) FROM viaturas")
    if cursor.fetchone()[0] == 0:
        viaturas_iniciais = [
            ("25-1001", "Chevrolet S10", "TRX-6I85", "Diesel", 40000, 10000, "Operacional"),
            ("25-1111", "Chevrolet S10", "TRX-4B85", "Diesel", 40000, 10000, "Operacional"),
            ("25-1329", "Chevrolet Spin", "TRZ-7E17", "Gasolina", 0, 10000, "Operacional"),
            ("25-1353", "Chevrolet Spin", "TSC-2D46", "Gasolina", 30000, 10000, "Operacional"),
            ("25-1394", "Chevrolet Spin", "UTS-3J57", "Gasolina", 0, 10000, "Operacional")
        ]
        cursor.executemany("""
            INSERT OR IGNORE INTO viaturas (prefixo, modelo, placa, combustivel, km_revisao_base, intervalo_revisao, status)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, viaturas_iniciais)
        
    conn.commit()
    conn.close()

inicializar_banco()

# -----------------------------------------------------------------------------
# REGRAS DE CÁLCULO DE REVISÃO PREVENTIVA
# -----------------------------------------------------------------------------
def calcular_status_revisao(km_atual, km_base, intervalo=10000):
    """Calcula a próxima revisão e quantos quilômetros restam."""
    if intervalo <= 0:
        intervalo = 10000

    if km_atual < intervalo and km_base == 0:
        prox = intervalo
    else:
        if km_atual >= km_base:
            ciclos = ((km_atual - km_base) // intervalo) + 1
            prox = km_base + (ciclos * intervalo)
        else:
            prox = km_base

    restante = prox - km_atual

    if restante <= 0:
        status = "🔴 VENCIDA / URGENTE"
        classe = "error"
    elif restante <= 1000:
        status = "🟡 ALERTA (< 1.000 km)"
        classe = "warning"
    else:
        status = "🟢 REGULAR"
        classe = "success"

    return prox, restante, status, classe

# -----------------------------------------------------------------------------
# PARSERS DE TEXTO E FOTO
# -----------------------------------------------------------------------------
def extrair_dados_whatsapp(texto):
    dados = {
        "prefixo": None,
        "km": 0,
        "litros": 0.0,
        "motorista": "",
        "cmt_guarnicao": "",
        "turno": "1º Turno (07h às 19h)",
        "horario": agora_manaus().strftime("%H:%M"),
        "observacao": ""
    }
    
    # Prefixo
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
                
    # Odômetro / KM
    km_match = re.search(r'(?:KM|ODOMETRO|QUILOMETRAGEM)[:\s]*([0-9\.\,]+)', texto, re.IGNORECASE)
    if km_match:
        km_clean = km_match.group(1).replace('.', '').replace(',', '')
        if km_clean.isdigit():
            dados["km"] = int(km_clean)
            
    # Litros
    l_match = re.search(r'(?:LITROS|LT|LTS|QTDE|QUANTIDADE)[:\s]*([0-9\.\,]+)', texto, re.IGNORECASE)
    if l_match:
        try:
            dados["litros"] = float(l_match.group(1).replace(',', '.'))
        except:
            pass
            
    # Motorista
    mot_match = re.search(r'(?:MOTORISTA|CONDUTOR|POLICIAL|FAROL)[:\s]*([^\n\r]+)', texto, re.IGNORECASE)
    if mot_match:
        dados["motorista"] = mot_match.group(1).strip()
        
    # Comandante de Guarnição
    cmt_match = re.search(r'(?:CMT|COMANDANTE|CHEFE|ENCARREGADO)[:\s]*([^\n\r]+)', texto, re.IGNORECASE)
    if cmt_match:
        dados["cmt_guarnicao"] = cmt_match.group(1).strip()

    # Horário
    h_match = re.search(r'(?:HORA|HORARIO|HORÁRIO)[:\s]*([0-9]{1,2}[:hH][0-9]{2})', texto, re.IGNORECASE)
    if h_match:
        h = h_match.group(1).replace('h', ':').replace('H', ':')
        if len(h) == 4 and h[1] == ':':
            h = '0' + h
        dados["horario"] = h[:5]
        
    # Observação / Avaria
    obs_match = re.search(r'(?:OBS|AVARIA|ALTERACAO|ALTERAÇÃO|OBSERVACAO|OBSERVAÇÃO)[:\s]*([^\n\r]+)', texto, re.IGNORECASE)
    if obs_match:
        dados["observacao"] = obs_match.group(1).strip()
        
    return dados

def extrair_km_imagem_ocr(imagem_bytes):
    try:
        import pytesseract
        img = Image.open(io.BytesIO(imagem_bytes))
        texto_ocr = pytesseract.image_to_string(img, config='--psm 6 digits')
        numeros = re.findall(r'\b\d{4,6}\b', texto_ocr)
        if numeros:
            return int(numeros[0])
    except:
        pass
    return 0

# -----------------------------------------------------------------------------
# CONSULTAS AUXILIARES
# -----------------------------------------------------------------------------
def obter_viaturas():
    conn = get_conexao()
    df = pd.read_sql_query("SELECT * FROM viaturas ORDER BY prefixo ASC", conn)
    conn.close()
    return df

def obter_historico_abastecimentos(prefixo=None):
    conn = get_conexao()
    if prefixo and prefixo != "Todas":
        q = "SELECT id, prefixo, data, horario, km_atual, litros, motorista, observacao, origem FROM abastecimentos WHERE prefixo = ? ORDER BY data DESC, km_atual DESC, id DESC"
        df = pd.read_sql_query(q, conn, params=(prefixo,))
    else:
        q = "SELECT id, prefixo, data, horario, km_atual, litros, motorista, observacao, origem FROM abastecimentos ORDER BY data DESC, km_atual DESC, id DESC"
        df = pd.read_sql_query(q, conn)
    conn.close()
    return df

def obter_historico_assuncoes(prefixo=None):
    conn = get_conexao()
    if prefixo and prefixo != "Todas":
        q = "SELECT id, prefixo, data, horario, turno, cmt_guarnicao, motorista, km_inicial, observacao, origem FROM assuncoes_servico WHERE prefixo = ? ORDER BY data DESC, km_inicial DESC, id DESC"
        df = pd.read_sql_query(q, conn, params=(prefixo,))
    else:
        q = "SELECT id, prefixo, data, horario, turno, cmt_guarnicao, motorista, km_inicial, observacao, origem FROM assuncoes_servico ORDER BY data DESC, km_inicial DESC, id DESC"
        df = pd.read_sql_query(q, conn)
    conn.close()
    return df

def obter_historico_revisoes(prefixo=None):
    conn = get_conexao()
    if prefixo and prefixo != "Todas":
        q = "SELECT * FROM historico_revisoes WHERE prefixo = ? ORDER BY data DESC, km_revisao DESC"
        df = pd.read_sql_query(q, conn, params=(prefixo,))
    else:
        q = "SELECT * FROM historico_revisoes ORDER BY data DESC, km_revisao DESC"
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
# INTERFACE DO USUÁRIO
# -----------------------------------------------------------------------------
st.title("🚔 Frota Operacional — 10ª CICOM")
st.caption(f"Horário Oficial de Manaus: {agora_manaus().strftime('%d/%m/%Y %H:%M:%S')} (Fuso UTC-4)")

df_vtrs = obter_viaturas()
lista_prefixos = df_vtrs['prefixo'].tolist()

# BARRA LATERAL COM AÇÕES DIRETAS
with st.sidebar:
    st.header("⚙️ Painel Operacional")
    filtro_vtr = st.selectbox("Filtrar Viatura:", ["Todas"] + lista_prefixos)
    st.write("---")
    st.subheader("🎯 Ações Rápidas")
    st.info("Utilize as abas superiores para registrar Abastecimentos, Assunções de Serviço ou Lançamento de Revisões.")

# ABAS DO SISTEMA
abas_nomes = [
    "📊 Visão Geral & Odômetros",
    "🛡️ Assunção de Serviço",
    "⛽ Registrar Abastecimento",
    "🛠️ Registrar Revisão",
    "⚠️ Avarias & Ocorrências",
    "📑 Histórico Completo",
    "🔍 Quadro Geral de Revisões"
]

tab_dash, tab_assuncao, tab_abast, tab_reg_rev, tab_avarias, tab_hist, tab_manut = st.tabs(abas_nomes)

# -----------------------------------------------------------------------------
# 1. DASHBOARD GERAL
# -----------------------------------------------------------------------------
with tab_dash:
    df_todos_abast = obter_historico_abastecimentos()
    st.subheader("Odômetro Mais Recente da Frota & Situação de Revisão")
    cols = st.columns(len(df_vtrs))
    
    for idx, vtr in df_vtrs.iterrows():
        p = vtr['prefixo']
        sub = df_todos_abast[df_todos_abast['prefixo'] == p]
        km_ultimo = sub['km_atual'].iloc[0] if not sub.empty else vtr['km_revisao_base']
        litros_total = sub['litros'].sum() if not sub.empty else 0
        data_recente = sub['data'].iloc[0] if not sub.empty else "Sem registro"
        
        prox_r, rest_r, st_r, _ = calcular_status_revisao(km_ultimo, vtr['km_revisao_base'], vtr['intervalo_revisao'])

        with cols[idx]:
            st.metric(
                label=f"VTR {p}",
                value=f"{km_ultimo:,} km".replace(",", "."),
                delta=f"Faltam {rest_r:,} km".replace(",", ".") if rest_r > 0 else f"Vencida há {abs(rest_r):,} km".replace(",", ".")
            )
            st.caption(f"**{vtr['modelo']}** | `{vtr['placa']}`\nRevisão: **{st_r}**\nÚltimo reg: `{data_recente}`")

    st.write("---")
    col_dash1, col_dash2 = st.columns(2)
    with col_dash1:
        st.markdown("##### 🛡️ Últimas Assunções de Serviço")
        df_ass = obter_historico_assuncoes()
        st.dataframe(df_ass.head(6), use_container_width=True)
    with col_dash2:
        st.markdown("##### ⛽ Últimos Abastecimentos")
        st.dataframe(df_todos_abast.head(6), use_container_width=True)

# -----------------------------------------------------------------------------
# 2. REGISTRO DE ASSUNÇÃO DE SERVIÇO COM CÁLCULO IMEDIATO DE REVISÃO
# -----------------------------------------------------------------------------
with tab_assuncao:
    st.subheader("🛡️ Registro de Assunção de Serviço da Viatura")
    st.caption("Cadastre a entrada de serviço. O sistema calcula na hora se a viatura está dentro do limite de revisão.")

    metodo_assuncao = st.radio(
        "Método de Registro da Assunção:",
        ["✍️ Manual", "📷 Leitura de Foto do Painel", "💬 Texto do WhatsApp"],
        horizontal=True,
        key="radio_metodo_assuncao"
    )

    dados_assuncao = {
        "prefixo": lista_prefixos[0] if lista_prefixos else "",
        "km": 0,
        "turno": "1º Turno (07h às 19h)",
        "cmt_guarnicao": "",
        "motorista": "",
        "horario": agora_manaus().strftime("%H:%M"),
        "data": agora_manaus().date(),
        "observacao": "",
        "origem": "MANUAL"
    }

    if metodo_assuncao == "💬 Texto do WhatsApp":
        txt_assuncao = st.text_area(
            "Cole a mensagem padrão de assunção de serviço do WhatsApp:",
            height=120,
            placeholder="Exemplo:\n*ASSUNÇÃO DE SERVIÇO*\nTurno: 1º Turno (07h às 19h)\nVTR: 1001\nKM Inicial: 44.200\nCmt: SGT PM CARDOSO\nMotorista: CB PM SILVA\nObs: Farol dianteiro direito com lâmpada fraca"
        )
        if st.button("🔍 Processar Mensagem WhatsApp", key="btn_proc_assuncao"):
            if txt_assuncao.strip():
                st.session_state['assuncao_ext'] = extrair_dados_whatsapp(txt_assuncao)
                st.success("Mensagem processada com sucesso!")

        if 'assuncao_ext' in st.session_state:
            e = st.session_state['assuncao_ext']
            if e['prefixo'] in lista_prefixos:
                dados_assuncao['prefixo'] = e['prefixo']
            dados_assuncao['km'] = e['km']
            dados_assuncao['cmt_guarnicao'] = e['cmt_guarnicao']
            dados_assuncao['motorista'] = e['motorista']
            dados_assuncao['horario'] = e['horario']
            dados_assuncao['observacao'] = e['observacao']
            dados_assuncao['origem'] = "WHATSAPP"

    elif metodo_assuncao == "📷 Leitura de Foto do Painel":
        c_p1, c_p2 = st.columns([1, 1])
        with c_p1:
            foto_painel_ass = st.file_uploader("Upload da foto do odômetro na assunção:", type=['jpg', 'jpeg', 'png'], key="up_foto_ass")
            foto_camera_ass = st.camera_input("Ou fotografe o odômetro pelo celular:", key="cam_foto_ass")
            
            foto_usar_ass = foto_camera_ass or foto_painel_ass
            if foto_usar_ass:
                km_detectado = extrair_km_imagem_ocr(foto_usar_ass.getvalue())
                if km_detectado > 0:
                    dados_assuncao['km'] = km_detectado
                    st.success(f"Odômetro detectado: {km_detectado:,} KM".replace(',', '.'))
                dados_assuncao['origem'] = "FOTO_PAINEL"
        with c_p2:
            if foto_usar_ass:
                st.image(foto_usar_ass, caption="Foto do Painel anexada", use_container_width=True)

    st.write("---")
    st.markdown("##### 📝 Formulário de Confirmação da Assunção de Serviço")
    
    c1, c2, c3 = st.columns(3)
    with c1:
        idx_vtr_a = lista_prefixos.index(dados_assuncao['prefixo']) if dados_assuncao['prefixo'] in lista_prefixos else 0
        vtr_selecionada = st.selectbox("Viatura:", lista_prefixos, index=idx_vtr_a, key="ass_vtr_sel")
        turno_selecionado = st.selectbox("Turno de Serviço:", ["1º Turno (07h às 19h)", "2º Turno (19h às 07h)", "Turno Extra / Operação Especial"], key="ass_turno_sel")
    with c2:
        data_digitada = st.date_input("Data:", value=dados_assuncao['data'], key="ass_dt_sel")
        hora_digitada = st.text_input("Horário (Manaus):", value=dados_assuncao['horario'], key="ass_hr_sel")
    with c3:
        km_digitado = st.number_input("KM Inicial (Odômetro Atual):", value=int(dados_assuncao['km']), step=1, key="ass_km_num")
        cmt_digitado = st.text_input("Comandante da Guarnição:", value=dados_assuncao['cmt_guarnicao'], placeholder="Ex: TEN PM MOURA / SGT PM CARDOSO", key="ass_cmt_txt")

    mot_digitado = st.text_input("Motorista / Condutor:", value=dados_assuncao['motorista'], placeholder="Ex: CB PM SILVA", key="ass_mot_txt")
    obs_digitada = st.text_area("Observações da Viatura / Avarias Verificadas no Início:", value=dados_assuncao['observacao'], placeholder="Ex: Nível de óleo e água conferidos; lanterna esquerda trincada.", key="ass_obs_txt")

    # -------------------------------------------------------------------------
    # CÁLCULO AUTOMÁTICO DE REVISÃO EXIBIDO EM TEMPO REAL ANTES DE GRAVAR
    # -------------------------------------------------------------------------
    if km_digitado > 0:
        info_vtr = df_vtrs[df_vtrs['prefixo'] == vtr_selecionada].iloc[0]
        base_vtr = info_vtr['km_revisao_base']
        int_vtr = info_vtr['intervalo_revisao']
        
        prox_rev, faltam_km, sit_rev, classe_rev = calcular_status_revisao(km_digitado, base_vtr, int_vtr)

        st.markdown("#### 🔍 Diagnóstico Automático de Manutenção Preventiva:")
        c_diag1, c_diag2, c_diag3 = st.columns(3)
        c_diag1.metric("KM Informado pelo Motorista", f"{km_digitado:,} km".replace(",", "."))
        c_diag2.metric("Próxima Revisão Prevista", f"{prox_rev:,} km".replace(",", "."))
        
        if faltam_km <= 0:
            c_diag3.metric("Situação", sit_rev, delta=f"Vencida há {abs(faltam_km):,} km".replace(",", "."), delta_color="inverse")
            st.error(f"🚨 **ATENÇÃO GUARNICAO / COMANDO:** A VTR {vtr_selecionada} ultrapassou o limite para revisão em **{abs(faltam_km):,} KM**. Solicitar encaminhamento à oficina credenciada.")
        elif faltam_km <= 1000:
            c_diag3.metric("Faltam para Revisão", f"{faltam_km:,} km".replace(",", "."), delta="Atenção: Menos de 1.000 km!", delta_color="inverse")
            st.warning(f"⚠️ **ALERTA DE REVISÃO PRÓXIMA:** Faltam apenas **{faltam_km:,} KM** para a próxima revisão preventiva da VTR {vtr_selecionada}.")
        else:
            c_diag3.metric("Faltam para Revisão", f"{faltam_km:,} km".replace(",", "."), delta="Em dia / Regular")
            st.success(f"✅ **SITUAÇÃO REGULAR:** A VTR {vtr_selecionada} possui **{faltam_km:,} KM** disponíveis até a próxima revisão.")

    if st.button("💾 Confirmar e Homologar Assunção de Serviço", type="primary", use_container_width=True, key="btn_gravar_assuncao"):
        conn = get_conexao()
        cur = conn.cursor()
        
        # 1. Grava na tabela de assunções
        cur.execute("""
            INSERT INTO assuncoes_servico (prefixo, data, horario, turno, cmt_guarnicao, motorista, km_inicial, observacao, origem)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (vtr_selecionada, data_digitada.strftime('%Y-%m-%d'), hora_digitada, turno_selecionado, cmt_digitado, mot_digitado, int(km_digitado), obs_digitada, dados_assuncao['origem']))
        
        # 2. Alimenta o odômetro no histórico geral para manter a frota sincronizada
        cur.execute("""
            INSERT INTO abastecimentos (prefixo, data, horario, km_atual, litros, motorista, observacao, origem)
            VALUES (?, ?, ?, ?, 0.0, ?, ?, ?)
        """, (vtr_selecionada, data_digitada.strftime('%Y-%m-%d'), hora_digitada, int(km_digitado), mot_digitado, f"Assunção ({turno_selecionado}) - {obs_digitada}".strip(), dados_assuncao['origem']))

        # 3. Registra avaria se houver descrição de alteração
        if obs_digitada.strip() and obs_digitada.lower() not in ['sem alteracao', 'sem alteração', 'ok', 'tudo ok']:
            cur.execute("""
                INSERT INTO ocorrencias (prefixo, data, tipo, descricao, status, registrado_por)
                VALUES (?, ?, 'Avaria / Assunção', ?, 'Pendente', ?)
            """, (vtr_selecionada, data_digitada.strftime('%Y-%m-%d'), obs_digitada, mot_digitado or cmt_digitado))

        conn.commit()
        conn.close()
        st.success(f"✅ Assunção de Serviço da VTR {vtr_selecionada} homologada com sucesso!")
        if 'assuncao_ext' in st.session_state:
            del st.session_state['assuncao_ext']
        st.rerun()

# -----------------------------------------------------------------------------
# 3. REGISTRO DE ABASTECIMENTO
# -----------------------------------------------------------------------------
with tab_abast:
    st.subheader("⛽ Registro de Abastecimento de Viaturas")
    st.caption("Cadastre os abastecimentos efetuados via Manual, Leitura de Foto da Bomba/Painel ou Mensagem do WhatsApp.")

    metodo_abast = st.radio(
        "Método de Registro de Abastecimento:",
        ["✍️ Manual", "📷 Leitura de Foto da Bomba / Painel", "💬 Texto do WhatsApp"],
        horizontal=True,
        key="radio_metodo_abast"
    )

    dados_abast = {
        "prefixo": lista_prefixos[0] if lista_prefixos else "",
        "km": 0,
        "litros": 0.0,
        "motorista": "",
        "horario": agora_manaus().strftime("%H:%M"),
        "data": agora_manaus().date(),
        "observacao": "",
        "origem": "MANUAL"
    }

    if metodo_abast == "💬 Texto do WhatsApp":
        txt_abast = st.text_area(
            "Cole a mensagem de abastecimento do WhatsApp:",
            height=110,
            placeholder="Exemplo:\nVTR: 1001\nKM: 42.150\nLitros: 55L\nMotorista: CB PM SOUZA\nObs: Posto Equador Djalma Batista"
        )
        if st.button("🔍 Processar Mensagem Abastecimento", key="btn_proc_abast"):
            if txt_abast.strip():
                st.session_state['abast_ext'] = extrair_dados_whatsapp(txt_abast)
                st.success("Dados do abastecimento extraídos com sucesso!")

        if 'abast_ext' in st.session_state:
            e = st.session_state['abast_ext']
            if e['prefixo'] in lista_prefixos:
                dados_abast['prefixo'] = e['prefixo']
            dados_abast['km'] = e['km']
            dados_abast['litros'] = e['litros']
            dados_abast['motorista'] = e['motorista']
            dados_abast['horario'] = e['horario']
            dados_abast['observacao'] = e['observacao']
            dados_abast['origem'] = "WHATSAPP"

    elif metodo_abast == "📷 Leitura de Foto da Bomba / Painel":
        col_f1, col_f2 = st.columns([1, 1])
        with col_f1:
            foto_abast = st.file_uploader("Upload da foto da bomba ou odômetro:", type=['jpg', 'jpeg', 'png'], key="up_foto_abast")
            foto_cam_abast = st.camera_input("Fotografar bomba ou odômetro:", key="cam_foto_abast")
            
            foto_usar_a = foto_cam_abast or foto_abast
            if foto_usar_a:
                km_detectado_a = extrair_km_imagem_ocr(foto_usar_a.getvalue())
                if km_detectado_a > 0:
                    dados_abast['km'] = km_detectado_a
                    st.success(f"Odômetro detectado: {km_detectado_a:,} KM".replace(',', '.'))
                dados_abast['origem'] = "FOTO_PAINEL"
        with col_f2:
            if foto_usar_a:
                st.image(foto_usar_a, caption="Comprovante / Foto do Abastecimento", use_container_width=True)

    with st.form("form_abastecimento"):
        st.markdown("##### 📝 Formulário de Gravação de Abastecimento")
        c1, c2, c3 = st.columns(3)
        with c1:
            idx_vtr = lista_prefixos.index(dados_abast['prefixo']) if dados_abast['prefixo'] in lista_prefixos else 0
            vtr_sel = st.selectbox("Viatura:", lista_prefixos, index=idx_vtr, key="abast_vtr")
            data_sel = st.date_input("Data do Abastecimento:", value=dados_abast['data'], key="abast_data")
        with c2:
            km_val = st.number_input("Odômetro (KM):", value=int(dados_abast['km']), step=1, key="abast_km")
            litros_val = st.number_input("Litros Abastecidos:", value=float(dados_abast['litros']), step=0.1, format="%.2f", key="abast_lt")
        with c3:
            motorista_val = st.text_input("Motorista / Condutor:", value=dados_abast['motorista'], placeholder="Ex: CB PM ALMEIDA", key="abast_mot")
            horario_val = st.text_input("Horário (Manaus):", value=dados_abast['horario'], key="abast_hr")

        obs_val = st.text_input("Observação / Posto / Tipo de Combustível:", value=dados_abast['observacao'], key="abast_obs")

        if st.form_submit_button("💾 Salvar Registro de Abastecimento", type="primary", use_container_width=True):
            conn = get_conexao()
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO abastecimentos (prefixo, data, horario, km_atual, litros, motorista, observacao, origem)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (vtr_sel, data_sel.strftime('%Y-%m-%d'), horario_val, int(km_val), float(litros_val), motorista_val, obs_val, dados_abast['origem']))
            
            if any(p in obs_val.lower() for p in ['avaria', 'quebrado', 'pneu', 'furado', 'luz', 'farol', 'defeito']):
                cur.execute("""
                    INSERT INTO ocorrencias (prefixo, data, tipo, descricao, status, registrado_por)
                    VALUES (?, ?, 'Avaria / Abastecimento', ?, 'Pendente', ?)
                """, (vtr_sel, data_sel.strftime('%Y-%m-%d'), obs_val, motorista_val))

            conn.commit()
            conn.close()
            st.success(f"✅ Abastecimento da VTR {vtr_sel} gravado com sucesso!")
            if 'abast_ext' in st.session_state:
                del st.session_state['abast_ext']
            st.rerun()

# -----------------------------------------------------------------------------
# 4. REGISTRO DE REVISÃO REALIZADA (NOVO BOTÃO / FLUXO OPERACIONAL)
# -----------------------------------------------------------------------------
with tab_reg_rev:
    st.subheader("🛠️ Registrar Revisão Mecânica Realizada")
    st.caption("Use esta área para homologar a revisão preventiva feita na concessionária ou oficina conveniada. A quilometragem base da viatura será atualizada automaticamente.")

    with st.form("form_revisao_feita", clear_on_submit=True):
        c_r1, c_r2, c_r3 = st.columns(3)
        with c_r1:
            rev_vtr = st.selectbox("Selecione a Viatura Revisada:", lista_prefixos)
            rev_tipo = st.selectbox("Tipo de Revisão:", [
                "Preventiva Programada (10.000 km)",
                "Preventiva de Fábrica (1ª Revisão)",
                "Corretiva / Suspensão / Freios",
                "Troca de Óleo e Filtros Avulsa",
                "Revisão Geral de Sistema Elétrico"
            ])
        with c_r2:
            rev_data = st.date_input("Data da Realização da Revisão:", value=agora_manaus().date())
            rev_km = st.number_input("Odômetro da Viatura na Revisão (KM):", min_value=0, step=1000, help="Esta quilometragem se tornará a nova base para calcular a próxima revisão.")
        with c_r3:
            rev_oficina = st.text_input("Oficina / Concessionária:", placeholder="Ex: Pedragon Chevrolet / Oficina PMAM")
            rev_resp = st.text_input("Policial Responsável pela Entrega / Recebimento:", placeholder="Ex: SGT PM MOURA (Armeiro/Transporte)")

        rev_obs = st.text_area("Serviços Executados e Peças Trocadas (OS):", placeholder="Ex: Troca de óleo do motor, filtro de óleo, filtro de combustível, pastilhas de freio dianteiras e alinhamento.")

        if st.form_submit_button("✅ Homologar e Gravar Revisão no Sistema", type="primary", use_container_width=True):
            if rev_km > 0:
                conn = get_conexao()
                cur = conn.cursor()
                
                # 1. Registra no histórico de revisões
                cur.execute("""
                    INSERT INTO historico_revisoes (prefixo, data, km_revisao, tipo_revisao, oficina, responsavel, observacoes)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (rev_vtr, rev_data.strftime('%Y-%m-%d'), int(rev_km), rev_tipo, rev_oficina, rev_resp, rev_obs))
                
                # 2. Atualiza a quilometragem base da viatura para recalibrar o cálculo automático
                cur.execute("""
                    UPDATE viaturas 
                    SET km_revisao_base = ? 
                    WHERE prefixo = ?
                """, (int(rev_km), rev_vtr))
                
                conn.commit()
                conn.close()
                st.success(f"🎉 Revisão da VTR {rev_vtr} homologada com sucesso! Base recalibrada para {rev_km:,} KM.")
                st.rerun()
            else:
                st.warning("Informe o odômetro exato da realização da revisão para atualizar a base da viatura.")

    st.write("---")
    st.markdown("##### 📋 Histórico de Revisões Realizadas na 10ª CICOM")
    df_rev_historico = obter_historico_revisoes(filtro_vtr)
    if df_rev_historico.empty:
        st.info("Nenhuma revisão homologada cadastrada até o momento.")
    else:
        st.dataframe(df_rev_historico, use_container_width=True)

# -----------------------------------------------------------------------------
# 5. AVARIAS & OCORRÊNCIAS
# -----------------------------------------------------------------------------
with tab_avarias:
    st.subheader("⚠️ Controle de Avarias e Ocorrências das Viaturas")
    
    col_cad, col_lista = st.columns([1, 1.6])
    
    with col_cad:
        st.markdown("#### 📝 Registrar Nova Avaria")
        with st.form("form_avaria", clear_on_submit=True):
            av_vtr = st.selectbox("Viatura:", lista_prefixos)
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

# -----------------------------------------------------------------------------
# 6. HISTÓRICO COMPLETO (ABASTECIMENTOS, ASSUNÇÕES E REVISÕES)
# -----------------------------------------------------------------------------
with tab_hist:
    st.subheader(f"Histórico Geral da Frota ({filtro_vtr})")
    
    sub_tab1, sub_tab2, sub_tab3 = st.tabs(["⛽ Abastecimentos", "🛡️ Assunções de Serviço", "🛠️ Revisões"])
    
    with sub_tab1:
        df_filtrado_abast = obter_historico_abastecimentos(filtro_vtr)
        st.dataframe(df_filtrado_abast, use_container_width=True)
        buffer1 = io.BytesIO()
        with pd.ExcelWriter(buffer1, engine='openpyxl') as writer:
            df_filtrado_abast.to_excel(writer, index=False, sheet_name='Abastecimentos')
        st.download_button("📥 Baixar Abastecimentos (.xlsx)", data=buffer1.getvalue(), file_name="abastecimentos_10cicom.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with sub_tab2:
        df_filtrado_assuncao = obter_historico_assuncoes(filtro_vtr)
        st.dataframe(df_filtrado_assuncao, use_container_width=True)
        buffer2 = io.BytesIO()
        with pd.ExcelWriter(buffer2, engine='openpyxl') as writer:
            df_filtrado_assuncao.to_excel(writer, index=False, sheet_name='Assuncoes')
        st.download_button("📥 Baixar Assunções (.xlsx)", data=buffer2.getvalue(), file_name="assuncoes_10cicom.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with sub_tab3:
        df_filtrado_rev = obter_historico_revisoes(filtro_vtr)
        st.dataframe(df_filtrado_rev, use_container_width=True)
        buffer3 = io.BytesIO()
        with pd.ExcelWriter(buffer3, engine='openpyxl') as writer:
            df_filtrado_rev.to_excel(writer, index=False, sheet_name='Revisoes')
        st.download_button("📥 Baixar Revisões (.xlsx)", data=buffer3.getvalue(), file_name="revisoes_10cicom.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------------------------------------------------------------
# 7. QUADRO GERAL DE REVISÕES PREVENTIVAS
# -----------------------------------------------------------------------------
with tab_manut:
    st.subheader('🔍 Quadro Geral de Monitoramento das Revisões (Intervalo: 10.000 KM)')

    df_vtrs = obter_viaturas()
    df_todos = obter_historico_abastecimentos()

    lista_m = []
    for _, vtr in df_vtrs.iterrows():
        p = vtr['prefixo']
        sub = df_todos[df_todos['prefixo'] == p]
        km_atual = sub['km_atual'].iloc[0] if not sub.empty else vtr['km_revisao_base']
        base = vtr['km_revisao_base']
        intervalo = vtr['intervalo_revisao'] if vtr['intervalo_revisao'] > 0 else 10000

        prox, restante, status, _ = calcular_status_revisao(km_atual, base, intervalo)

        lista_m.append({
            'Viatura': p,
            'Modelo': vtr['modelo'],
            'KM Atual': f'{km_atual:,}'.replace(',', '.'),
            'KM Última Revisão': f'{base:,} km'.replace(',', '.') if base > 0 else 'Zero KM (Nova)',
            'Próxima Revisão': f'{prox:,}'.replace(',', '.'),
            'Faltam': f'{restante:,} km'.replace(',', '.') if restante > 0 else f'Vencida há {abs(restante):,} km'.replace(',', '.'),
            'Situação': status,
        })

    st.dataframe(pd.DataFrame(lista_m), use_container_width=True)
    st.write('---')

    st.markdown('#### ⚙️ Calibrar Parâmetros da Viatura Manualmente')
    with st.form('form_ajuste_revisao_manual'):
        col_v, col_base, col_int = st.columns(3)
        with col_v:
            vtr_escolhida = st.selectbox('Selecione a Viatura:', lista_prefixos, key="calib_vtr")
        with col_base:
            km_base_novo = st.number_input('KM Base da Última Revisão:', min_value=0, step=1000, value=0, key="calib_base")
        with col_int:
            int_novo = st.number_input('Intervalo de Manutenção (Padrão: 10.000):', min_value=1000, step=1000, value=10000, key="calib_int")

        if st.form_submit_button('💾 Gravar Novos Parâmetros', use_container_width=True):
            conn_up = get_conexao()
            cur_up = conn_up.cursor()
            cur_up.execute("""
                UPDATE viaturas 
                SET km_revisao_base = ?, intervalo_revisao = ?
                WHERE prefixo = ?
            """, (int(km_base_novo), int(int_novo), vtr_escolhida))
            conn_up.commit()
            conn_up.close()
            st.success(f'✅ Parâmetros da VTR {vtr_escolhida} calibrados com sucesso!')
            st.rerun()
            