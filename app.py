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

# Injeção de CSS para estilizar botões verdes de sucesso
st.markdown("""
    <style>
    div[data-testid="stButton"] button[kind="primary"] {
        border-radius: 8px;
    }
    .btn-verde button {
        background-color: #28a745 !important;
        color: white !important;
        border: 1px solid #28a745 !important;
        font-weight: bold !important;
    }
    </style>
""", unsafe_allow_html=True)

FUSO_MANAUS = ZoneInfo("America/Manaus")

def agora_manaus():
    return datetime.now(FUSO_MANAUS)

NOME_BANCO = "frota_10cicom.db"
ARQUIVO_EXCEL = "kms_abastecimento_10cicom.xlsx"
SENHA_PADRAO = "10cicom"

# -----------------------------------------------------------------------------
# AUTENTICAÇÃO POR SENHA (10cicom)
# -----------------------------------------------------------------------------
def verificar_login():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if not st.session_state["autenticado"]:
        col1, col2, col3 = st.columns([1, 1.2, 1])
        with col2:
            st.write("")
            st.write("")
            st.markdown("### 🚔 10ª CICOM — PMAM")
            st.markdown("#### Sistema de Gestão de Frota Operacional")
            st.info("Acesso restrito ao efetivo de serviço e administração.")
            
            with st.form("form_login"):
                senha = st.text_input("Digite a senha de acesso:", type="password", placeholder="Informe a senha...")
                btn_entrar = st.form_submit_button("🔓 Acessar Sistema", use_container_width=True)
                
                if btn_entrar:
                    if senha == SENHA_PADRAO:
                        st.session_state["autenticado"] = True
                        st.success("Acesso autorizado com sucesso!")
                        st.rerun()
                    else:
                        st.error("❌ Senha incorreta. Tente novamente.")
        st.stop()

verificar_login()

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
    
    # 2. Tabela de Abastecimentos / Odômetro Geral
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
            assuncao_id INTEGER DEFAULT NULL,
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

    # Migração defensiva para garantir campo assuncao_id
    cursor.execute("PRAGMA table_info(abastecimentos)")
    colunas_abast = [c[1] for c in cursor.fetchall()]
    if "assuncao_id" not in colunas_abast:
        cursor.execute("ALTER TABLE abastecimentos ADD COLUMN assuncao_id INTEGER DEFAULT NULL")
    
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
                
    km_match = re.search(r'(?:KM|ODOMETRO|QUILOMETRAGEM)[:\s]*([0-9\.\,]+)', texto, re.IGNORECASE)
    if km_match:
        km_clean = km_match.group(1).replace('.', '').replace(',', '')
        if km_clean.isdigit():
            dados["km"] = int(km_clean)
            
    l_match = re.search(r'(?:LITROS|LT|LTS|QTDE|QUANTIDADE)[:\s]*([0-9\.\,]+)', texto, re.IGNORECASE)
    if l_match:
        try:
            dados["litros"] = float(l_match.group(1).replace(',', '.'))
        except:
            pass
            
    mot_match = re.search(r'(?:MOTORISTA|CONDUTOR|POLICIAL|FAROL)[:\s]*([^\n\r]+)', texto, re.IGNORECASE)
    if mot_match:
        dados["motorista"] = mot_match.group(1).strip()
        
    cmt_match = re.search(r'(?:CMT|COMANDANTE|CHEFE|ENCARREGADO)[:\s]*([^\n\r]+)', texto, re.IGNORECASE)
    if cmt_match:
        dados["cmt_guarnicao"] = cmt_match.group(1).strip()

    h_match = re.search(r'(?:HORA|HORARIO|HORÁRIO)[:\s]*([0-9]{1,2}[:hH][0-9]{2})', texto, re.IGNORECASE)
    if h_match:
        h = h_match.group(1).replace('h', ':').replace('H', ':')
        if len(h) == 4 and h[1] == ':':
            h = '0' + h
        dados["horario"] = h[:5]
        
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
        q = "SELECT id, prefixo, data, horario, km_atual, litros, motorista, observacao, origem, assuncao_id FROM abastecimentos WHERE prefixo = ? ORDER BY data DESC, km_atual DESC, id DESC"
        df = pd.read_sql_query(q, conn, params=(prefixo,))
    else:
        q = "SELECT id, prefixo, data, horario, km_atual, litros, motorista, observacao, origem, assuncao_id FROM abastecimentos ORDER BY data DESC, km_atual DESC, id DESC"
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

df_vtrs = obter_viaturas()
lista_prefixos = df_vtrs['prefixo'].tolist()

# -----------------------------------------------------------------------------
# BARRA LATERAL (MENU OPERACIONAL COM BOTÕES NA LATERAL ESQUERDA)
# -----------------------------------------------------------------------------
if "tela_ativa" not in st.session_state:
    st.session_state["tela_ativa"] = "📊 Visão Geral & Odômetros"

with st.sidebar:
    st.markdown("### 🚔 10ª CICOM — PMAM")
    st.caption(f"Fuso Oficial: UTC-4 (Manaus)\n{agora_manaus().strftime('%d/%m/%Y %H:%M')}")
    st.write("---")

    filtro_vtr = st.selectbox("Filtrar Viatura:", ["Todas"] + lista_prefixos)
    st.write("---")
    st.markdown("#### 🔘 Módulos de Operação")

    telas = [
        ("📊 Visão Geral & Odômetros", "btn_menu_dash"),
        ("🛡️ Assunção de Serviço", "btn_menu_assuncao"),
        ("⛽ Registrar Abastecimento", "btn_menu_abast"),
        ("✏️ Corrigir / Editar Lançamentos", "btn_menu_corrigir"),
        ("🛠️ Registrar Revisão Realizada", "btn_menu_rev"),
        ("⚠️ Avarias & Ocorrências", "btn_menu_avarias"),
        ("📑 Histórico Completo", "btn_menu_hist"),
        ("🔍 Monitor de Revisões", "btn_menu_monit")
    ]

    for label, chave in telas:
        estilo = "primary" if st.session_state["tela_ativa"] == label else "secondary"
        if st.button(label, key=chave, use_container_width=True, type=estilo):
            st.session_state["tela_ativa"] = label
            # Limpa estados de gravação anterior ao alternar telas
            st.session_state.pop("gravado_assuncao", None)
            st.session_state.pop("gravado_abastecimento", None)
            st.rerun()

    st.write("---")
    if st.button("🔒 Sair do Sistema (Logout)", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()

# Cabeçalho Principal
st.title("🚔 Frota Operacional — 10ª CICOM")
st.caption(f"Módulo Ativo: **{st.session_state['tela_ativa']}** | Horário Manaus: {agora_manaus().strftime('%d/%m/%Y %H:%M:%S')}")
st.write("---")

# =============================================================================
# TELA 1: DASHBOARD GERAL
# =============================================================================
if st.session_state["tela_ativa"] == "📊 Visão Geral & Odômetros":
    df_todos_abast = obter_historico_abastecimentos()
    st.subheader("Odômetro Mais Recente da Frota & Diagnóstico de Revisão")
    cols = st.columns(len(df_vtrs))
    
    for idx, vtr in df_vtrs.iterrows():
        p = vtr['prefixo']
        sub = df_todos_abast[df_todos_abast['prefixo'] == p]
        km_ultimo = sub['km_atual'].iloc[0] if not sub.empty else vtr['km_revisao_base']
        litros_total = sub[sub['litros'] > 0]['litros'].sum() if not sub.empty else 0
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
        st.dataframe(obter_historico_assuncoes().head(6), use_container_width=True)
    with col_dash2:
        st.markdown("##### ⛽ Últimos Abastecimentos")
        st.dataframe(df_todos_abast[df_todos_abast['litros'] > 0].head(6), use_container_width=True)

# =============================================================================
# TELA 2: ASSUNÇÃO DE SERVIÇO COM BOTÃO VERDE DE SUCESSO
# =============================================================================
elif st.session_state["tela_ativa"] == "🛡️ Assunção de Serviço":
    st.subheader("🛡️ Registro de Assunção de Serviço da Viatura")
    st.caption("Cadastre a entrada de serviço. O sistema calcula na hora se a viatura está dentro do limite de revisão.")

    # Exibe badge de sucesso e botão verde se acabou de ser gravado
    ja_gravou_ass = st.session_state.get("gravado_assuncao", False)
    if ja_gravou_ass:
        st.success(f"🎉 **Assunção de Serviço Homologada com Sucesso!** VTR {st.session_state.get('ultima_vtr_ass')} registrada.")
        if st.button("➕ Realizar Outra Assunção", key="btn_novo_ass"):
            st.session_state["gravado_assuncao"] = False
            st.rerun()

    metodo_assuncao = st.radio(
        "Método de Entrada:",
        ["✍️ Manual", "📷 Leitura de Foto do Painel", "💬 Texto do WhatsApp"],
        horizontal=True
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
            "Cole a mensagem de assunção do WhatsApp:",
            height=110,
            placeholder="Exemplo:\n*ASSUNÇÃO DE SERVIÇO*\nTurno: 1º Turno (07h às 19h)\nVTR: 1001\nKM: 44.200\nCmt: SGT PM CARDOSO\nMotorista: CB PM SILVA\nObs: Sem alterações"
        )
        if st.button("🔍 Processar Mensagem WhatsApp", key="btn_p_ass"):
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
            foto_ass = st.file_uploader("Upload da foto do painel/odômetro:", type=['jpg', 'jpeg', 'png'], key="f_up_ass")
            foto_cam_ass = st.camera_input("Fotografar odômetro:", key="f_cam_ass")
            foto_u_ass = foto_cam_ass or foto_ass
            if foto_u_ass:
                km_det = extrair_km_imagem_ocr(foto_u_ass.getvalue())
                if km_det > 0:
                    dados_assuncao['km'] = km_det
                    st.success(f"Odômetro detectado: {km_det:,} KM".replace(',', '.'))
                dados_assuncao['origem'] = "FOTO_PAINEL"
        with c_p2:
            if foto_u_ass:
                st.image(foto_u_ass, caption="Foto do Painel anexada", use_container_width=True)

    st.write("---")
    st.markdown("##### 📝 Formulário de Confirmação da Assunção de Serviço")
    c1, c2, c3 = st.columns(3)
    with c1:
        idx_vtr_a = lista_prefixos.index(dados_assuncao['prefixo']) if dados_assuncao['prefixo'] in lista_prefixos else 0
        vtr_sel = st.selectbox("Viatura:", lista_prefixos, index=idx_vtr_a, key="as_vtr")
        turno_sel = st.selectbox("Turno de Serviço:", ["1º Turno (07h às 19h)", "2º Turno (19h às 07h)", "Turno Extra / Operação Especial"], key="as_trn")
    with c2:
        dt_sel = st.date_input("Data:", value=dados_assuncao['data'], key="as_dt")
        hr_sel = st.text_input("Horário (Manaus):", value=dados_assuncao['horario'], key="as_hr")
    with c3:
        km_sel = st.number_input("KM Inicial (Odômetro Atual):", value=int(dados_assuncao['km']), step=1, key="as_km")
        cmt_sel = st.text_input("Comandante da Guarnição:", value=dados_assuncao['cmt_guarnicao'], placeholder="Ex: SGT PM CARDOSO", key="as_cmt")

    mot_sel = st.text_input("Motorista / Condutor:", value=dados_assuncao['motorista'], placeholder="Ex: CB PM SILVA", key="as_mot")
    obs_sel = st.text_area("Observações da Viatura / Avarias na Entrada:", value=dados_assuncao['observacao'], placeholder="Ex: Nível de óleo e água checados; sem alterações.", key="as_obs")

    if km_sel > 0:
        info_vtr = df_vtrs[df_vtrs['prefixo'] == vtr_sel].iloc[0]
        prox_rev, faltam_km, sit_rev, _ = calcular_status_revisao(km_sel, info_vtr['km_revisao_base'], info_vtr['intervalo_revisao'])

        st.markdown("#### 🔍 Diagnóstico Automático de Revisão:")
        cd1, cd2, cd3 = st.columns(3)
        cd1.metric("KM Informado pelo Motorista", f"{km_sel:,} km".replace(",", "."))
        cd2.metric("Próxima Revisão Prevista", f"{prox_rev:,} km".replace(",", "."))
        if faltam_km <= 0:
            cd3.metric("Situação", sit_rev, delta=f"Vencida há {abs(faltam_km):,} km".replace(",", "."), delta_color="inverse")
            st.error(f"🚨 **URGENTE:** A VTR {vtr_sel} excedeu o limite de revisão preventiva em **{abs(faltam_km):,} KM**.")
        elif faltam_km <= 1000:
            cd3.metric("Faltam para Revisão", f"{faltam_km:,} km".replace(",", "."), delta="Atenção: Menos de 1.000 km!", delta_color="inverse")
            st.warning(f"⚠️ **ALERTA:** Faltam apenas **{faltam_km:,} KM** para a próxima revisão da VTR {vtr_sel}.")
        else:
            cd3.metric("Faltam para Revisão", f"{faltam_km:,} km".replace(",", "."), delta="Em dia")
            st.success(f"✅ **REGULAR:** A VTR {vtr_sel} tem **{faltam_km:,} KM** restantes até a próxima revisão.")

    # Renderização do Botão de Confirmação com cor Verde quando registrado
    if ja_gravou_ass:
        st.markdown('<div class="btn-verde">', unsafe_allow_html=True)
        st.button("✅ Assunção Registrada com Sucesso!", use_container_width=True, disabled=True, key="btn_confirmado_ass")
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        if st.button("💾 Confirmar e Homologar Assunção de Serviço", type="primary", use_container_width=True, key="btn_salvar_ass"):
            conn = get_conexao()
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO assuncoes_servico (prefixo, data, horario, turno, cmt_guarnicao, motorista, km_inicial, observacao, origem)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (vtr_sel, dt_sel.strftime('%Y-%m-%d'), hr_sel, turno_sel, cmt_sel, mot_sel, int(km_sel), obs_sel, dados_assuncao['origem']))
            
            id_nova_assuncao = cur.lastrowid
            
            # Sincroniza odômetro no histórico geral vinculando o assuncao_id
            cur.execute("""
                INSERT INTO abastecimentos (prefixo, data, horario, km_atual, litros, motorista, observacao, origem, assuncao_id)
                VALUES (?, ?, ?, ?, 0.0, ?, ?, ?, ?)
            """, (vtr_sel, dt_sel.strftime('%Y-%m-%d'), hr_sel, int(km_sel), mot_sel, f"Assunção ({turno_sel}) - {obs_sel}".strip(), dados_assuncao['origem'], id_nova_assuncao))

            if obs_sel.strip() and obs_sel.lower() not in ['sem alteracao', 'sem alteração', 'ok', 'tudo ok']:
                cur.execute("""
                    INSERT INTO ocorrencias (prefixo, data, tipo, descricao, status, registrado_por)
                    VALUES (?, ?, 'Avaria / Assunção', ?, 'Pendente', ?)
                """, (vtr_sel, dt_sel.strftime('%Y-%m-%d'), obs_sel, mot_sel or cmt_sel))

            conn.commit()
            conn.close()
            st.session_state["gravado_assuncao"] = True
            st.session_state["ultima_vtr_ass"] = vtr_sel
            if 'assuncao_ext' in st.session_state:
                del st.session_state['assuncao_ext']
            st.rerun()

# =============================================================================
# TELA 3: REGISTRO DE ABASTECIMENTO COM BOTÃO VERDE DE SUCESSO
# =============================================================================
elif st.session_state["tela_ativa"] == "⛽ Registrar Abastecimento":
    st.subheader("⛽ Registro de Abastecimento de Viaturas")
    st.caption("Cadastre os abastecimentos efetuados via Manual, Leitura de Foto da Bomba/Painel ou Mensagem do WhatsApp.")

    ja_gravou_abast = st.session_state.get("gravado_abastecimento", False)
    if ja_gravou_abast:
        st.success(f"🎉 **Abastecimento Registrado com Sucesso!** VTR {st.session_state.get('ultima_vtr_abast')} homologada.")
        if st.button("➕ Realizar Novo Abastecimento", key="btn_novo_abast"):
            st.session_state["gravado_abastecimento"] = False
            st.rerun()

    metodo_abast = st.radio(
        "Método de Entrada:",
        ["✍️ Manual", "📷 Leitura de Foto da Bomba / Painel", "💬 Texto do WhatsApp"],
        horizontal=True
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
            height=100,
            placeholder="Exemplo:\nVTR: 1001\nKM: 42.150\nLitros: 55L\nMotorista: CB PM SOUZA\nObs: Posto Equador"
        )
        if st.button("🔍 Processar Mensagem Abastecimento", key="btn_p_ab"):
            if txt_abast.strip():
                st.session_state['abast_ext'] = extrair_dados_whatsapp(txt_abast)
                st.success("Dados identificados!")

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
            foto_ab = st.file_uploader("Upload da foto da bomba ou odômetro:", type=['jpg', 'jpeg', 'png'], key="f_up_ab")
            foto_cam_ab = st.camera_input("Fotografar:", key="f_cam_ab")
            foto_u_ab = foto_cam_ab or foto_ab
            if foto_u_ab:
                km_det_ab = extrair_km_imagem_ocr(foto_u_ab.getvalue())
                if km_det_ab > 0:
                    dados_abast['km'] = km_det_ab
                    st.success(f"Odômetro detectado: {km_det_ab:,} KM".replace(',', '.'))
                dados_abast['origem'] = "FOTO_PAINEL"
        with col_f2:
            if foto_u_ab:
                st.image(foto_u_ab, caption="Comprovante / Foto anexada", use_container_width=True)

    st.markdown("##### 📝 Formulário de Confirmação de Abastecimento")
    c1, c2, c3 = st.columns(3)
    with c1:
        idx_vtr_b = lista_prefixos.index(dados_abast['prefixo']) if dados_abast['prefixo'] in lista_prefixos else 0
        vtr_b = st.selectbox("Viatura:", lista_prefixos, index=idx_vtr_b, key="ab_vtr")
        data_b = st.date_input("Data:", value=dados_abast['data'], key="ab_dt")
    with c2:
        km_b = st.number_input("Odômetro (KM):", value=int(dados_abast['km']), step=1, key="ab_km")
        litros_b = st.number_input("Litros:", value=float(dados_abast['litros']), step=0.1, format="%.2f", key="ab_lt")
    with c3:
        mot_b = st.text_input("Motorista / Condutor:", value=dados_abast['motorista'], placeholder="Ex: CB PM ALMEIDA", key="ab_mot")
        hr_b = st.text_input("Horário (Manaus):", value=dados_abast['horario'], key="ab_hr")

    obs_b = st.text_input("Observação / Posto / Combustível:", value=dados_abast['observacao'], key="ab_obs")

    # Botão com estilização verde quando gravado
    if ja_gravou_abast:
        st.markdown('<div class="btn-verde">', unsafe_allow_html=True)
        st.button("✅ Abastecimento Registrado com Sucesso!", use_container_width=True, disabled=True, key="btn_confirmado_abast")
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        if st.button("💾 Salvar Registro de Abastecimento", type="primary", use_container_width=True, key="btn_salvar_abast_exec"):
            conn = get_conexao()
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO abastecimentos (prefixo, data, horario, km_atual, litros, motorista, observacao, origem)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (vtr_b, data_b.strftime('%Y-%m-%d'), hr_b, int(km_b), float(litros_b), mot_b, obs_b, dados_abast['origem']))
            
            if any(p in obs_b.lower() for p in ['avaria', 'quebrado', 'pneu', 'furado', 'luz', 'farol', 'defeito']):
                cur.execute("""
                    INSERT INTO ocorrencias (prefixo, data, tipo, descricao, status, registrado_por)
                    VALUES (?, ?, 'Avaria / Abastecimento', ?, 'Pendente', ?)
                """, (vtr_b, data_b.strftime('%Y-%m-%d'), obs_b, mot_b))

            conn.commit()
            conn.close()
            st.session_state["gravado_abastecimento"] = True
            st.session_state["ultima_vtr_abast"] = vtr_b
            if 'abast_ext' in st.session_state:
                del st.session_state['abast_ext']
            st.rerun()

# =============================================================================
# TELA 4: CORREÇÃO E EXCLUSÃO DE LANÇAMENTOS (LIMPA TOTALMENTE DO HISTÓRICO)
# =============================================================================
elif st.session_state["tela_ativa"] == "✏️ Corrigir / Editar Lançamentos":
    st.subheader("✏️ Módulo de Correção e Exclusão de Registros")
    st.caption("Ao corrigir ou excluir um registro, ele é imediatamente atualizado ou removido do histórico oficial.")

    tipo_correcao = st.radio(
        "Qual lançamento deseja corrigir ou excluir?",
        ["⛽ Abastecimento", "🛡️ Assunção de Serviço"],
        horizontal=True
    )

    conn = get_conexao()
    
    if tipo_correcao == "⛽ Abastecimento":
        df_ab = pd.read_sql_query("SELECT id, prefixo, data, horario, km_atual, litros, motorista, observacao FROM abastecimentos WHERE litros > 0 OR origem != 'MANUAL' ORDER BY id DESC LIMIT 50", conn)
        
        if df_ab.empty:
            st.info("Nenhum lançamento de abastecimento localizado para edição.")
        else:
            opcoes_ab = {
                f"ID #{row['id']} | VTR {row['prefixo']} | Data {row['data']} {row['horario']} | {row['km_atual']} KM | {row['litros']}L | {row['motorista']}": row['id']
                for _, row in df_ab.iterrows()
            }
            sel_ab_texto = st.selectbox("Selecione o abastecimento a corrigir:", list(opcoes_ab.keys()))
            id_sel_ab = opcoes_ab[sel_ab_texto]
            
            registro_ab = df_ab[df_ab['id'] == id_sel_ab].iloc[0]

            with st.form("form_edita_abast"):
                st.markdown(f"#### Editando Abastecimento (ID #{id_sel_ab})")
                col_e1, col_e2, col_e3 = st.columns(3)
                with col_e1:
                    novo_vtr_ab = st.selectbox("Viatura:", lista_prefixos, index=lista_prefixos.index(registro_ab['prefixo']) if registro_ab['prefixo'] in lista_prefixos else 0)
                    nova_data_ab = st.date_input("Data:", value=pd.to_datetime(registro_ab['data']).date())
                with col_e2:
                    novo_km_ab = st.number_input("KM Atual:", value=int(registro_ab['km_atual']), step=1)
                    novos_litros_ab = st.number_input("Litros:", value=float(registro_ab['litros']), step=0.1, format="%.2f")
                with col_e3:
                    novo_mot_ab = st.text_input("Motorista:", value=registro_ab['motorista'] or "")
                    novo_hr_ab = st.text_input("Horário:", value=registro_ab['horario'] or "")

                nova_obs_ab = st.text_input("Observação:", value=registro_ab['observacao'] or "")

                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    salvar_edicao_ab = st.form_submit_button("💾 Salvar Correções no Registro", type="primary", use_container_width=True)
                with col_btn2:
                    excluir_ab = st.form_submit_button("🗑️ Excluir Definitivamente do Histórico", use_container_width=True)

                if salvar_edicao_ab:
                    cur = conn.cursor()
                    cur.execute("""
                        UPDATE abastecimentos 
                        SET prefixo = ?, data = ?, horario = ?, km_atual = ?, litros = ?, motorista = ?, observacao = ?
                        WHERE id = ?
                    """, (novo_vtr_ab, nova_data_ab.strftime('%Y-%m-%d'), novo_hr_ab, int(novo_km_ab), float(novos_litros_ab), novo_mot_ab, nova_obs_ab, id_sel_ab))
                    conn.commit()
                    st.success(f"Registro #{id_sel_ab} corrigido com sucesso!")
                    st.rerun()

                if excluir_ab:
                    cur = conn.cursor()
                    cur.execute("DELETE FROM abastecimentos WHERE id = ?", (id_sel_ab,))
                    conn.commit()
                    st.warning(f"Registro #{id_sel_ab} excluído e removido do histórico!")
                    st.rerun()

    else:
        df_as = pd.read_sql_query("SELECT id, prefixo, data, horario, turno, cmt_guarnicao, motorista, km_inicial, observacao FROM assuncoes_servico ORDER BY id DESC LIMIT 50", conn)
        
        if df_as.empty:
            st.info("Nenhuma assunção de serviço localizada para edição.")
        else:
            opcoes_as = {
                f"ID #{row['id']} | VTR {row['prefixo']} | Data {row['data']} {row['horario']} | {row['turno']} | {row['km_inicial']} KM | Cmt: {row['cmt_guarnicao']}": row['id']
                for _, row in df_as.iterrows()
            }
            sel_as_texto = st.selectbox("Selecione a assunção de serviço a corrigir:", list(opcoes_as.keys()))
            id_sel_as = opcoes_as[sel_as_texto]
            
            registro_as = df_as[df_as['id'] == id_sel_as].iloc[0]

            with st.form("form_edita_assuncao"):
                st.markdown(f"#### Editando Assunção de Serviço (ID #{id_sel_as})")
                col_s1, col_s2, col_s3 = st.columns(3)
                with col_s1:
                    novo_vtr_as = st.selectbox("Viatura:", lista_prefixos, index=lista_prefixos.index(registro_as['prefixo']) if registro_as['prefixo'] in lista_prefixos else 0)
                    nova_data_as = st.date_input("Data:", value=pd.to_datetime(registro_as['data']).date())
                with col_s2:
                    novo_km_as = st.number_input("KM Inicial:", value=int(registro_as['km_inicial']), step=1)
                    turnos_disp = ["1º Turno (07h às 19h)", "2º Turno (19h às 07h)", "Turno Extra / Operação Especial"]
                    idx_trn = turnos_disp.index(registro_as['turno']) if registro_as['turno'] in turnos_disp else 0
                    novo_turno_as = st.selectbox("Turno:", turnos_disp, index=idx_trn)
                with col_s3:
                    novo_cmt_as = st.text_input("Comandante Guarnição:", value=registro_as['cmt_guarnicao'] or "")
                    novo_hr_as = st.text_input("Horário:", value=registro_as['horario'] or "")

                novo_mot_as = st.text_input("Motorista:", value=registro_as['motorista'] or "")
                nova_obs_as = st.text_area("Observações / Avarias na Assunção:", value=registro_as['observacao'] or "")

                col_btn_s1, col_btn_s2 = st.columns(2)
                with col_btn_s1:
                    salvar_edicao_as = st.form_submit_button("💾 Salvar Correções na Assunção", type="primary", use_container_width=True)
                with col_btn_s2:
                    excluir_as = st.form_submit_button("🗑️ Excluir Definitivamente do Histórico", use_container_width=True)

                if salvar_edicao_as:
                    cur = conn.cursor()
                    # Atualiza tabela de assunções
                    cur.execute("""
                        UPDATE assuncoes_servico 
                        SET prefixo = ?, data = ?, horario = ?, turno = ?, cmt_guarnicao = ?, motorista = ?, km_inicial = ?, observacao = ?
                        WHERE id = ?
                    """, (novo_vtr_as, nova_data_as.strftime('%Y-%m-%d'), novo_hr_as, novo_turno_as, novo_cmt_as, novo_mot_as, int(novo_km_as), nova_obs_as, id_sel_as))
                    
                    # Sincroniza o espelho do odômetro no histórico geral de abastecimentos
                    cur.execute("""
                        UPDATE abastecimentos
                        SET prefixo = ?, data = ?, horario = ?, km_atual = ?, motorista = ?, observacao = ?
                        WHERE assuncao_id = ?
                    """, (novo_vtr_as, nova_data_as.strftime('%Y-%m-%d'), novo_hr_as, int(novo_km_as), novo_mot_as, f"Assunção ({novo_turno_as}) - {nova_obs_as}".strip(), id_sel_as))
                    
                    conn.commit()
                    st.success(f"Assunção #{id_sel_as} e histórico sincronizados com sucesso!")
                    st.rerun()

                if excluir_as:
                    cur = conn.cursor()
                    # Remove da tabela de assunção
                    cur.execute("DELETE FROM assuncoes_servico WHERE id = ?", (id_sel_as,))
                    # Remove o espelho da assunção da tabela de abastecimentos/histórico
                    cur.execute("DELETE FROM abastecimentos WHERE assuncao_id = ?", (id_sel_as,))
                    conn.commit()
                    st.warning(f"Assunção #{id_sel_as} removida completamente do histórico!")
                    st.rerun()

    conn.close()

# =============================================================================
# TELA 5: REGISTRO DE REVISÃO REALIZADA
# =============================================================================
elif st.session_state["tela_ativa"] == "🛠️ Registrar Revisão Realizada":
    st.subheader("🛠️ Registrar Revisão Mecânica Realizada")
    st.caption("Use esta área para homologar a revisão preventiva feita na concessionária ou oficina conveniada. A base da viatura será atualizada automaticamente.")

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
            rev_km = st.number_input("Odômetro da Viatura na Revisão (KM):", min_value=0, step=1000)
        with c_r3:
            rev_oficina = st.text_input("Oficina / Concessionária:", placeholder="Ex: Concessionária Chevrolet / Oficina PMAM")
            rev_resp = st.text_input("Policial Responsável pelo Recebimento:", placeholder="Ex: SGT PM MOURA")

        rev_obs = st.text_area("Serviços Executados e Peças Trocadas (OS):", placeholder="Ex: Troca de óleo, pastilhas de freio e filtros.")

        if st.form_submit_button("✅ Homologar e Gravar Revisão", type="primary", use_container_width=True):
            if rev_km > 0:
                conn = get_conexao()
                cur = conn.cursor()
                cur.execute("""
                    INSERT INTO historico_revisoes (prefixo, data, km_revisao, tipo_revisao, oficina, responsavel, observacoes)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (rev_vtr, rev_data.strftime('%Y-%m-%d'), int(rev_km), rev_tipo, rev_oficina, rev_resp, rev_obs))
                
                cur.execute("""
                    UPDATE viaturas 
                    SET km_revisao_base = ? 
                    WHERE prefixo = ?
                """, (int(rev_km), rev_vtr))
                
                conn.commit()
                conn.close()
                st.success(f"🎉 Revisão da VTR {rev_vtr} homologada com sucesso! Base atualizada para {rev_km:,} KM.")
                st.rerun()
            else:
                st.warning("Informe o odômetro exato da revisão.")

    st.write("---")
    st.markdown("##### 📋 Histórico de Revisões Cadastradas")
    df_rev_historico = obter_historico_revisoes(filtro_vtr)
    if df_rev_historico.empty:
        st.info("Nenhuma revisão homologada cadastrada.")
    else:
        st.dataframe(df_rev_historico, use_container_width=True)

# =============================================================================
# TELA 6: AVARIAS & OCORRÊNCIAS
# =============================================================================
elif st.session_state["tela_ativa"] == "⚠️ Avarias & Ocorrências":
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
            av_desc = st.text_area("Descrição do Problema:", placeholder="Ex: Pneu dianteiro esquerdo furado.")
            av_policial = st.text_input("Policial Relatante:", placeholder="Ex: SGT PM CARDOSO")
            
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
                    st.success(f"Avaria da VTR {av_vtr} registrada!")
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
                            st.success("Ocorrência marcada como solucionada!")
                            st.rerun()

# =============================================================================
# TELA 7: HISTÓRICO COMPLETO
# =============================================================================
elif st.session_state["tela_ativa"] == "📑 Histórico Completo":
    st.subheader(f"Histórico Geral da Frota ({filtro_vtr})")
    
    sub1, sub2, sub3 = st.tabs(["⛽ Abastecimentos", "🛡️ Assunções de Serviço", "🛠️ Revisões"])
    
    with sub1:
        df_f_ab = obter_historico_abastecimentos(filtro_vtr)
        st.dataframe(df_f_ab.drop(columns=['assuncao_id'], errors='ignore'), use_container_width=True)
        buf1 = io.BytesIO()
        with pd.ExcelWriter(buf1, engine='openpyxl') as writer:
            df_f_ab.to_excel(writer, index=False, sheet_name='Abastecimentos')
        st.download_button("📥 Baixar Abastecimentos (.xlsx)", data=buf1.getvalue(), file_name="abastecimentos_10cicom.xlsx")

    with sub2:
        df_f_as = obter_historico_assuncoes(filtro_vtr)
        st.dataframe(df_f_as, use_container_width=True)
        buf2 = io.BytesIO()
        with pd.ExcelWriter(buf2, engine='openpyxl') as writer:
            df_f_as.to_excel(writer, index=False, sheet_name='Assuncoes')
        st.download_button("📥 Baixar Assunções (.xlsx)", data=buf2.getvalue(), file_name="assuncoes_10cicom.xlsx")

    with sub3:
        df_f_rv = obter_historico_revisoes(filtro_vtr)
        st.dataframe(df_f_rv, use_container_width=True)
        buf3 = io.BytesIO()
        with pd.ExcelWriter(buf3, engine='openpyxl') as writer:
            df_f_rv.to_excel(writer, index=False, sheet_name='Revisoes')
        st.download_button("📥 Baixar Revisões (.xlsx)", data=buf3.getvalue(), file_name="revisoes_10cicom.xlsx")

# =============================================================================
# TELA 8: QUADRO GERAL DE REVISÕES
# =============================================================================
elif st.session_state["tela_ativa"] == "🔍 Monitor de Revisões":
    st.subheader('🔍 Quadro Geral de Monitoramento das Revisões (Intervalo: 10.000 KM)')

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
            vtr_esc = st.selectbox('Selecione a Viatura:', lista_prefixos, key="cal_vtr")
        with col_base:
            km_b_nov = st.number_input('KM Base da Última Revisão:', min_value=0, step=1000, value=0, key="cal_b")
        with col_int:
            int_nov = st.number_input('Intervalo de Manutenção (Padrão: 10.000):', min_value=1000, step=1000, value=10000, key="cal_i")

        if st.form_submit_button('💾 Gravar Novos Parâmetros', use_container_width=True):
            conn_up = get_conexao()
            cur_up = conn_up.cursor()
            cur_up.execute("""
                UPDATE viaturas 
                SET km_revisao_base = ?, intervalo_revisao = ?
                WHERE prefixo = ?
            """, (int(km_b_nov), int(int_nov), vtr_esc))
            conn_up.commit()
            conn_up.close()
            st.success(f'✅ Parâmetros da VTR {vtr_esc} calibrados com sucesso!')
            st.rerun()
            