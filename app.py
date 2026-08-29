import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import sqlite3
import datetime
import docx
import os
import io
import re
from PIL import Image

# Tenta carregar OCR (easyocr ou pytesseract se disponível)
try:
    import easyocr
    OCR_ENGINE = "easyocr"
except ImportError:
    try:
        import pytesseract
        OCR_ENGINE = "pytesseract"
    except ImportError:
        OCR_ENGINE = "manual"

# Configuração da Página
st.set_page_config(page_title="Gestão de Frota - 10ª CICOM", layout="wide", page_icon="🚔")

DB_FILE = "frota_10cicom.db"
FOTOS_DIR = "fotos_paineis"
os.makedirs(FOTOS_DIR, exist_ok=True)

# --- CONTROLE DE ACESSO (LOGIN) ---
SENHA_PADRAO = "10cicom"

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

# --- EXTRAÇÃO DE TEXTO DO PRINT / FOTO (OCR + REGEX) ---
@st.cache_resource
def carregar_leitor_ocr():
    if OCR_ENGINE == "easyocr":
        return easyocr.Reader(['pt', 'en'], gpu=False)
    return None

def extrair_texto_imagem(imagem_bytes):
    texto_completo = ""
    if OCR_ENGINE == "easyocr":
        reader = carregar_leitor_ocr()
        resultados = reader.readtext(imagem_bytes.getvalue(), detail=0)
        texto_completo = " ".join(resultados)
    elif OCR_ENGINE == "pytesseract":
        img = Image.open(imagem_bytes)
        texto_completo = pytesseract.image_to_string(img, lang='por')
    return texto_completo

def interpretar_dados_mensagem(texto):
    dados = {
        "prefixo": "25-1001",
        "motorista": "",
        "km_atual": 0.0,
        "autonomia": 0.0,
        "tipo_operacao": "⛽ Realizar Abastecimento",
        "status_registro": "Abastecido"
    }
    
    if not texto:
        return dados
        
    texto_clean = texto.upper()
    
    # 1. Identifica Prefixo da Viatura
    if "1001" in texto_clean or "25-1001" in texto_clean:
        dados["prefixo"] = "25-1001"
    elif "1111" in texto_clean or "25-1111" in texto_clean:
        dados["prefixo"] = "25-1111"
    elif "1329" in texto_clean or "25-1329" in texto_clean:
        dados["prefixo"] = "25-1329"
    elif "1353" in texto_clean or "25-1353" in texto_clean:
        dados["prefixo"] = "25-1353"
    elif "1394" in texto_clean or "25-1394" in texto_clean:
        dados["prefixo"] = "25-1394"

    # 2. Identifica KM / Odômetro
    match_km = re.search(r'(?:KM|ODO|ODÔMETRO|INICIAL)[\s:\.\-]*([0-9]{2,6}(?:[\.,][0-9]{1,3})?)', texto_clean)
    if match_km:
        num_str = match_km.group(1).replace('.', '').replace(',', '.')
        try:
            dados["km_atual"] = float(num_str)
        except ValueError:
            pass
    else:
        match_numeros = re.findall(r'\b[0-9]{4,6}\b', texto_clean)
        if match_numeros:
            for num in match_numeros:
                if num not in ["1001", "1111", "1329", "1353", "1394"]:
                    dados["km_atual"] = float(num)
                    break

    # 3. Identifica Motorista / Farol
    match_mot = re.search(r'(?:MOTORISTA|CONDUTOR|FAROL|SD|CB|SGT|SUB|TEN|CAP)[\s:\.\-]*([A-Z\s]+?)(?=\n|KM|VTR|AUTO|$)', texto_clean)
    if match_mot:
        dados["motorista"] = match_mot.group(0).strip().title()

    # 4. Identifica Autonomia
    match_auto = re.search(r'(?:AUTONOMIA|AUTO)[\s:\.\-]*([0-9]{1,4})', texto_clean)
    if match_auto:
        try:
            dados["autonomia"] = float(match_auto.group(1))
        except ValueError:
            pass
            
    # 5. Classifica o Tipo de Operação
    if "ASSUNÇÃO" in texto_clean or "ENTRADA" in texto_clean or "INÍCIO" in texto_clean:
        dados["tipo_operacao"] = "🕒 Assunção de Serviço (Entrada de Turno / KM Inicial)"
        dados["status_registro"] = "Assunção de Serviço"
    elif "NÃO ABASTECEU" in texto_clean or "N/A" in texto_clean:
        dados["tipo_operacao"] = "🚫 Informar que a Viatura Não Abasteceu Hoje (N/A)"
        dados["status_registro"] = "Não Abasteceu (N/A)"

    return dados

# --- COMPONENTE DE DITADO POR VOZ ---
def componente_ditado_voz(chave_id="transcricao_box"):
    html_code = f"""
    <div style="background-color: #f0f2f6; border-radius: 8px; padding: 12px; margin-bottom: 10px; font-family: sans-serif;">
        <button id="btn_rec_{chave_id}" type="button" style="background-color: #ff4b4b; color: white; border: none; padding: 8px 16px; border-radius: 5px; font-weight: bold; cursor: pointer;">
            🎙️ Iniciar Ditado por Voz
        </button>
        <span id="status_{chave_id}" style="margin-left: 10px; font-size: 13px; color: #555;">Clique para falar...</span>
        <div style="margin-top: 8px;">
            <textarea id="{chave_id}" rows="3" style="width: 100%; border: 1px solid #ccc; border-radius: 5px; padding: 6px; font-size: 14px;" placeholder="O texto falado aparecerá aqui..."></textarea>
        </div>
        <button id="btn_copiar_{chave_id}" type="button" style="margin-top: 5px; background-color: #0e1117; color: white; border: none; padding: 5px 10px; border-radius: 4px; font-size: 12px; cursor: pointer;">
            📋 Copiar Texto Ditado
        </button>
    </div>

    <script>
        const btnRec = document.getElementById('btn_rec_{chave_id}');
        const statusTxt = document.getElementById('status_{chave_id}');
        const txtArea = document.getElementById('{chave_id}');
        const btnCopiar = document.getElementById('btn_copiar_{chave_id}');

        const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
        if (SpeechRecognition) {{
            const recognition = new SpeechRecognition();
            recognition.lang = 'pt-BR';
            recognition.continuous = false;
            recognition.interimResults = false;

            btnRec.onclick = () => {{
                try {{
                    recognition.start();
                    statusTxt.innerText = '🔴 Ouvindo... Pode falar.';
                    btnRec.style.backgroundColor = '#d32f2f';
                }} catch (e) {{
                    recognition.stop();
                }}
            }};

            recognition.onresult = (event) => {{
                const transcript = event.results[0][0].transcript;
                if (txtArea.value) {{
                    txtArea.value += ' ' + transcript;
                }} else {{
                    txtArea.value = transcript;
                }}
                statusTxt.innerText = '✅ Fala convertida em texto!';
                btnRec.style.backgroundColor = '#ff4b4b';
            }};

            recognition.onerror = (event) => {{
                statusTxt.innerText = '⚠️ Erro ao capturar áudio (' + event.error + ')';
                btnRec.style.backgroundColor = '#ff4b4b';
            }};

            recognition.onend = () => {{
                btnRec.style.backgroundColor = '#ff4b4b';
            }};
        }} else {{
            btnRec.disabled = true;
            statusTxt.innerText = 'Navegador sem suporte a Web Speech API. Utilize o microfone do teclado.';
        }}

        btnCopiar.onclick = () => {{
            txtArea.select();
            document.execCommand('copy');
            statusTxt.innerText = '📋 Texto copiado para a área de transferência!';
        }};
    </script>
    """
    components.html(html_code, height=160)

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
        status_registro TEXT DEFAULT 'Abastecido',
        foto_painel TEXT,
        observacao TEXT,
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
    
    frota = [
        ("25-1001", "S10", "Diesel", "TRX-6I85", "Ativa", 30000, 40000),
        ("25-1111", "S10", "Diesel", "TRX-4B85", "Ativa", 50000, 60000),
        ("25-1329", "Spin", "Gasolina", "TRZ-7E17", "Ativa", 0, 10000),
        ("25-1353", "Spin", "Gasolina", "TSC-2D46", "Emprestada", 0, 10000),
        ("25-1394", "Spin", "Gasolina", "UTS-3J57", "Ativa", 0, 10000)
    ]
    c.executemany("INSERT OR IGNORE INTO viaturas VALUES (?, ?, ?, ?, ?, ?, ?)", frota)
    
    try:
        c.execute("ALTER TABLE abastecimentos ADD COLUMN status_registro TEXT DEFAULT 'Abastecido'")
    except sqlite3.OperationalError:
        pass
    try:
        c.execute("ALTER TABLE abastecimentos ADD COLUMN foto_painel TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        c.execute("ALTER TABLE abastecimentos ADD COLUMN observacao TEXT")
    except sqlite3.OperationalError:
        pass

    conn.commit()
    conn.close()

def salvar_foto(foto_file, prefixo, tipo_evento):
    if foto_file is not None:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"{prefixo}_{tipo_evento}_{timestamp}.jpg"
        caminho = os.path.join(FOTOS_DIR, nome_arquivo)
        with open(caminho, "wb") as f:
            f.write(foto_file.getbuffer())
        return caminho
    return None

def converter_dfs_para_excel_multiplas_abas(df_abast, df_assuncao):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_abast.to_excel(writer, index=False, sheet_name='Abastecimentos')
        df_assuncao.to_excel(writer, index=False, sheet_name='Assuncao_Servico')
    return output.getvalue()

init_db()

if not check_login():
    st.stop()

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("🚔 10ª CICOM")
    st.write("Gestão Operacional de Frota")
    menu = st.radio("Navegação", [
        "📊 Painel de Controle (Dashboard)",
        "🤖 Leitura Automática de Print (OCR)",
        "📸 Check-in & Abastecimento", 
        "✏️ Corrigir / Editar Lançamento",
        "🛠️ Status da Frota & Revisões", 
        "📋 Livro de Ocorrências / Avarias", 
        "🔎 Consulta, Fotos & Exportação", 
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
    
    df_vtr_rev = pd.read_sql_query("SELECT * FROM viaturas", conn)
    for _, r_v in df_vtr_rev.iterrows():
        pref = r_v['prefixo']
        res = conn.execute("SELECT km_atual FROM abastecimentos WHERE prefixo = ? AND km_atual > 0 ORDER BY km_atual DESC LIMIT 1", (pref,)).fetchone()
        km_at = res[0] if res else 0.0
        km_rest = r_v['km_proxima_revisao'] - km_at
        if r_v['status'] == 'Ativa' and km_rest <= 500 and km_at > 0:
            st.error(f"🚨 **ALERTA CRÍTICO:** Viatura **{pref}** ({r_v['modelo']}) está a apenas **{km_rest:.0f} km** da revisão de {r_v['km_proxima_revisao']:.0f} km!")
        elif r_v['status'] == 'Ativa' and km_rest <= 1500 and km_at > 0:
            st.warning(f"⚠️ **ATENÇÃO:** Viatura **{pref}** ({r_v['modelo']}) está próxima da revisão ({km_rest:.0f} km restantes).")

    df_abast = pd.read_sql_query("SELECT * FROM abastecimentos", conn)
    df_vtr = pd.read_sql_query("SELECT * FROM viaturas", conn)
    
    if not df_abast.empty:
        df_so_abast = df_abast[df_abast['status_registro'] == 'Abastecido'].copy()
        df_so_abast['litros_num'] = df_so_abast['litros_liberados'].str.replace('L', '', case=False).astype(float, errors='ignore')
        df_so_abast['litros_num'] = pd.to_numeric(df_so_abast['litros_num'], errors='coerce').fillna(0)
        
        total_litros = df_so_abast['litros_num'].sum()
        total_abast = len(df_so_abast[df_so_abast['litros_num'] > 0])
        total_assuncoes = len(df_abast[df_abast['status_registro'] == 'Assunção de Serviço'])
        total_nao_abast = len(df_abast[df_abast['status_registro'] == 'Não Abasteceu (N/A)'])
        vtrs_ativas = len(df_vtr[df_vtr['status'] == 'Ativa'])
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Abastecimentos Realizados", f"{total_abast}")
        c2.metric("Assunções de Serviço (KM Inicial)", f"{total_assuncoes}")
        c3.metric("Total Combustível", f"{total_litros:.0f} Litros")
        c4.metric("Viaturas Ativas", f"{vtrs_ativas}")
        
        st.write("---")
        cg1, cg2 = st.columns(2)
        with cg1:
            st.write("##### ⛽ Consumo Real de Combustível por Viatura (Exclui KM Inicial)")
            consumo_vtr = df_so_abast.groupby('prefixo')['litros_num'].sum()
            st.bar_chart(consumo_vtr)
        with cg2:
            st.write("##### 📈 Distribuição dos Registros da Unidade")
            status_cont = df_abast['status_registro'].value_counts()
            st.bar_chart(status_cont)
    else:
        st.info("Nenhum registro cadastrado no momento.")

# 2. LEITURA AUTOMÁTICA DE PRINT / WHATSAPP (OCR)
elif menu == "🤖 Leitura Automática de Print (OCR)":
    st.subheader("🤖 Leitura Automática de Print do WhatsApp ou Foto do Painel")
    st.write("Envie o **print da conversa do WhatsApp** (com a foto e o texto) ou cole a mensagem. O sistema separará automaticamente se é Abastecimento ou KM Inicial.")
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        print_upload = st.file_uploader("Selecione o Print / Foto do WhatsApp", type=["jpg", "jpeg", "png"])
    with col_up2:
        texto_colado = st.text_area("Ou Cole o Texto do WhatsApp aqui (Opcional):", placeholder="Ex: VTR 25-1001, Cb Silva, KM Inicial: 45200")
        
    dados_extraidos = {
        "prefixo": "25-1001",
        "motorista": "",
        "km_atual": 0.0,
        "autonomia": 0.0,
        "tipo_operacao": "⛽ Realizar Abastecimento",
        "status_registro": "Abastecido"
    }

    if print_upload is not None:
        st.image(print_upload, caption="Print / Imagem Carregada", width=300)
        with st.spinner("🤖 Analisando imagem e extraindo texto com OCR..."):
            texto_ocr = extrair_texto_imagem(print_upload)
            texto_combinado = f"{texto_ocr} {texto_colado}"
            dados_extraidos = interpretar_dados_mensagem(texto_combinado)
            if texto_ocr:
                st.success("✅ Texto lido com sucesso pela imagem!")
                with st.expander("Ver texto detectado pelo OCR"):
                    st.write(texto_ocr)
    elif texto_colado:
        dados_extraidos = interpretar_dados_mensagem(texto_colado)

    st.write("---")
    st.write("#### 📝 Conferência dos Dados Extraídos (Confirme antes de Gravar)")
    
    with st.form("form_ocr_confirmar"):
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            lista_vtrs = ["25-1001", "25-1111", "25-1329", "25-1353", "25-1394"]
            idx_vtr = lista_vtrs.index(dados_extraidos["prefixo"]) if dados_extraidos["prefixo"] in lista_vtrs else 0
            prefixo_conf = st.selectbox("Prefixo da Viatura", lista_vtrs, index=idx_vtr)
            motorista_conf = st.text_input("Motorista", value=dados_extraidos["motorista"])
            km_conf = st.number_input("KM Odômetro", value=dados_extraidos["km_atual"], step=1.0)
        with col_c2:
            data_conf = st.date_input("Data", datetime.date.today())
            hora_conf = st.time_input("Horário", datetime.datetime.now().time())
            tipo_conf = st.selectbox("Tipo de Registro", [
                "⛽ Realizar Abastecimento", 
                "🕒 Assunção de Serviço (Entrada de Turno / KM Inicial)",
                "🚫 Informar que a Viatura Não Abasteceu Hoje (N/A)"
            ], index=["⛽ Realizar Abastecimento", "🕒 Assunção de Serviço (Entrada de Turno / KM Inicial)", "🚫 Informar que a Viatura Não Abasteceu Hoje (N/A)"].index(dados_extraidos["tipo_operacao"]))
            autonomia_conf = st.number_input("Autonomia informada (km)", value=dados_extraidos["autonomia"], step=1.0)

        liberado = True
        litros = "0L"
        status_reg_salvar = "Abastecido"
        obs = "Lançamento via Leitura Automática (OCR)"
        
        if tipo_conf == "⛽ Realizar Abastecimento":
            status_reg_salvar = "Abastecido"
            if prefixo_conf in ["25-1001", "25-1111"]:
                if 0 < autonomia_conf <= 70:
                    litros = "60L"
                elif 70 < autonomia_conf <= 180:
                    litros = "50L"
                else:
                    litros = "0L"
            else:
                if 0 < autonomia_conf <= 70:
                    litros = "40L"
                elif 95 <= autonomia_conf <= 115:
                    litros = "30L"
                elif 115 < autonomia_conf <= 125:
                    litros = "35L"
                else:
                    litros = "0L"
        elif tipo_conf == "🕒 Assunção de Serviço (Entrada de Turno / KM Inicial)":
            status_reg_salvar = "Assunção de Serviço"
            litros = "0L"
            obs = "KM Inicial de Entrada de Turno"
        elif tipo_conf == "🚫 Informar que a Viatura Não Abasteceu Hoje (N/A)":
            status_reg_salvar = "Não Abasteceu (N/A)"
            litros = "N/A"
            obs = "Não Abasteceu (N/A)"

        gravar_ocr = st.form_submit_button("💾 Confirmar e Gravar no Sistema", use_container_width=True)
        if gravar_ocr:
            if not motorista_conf:
                st.warning("Preencha o nome do motorista.")
            else:
                caminho_foto = salvar_foto(print_upload, prefixo_conf, status_reg_salvar[:5])
                c = conn.cursor()
                c.execute("""
                    INSERT OR IGNORE INTO abastecimentos 
                    (data, prefixo, motorista, km_atual, litros_liberados, horario, status_registro, foto_painel, observacao) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (data_conf.strftime("%Y-%m-%d"), prefixo_conf, motorista_conf, km_conf, litros, hora_conf.strftime("%H:%M"), status_reg_salvar, caminho_foto, obs))
                conn.commit()
                st.success(f"✅ Registro de **{status_reg_salvar}** gravado separadamente com sucesso!")
                st.rerun()

# 3. CHECK-IN & ABASTECIMENTO MANUAL
elif menu == "📸 Check-in & Abastecimento":
    st.subheader("📸 Registro Diário do Motorista")
    
    tipo_operacao = st.radio("Selecione o Tipo de Registro:", [
        "⛽ Realizar Abastecimento", 
        "🕒 Assunção de Serviço (Entrada de Turno / KM Inicial)",
        "🚫 Informar que a Viatura Não Abasteceu Hoje (N/A)"
    ], horizontal=True)
    
    st.write("---")
    
    col1, col2 = st.columns(2)
    with col1:
        prefixo = st.selectbox("Prefixo da Viatura", ["25-1001", "25-1111", "25-1329", "25-1353", "25-1394"])
        motorista = st.text_input("Nome de Guerra do Motorista (Farol)")
        km_atual = st.number_input("Quilometragem (KM Odômetro no Painel)", min_value=0.0, step=1.0)
    with col2:
        data = st.date_input("Data do Registro", datetime.date.today())
        hora = st.time_input("Horário", datetime.datetime.now().time())
        if tipo_operacao == "⛽ Realizar Abastecimento":
            autonomia = st.number_input("Autonomia Restante informada no painel (km)", min_value=0.0, step=1.0)
        else:
            autonomia = 0.0

    st.write("##### 📷 Foto do Painel da Viatura (Odômetro)")
    op_foto = st.radio("Como deseja enviar a foto?", ["📸 Câmera do Celular / Web", "📁 Galeria de Fotos / Arquivo"], horizontal=True)
    foto_capturada = None
    if op_foto == "📸 Câmera do Celular / Web":
        foto_capturada = st.camera_input("Tire a foto do painel mostrando o KM")
    else:
        foto_capturada = st.file_uploader("Selecione a foto do painel", type=["jpg", "jpeg", "png"])

    liberado = False
    litros = "0L"
    status_reg = "Abastecido"
    obs = ""

    if tipo_operacao == "⛽ Realizar Abastecimento":
        status_reg = "Abastecido"
        if prefixo == "25-1353":
            obs = "Viatura 25-1353 cedida/emprestada."
            st.error("⚠️ Viatura 25-1353 está cedida. Abastecimento bloqueado.")
        elif prefixo in ["25-1001", "25-1111"]:
            if 0 < autonomia <= 70:
                liberado = True
                litros = "60L"
                st.success("✅ COTA LIBERADA: 60 Litros de Diesel (Autonomia <= 70km)")
            elif 70 < autonomia <= 180:
                liberado = True
                litros = "50L"
                st.success("✅ COTA LIBERADA: 50 Litros de Diesel (Autonomia até 180km)")
            else:
                st.error(f"❌ Bloqueado: Autonomia ({autonomia}km) acima do limite permitido (>180km).")
        else:
            if 0 < autonomia <= 70:
                liberado = True
                litros = "40L"
                st.success("✅ COTA LIBERADA: 40 Litros de Gasolina (Autonomia <= 70km)")
            elif 95 <= autonomia <= 115:
                liberado = True
                litros = "30L"
                st.success("✅ COTA LIBERADA: 30 Litros de Gasolina (Autonomia entre 95 e 115km)")
            elif 115 < autonomia <= 125:
                liberado = True
                litros = "35L"
                st.success("✅ COTA LIBERADA: 35 Litros de Gasolina (Autonomia entre 115 e 125km)")
            else:
                st.error(f"❌ Bloqueado: Autonomia ({autonomia}km) fora das faixas permitidas.")

    elif tipo_operacao == "🕒 Assunção de Serviço (Entrada de Turno / KM Inicial)":
        liberado = True
        status_reg = "Assunção de Serviço"
        litros = "0L"
        st.write("##### 🎙️ Ditado por Voz para Observações de Entrada:")
        componente_ditado_voz("ditado_assuncao")
        obs = st.text_input("Observações de Entrada (Cole aqui o texto ditado ou digite)")

    elif tipo_operacao == "🚫 Informar que a Viatura Não Abasteceu Hoje (N/A)":
        liberado = True
        status_reg = "Não Abasteceu (N/A)"
        litros = "N/A"
        motivo_nao = st.selectbox("Motivo de Não Abastecer:", [
            "Autonomia suficiente para o turno", 
            "Viatura em Manutenção / Oficina",
            "Viatura de Reserva / Parada",
            "Viatura Cedida / Emprestada",
            "Cota não liberada"
        ])
        obs = motivo_nao

    if st.button("💾 Gravar Registro no Sistema", use_container_width=True):
        if not motorista:
            st.warning("Informe o nome do motorista.")
        elif tipo_operacao == "⛽ Realizar Abastecimento" and not liberado:
            st.error("Não é possível salvar abastecimento fora da cota autorizada.")
        else:
            caminho_foto = salvar_foto(foto_capturada, prefixo, status_reg[:5])
            c = conn.cursor()
            c.execute("""
                INSERT OR IGNORE INTO abastecimentos 
                (data, prefixo, motorista, km_atual, litros_liberados, horario, status_registro, foto_painel, observacao) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (data.strftime("%Y-%m-%d"), prefixo, motorista, km_atual, litros, hora.strftime("%H:%M"), status_reg, caminho_foto, obs))
            conn.commit()
            st.success(f"✅ Registro de **{status_reg}** gravado com sucesso!")

# 4. CORRIGIR / EDITAR LANÇAMENTO
elif menu == "✏️ Corrigir / Editar Lançamento":
    st.subheader("✏️ Correção e Edição de Lançamentos Anteriores")
    st.info("Selecione o registro para ajustar o KM, prefixo ou o tipo de evento (Abastecimento vs. Assunção):")
    
    df_todos = pd.read_sql_query("SELECT id, data, prefixo, motorista, km_atual, litros_liberados, horario, status_registro, observacao FROM abastecimentos ORDER BY id DESC", conn)
    
    if not df_todos.empty:
        opcoes_regs = [f"ID {r['id']} | [{r['status_registro']}] {r['data']} | VTR: {r['prefixo']} | KM: {r['km_atual']} | Mot: {r['motorista']}" for _, r in df_todos.iterrows()]
        escolha = st.selectbox("Selecione o registro para editar:", opcoes_regs)
        
        reg_id = int(escolha.split(" | ")[0].replace("ID ", ""))
        item = df_todos[df_todos['id'] == reg_id].iloc[0]
        
        st.write("---")
        with st.form("form_edicao"):
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                vtr_edit = st.selectbox("Prefixo da Viatura", ["25-1001", "25-1111", "25-1329", "25-1353", "25-1394"], index=["25-1001", "25-1111", "25-1329", "25-1353", "25-1394"].index(item['prefixo']) if item['prefixo'] in ["25-1001", "25-1111", "25-1329", "25-1353", "25-1394"] else 0)
                mot_edit = st.text_input("Motorista", value=str(item['motorista']))
                km_edit = st.number_input("KM Odômetro Correto", value=float(item['km_atual']), step=1.0)
                status_edit = st.selectbox("Classificação do Registro", ["Abastecido", "Assunção de Serviço", "Não Abasteceu (N/A)"], index=["Abastecido", "Assunção de Serviço", "Não Abasteceu (N/A)"].index(item['status_registro']) if item['status_registro'] in ["Abastecido", "Assunção de Serviço", "Não Abasteceu (N/A)"] else 0)
            with col_e2:
                try:
                    data_val = datetime.datetime.strptime(item['data'], "%Y-%m-%d").date()
                except:
                    data_val = datetime.date.today()
                data_edit = st.date_input("Data", value=data_val)
                litros_edit = st.text_input("Quantidade de Litros (ou N/A)", value=str(item['litros_liberados']))
                obs_edit = st.text_input("Observação / Motivo da Correção", value=str(item['observacao']) if item['observacao'] else "")
                
            c_btn1, c_btn2 = st.columns(2)
            with c_btn1:
                salvar_edit = st.form_submit_button("💾 Salvar Alterações Corrigidas", use_container_width=True)
            with c_btn2:
                excluir = st.form_submit_button("🗑️ Excluir Este Registro", use_container_width=True)
                
            if salvar_edit:
                c = conn.cursor()
                c.execute("""
                    UPDATE abastecimentos 
                    SET data = ?, prefixo = ?, motorista = ?, km_atual = ?, litros_liberados = ?, status_registro = ?, observacao = ?
                    WHERE id = ?
                """, (data_edit.strftime("%Y-%m-%d"), vtr_edit, mot_edit, km_edit, litros_edit, status_edit, obs_edit, reg_id))
                conn.commit()
                st.success("✅ Registro atualizado com sucesso! Os relatórios foram recalculados.")
                st.rerun()
                
            if excluir:
                c = conn.cursor()
                c.execute("DELETE FROM abastecimentos WHERE id = ?", (reg_id,))
                conn.commit()
                st.warning("🗑️ Registro excluído com sucesso!")
                st.rerun()
    else:
        st.info("Nenhum lançamento encontrado para edição.")

# 5. STATUS & REVISÕES
elif menu == "🛠️ Status da Frota & Revisões":
    st.subheader("🛠️ Monitoramento de Manutenção Preventiva")
    df_vtr = pd.read_sql_query("SELECT * FROM viaturas", conn)
    
    ultimos_kms = []
    restantes = []
    alertas = []
    for _, row in df_vtr.iterrows():
        pref = row['prefixo']
        res = conn.execute("SELECT km_atual FROM abastecimentos WHERE prefixo = ? AND km_atual > 0 ORDER BY km_atual DESC LIMIT 1", (pref,)).fetchone()
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

# 6. OCORRÊNCIAS / AVARIAS
elif menu == "📋 Livro de Ocorrências / Avarias":
    st.subheader("📋 Registro de Alterações, Arranhões e Sinistros")
    tab1, tab2 = st.tabs(["✍️ Digitação & Voz", "📄 Importar do Word (.docx)"])
    
    with tab1:
        st.write("##### 🎙️ Ditado por Voz para Relato da Avaria:")
        componente_ditado_voz("ditado_avaria")
        
        with st.form("form_ocorrencia_manual"):
            col1, col2 = st.columns(2)
            with col1:
                prefixo_oc = st.selectbox("Viatura", ["25-1001", "25-1111", "25-1329", "25-1353", "25-1394"], key="pref_m")
                tipo_oc = st.selectbox("Tipo de Ocorrência", ["Arranhão/Avaria Leve", "Sinistro/Colisão", "Problema Mecânico/Elétrico", "Cautela/Empréstimo", "Outro"], key="tipo_m")
                motorista_oc = st.text_input("Policial / Motorista Envolvido", key="mot_m")
            with col2:
                data_oc = st.date_input("Data do Ocorrido", datetime.date.today(), key="data_m")
                descricao = st.text_area("Descrição Detalhada (Cole aqui o texto ditado ou digite)", placeholder="Descreva o arranhão, sinistro ou alteração...", key="desc_m")
                
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
            doc = docx.Document(arquivo_word)
            texto_extraido = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
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
    df_oc = pd.read_sql_query("SELECT data_hora as Data, prefixo as Viatura, tipo as Tipo, motorista as Envolvido, descricao as Detalhes FROM ocorrencias ORDER BY id DESC", conn)
    st.dataframe(df_oc, use_container_width=True)

# 7. CONSULTA, FOTOS & EXPORTAÇÃO SEPARADA
elif menu == "🔎 Consulta, Fotos & Exportação":
    st.subheader("🔎 Consulta Separada por Categoria e Exportação")
    
    df_abast = pd.read_sql_query("""
        SELECT id, data as Data, prefixo as Viatura, motorista as Motorista, 
               km_atual as [KM Odômetro], litros_liberados as [Qtd Liberada], 
               horario as Horário, status_registro as [Tipo Evento], 
               observacao as [Motivo/Obs], foto_painel
        FROM abastecimentos 
        ORDER BY data DESC, id DESC
    """, conn)
    
    if not df_abast.empty:
        # Filtros no topo
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            filtro_vtr = st.selectbox("Filtrar Viatura", ["Todas"] + list(df_abast['Viatura'].unique()))
        with col_f2:
            filtro_motorista = st.text_input("Filtrar Motorista")
        with col_f3:
            filtro_data = st.date_input("Filtrar Data Específica", value=None)
            
        df_filtrado = df_abast.copy()
        if filtro_vtr != "Todas":
            df_filtrado = df_filtrado[df_filtrado['Viatura'] == filtro_vtr]
        if filtro_motorista:
            df_filtrado = df_filtrado[df_filtrado['Motorista'].str.contains(filtro_motorista, case=False, na=False)]
        if filtro_data is not None:
            df_filtrado = df_filtrado[df_filtrado['Data'] == filtro_data.strftime("%Y-%m-%d")]

        # Divisão em Abas Separadas para NÃO MISTURAR
        tab_abast, tab_assuncao, tab_nao, tab_todos = st.tabs([
            "⛽ Abastecimentos Reais", 
            "🕒 Assunção de Serviço (KM Inicial)", 
            "🚫 Não Abasteceu (N/A)", 
            "📋 Todos os Eventos"
        ])
        
        colunas_exibir = [c for c in df_filtrado.columns if c not in ['id', 'foto_painel']]
        
        with tab_abast:
            df_aba_abast = df_filtrado[df_filtrado['Tipo Evento'] == 'Abastecido']
            st.write(f"**Total de Abastecimentos Registrados:** {len(df_aba_abast)}")
            st.dataframe(df_aba_abast[colunas_exibir], use_container_width=True)
            
            # Galeria de fotos de abastecimento
            fotos_abast = df_aba_abast[df_aba_abast['foto_painel'].notna() & (df_aba_abast['foto_painel'] != '')]
            if not fotos_abast.empty:
                st.write("##### 📸 Comprovações Fotográficas dos Abastecimentos")
                cols = st.columns(3)
                for idx, (_, r_f) in enumerate(fotos_abast.iterrows()):
                    if os.path.exists(str(r_f['foto_painel'])):
                        with cols[idx % 3]:
                            st.image(r_f['foto_painel'], caption=f"{r_f['Viatura']} - {r_f['Data']} ({r_f['KM Odômetro']} km | {r_f['Qtd Liberada']})", use_container_width=True)

        with tab_assuncao:
            df_aba_assuncao = df_filtrado[df_filtrado['Tipo Evento'] == 'Assunção de Serviço']
            st.write(f"**Total de Entradas de Plantão (KM Inicial):** {len(df_aba_assuncao)}")
            st.dataframe(df_aba_assuncao[colunas_exibir], use_container_width=True)
            
            # Galeria de fotos de assunção
            fotos_assuncao = df_aba_assuncao[df_aba_assuncao['foto_painel'].notna() & (df_aba_assuncao['foto_painel'] != '')]
            if not fotos_assuncao.empty:
                st.write("##### 📸 Fotos de Painel na Entrada do Turno")
                cols = st.columns(3)
                for idx, (_, r_f) in enumerate(fotos_assuncao.iterrows()):
                    if os.path.exists(str(r_f['foto_painel'])):
                        with cols[idx % 3]:
                            st.image(r_f['foto_painel'], caption=f"{r_f['Viatura']} - {r_f['Data']} (KM Inicial: {r_f['KM Odômetro']})", use_container_width=True)

        with tab_nao:
            df_aba_nao = df_filtrado[df_filtrado['Tipo Evento'] == 'Não Abasteceu (N/A)']
            st.write(f"**Total de Justificativas (Sem Abastecer):** {len(df_aba_nao)}")
            st.dataframe(df_aba_nao[colunas_exibir], use_container_width=True)

        with tab_todos:
            st.write(f"**Histórico Completo Consolidado:** {len(df_filtrado)}")
            st.dataframe(df_filtrado[colunas_exibir], use_container_width=True)
        
        st.write("---")
        # Exportação com 2 abas no Excel
        excel_bytes = converter_dfs_para_excel_multiplas_abas(
            df_filtrado[df_filtrado['Tipo Evento'] == 'Abastecido'][colunas_exibir],
            df_filtrado[df_filtrado['Tipo Evento'] == 'Assunção de Serviço'][colunas_exibir]
        )
        st.download_button(
            label="📥 Baixar Planilha Excel Organizada (Abas Separadas: Abastecimento e Assunção)",
            data=excel_bytes,
            file_name="Relatorio_Frota_10CICOM_Abas_Separadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Nenhum dado encontrado.")

# 8. IMPORTAR PLANILHA EXCEL
elif menu == "📥 Importar Planilha Excel":
    st.subheader("📥 Importar Dados de Planilha de Abastecimento (.xlsx)")
    arquivo_excel = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xls", "XLSX", "XLS"])
    
    if arquivo_excel is not None:
        if st.button("Processar e Carregar Dados"):
            raw_df = pd.read_excel(arquivo_excel, header=None)
            vtr_rows = {4: "25-1001", 5: "25-1111", 6: "25-1329", 7: "25-1353", 8: "25-1394"}
            date_cols = [c for c in range(len(raw_df.columns)) if pd.notna(raw_df.iloc[2, c])]
            
            c = conn.cursor()
            novos = 0
            for col in date_cols:
                d_val = raw_df.iloc[2, col]
                data_str = d_val.strftime("%Y-%m-%d") if isinstance(d_val, pd.Timestamp) else str(d_val)[:10]
                for r, prefixo in vtr_rows.items():
                    km = raw_df.iloc[r, col]
                    motorista = raw_df.iloc[r, col+1] if col+1 < len(raw_df.columns) else ""
                    qtde = raw_df.iloc[r, col+2] if col+2 < len(raw_df.columns) else ""
                    horario = raw_df.iloc[r, col+3] if col+3 < len(raw_df.columns) else ""
                    
                    status_impor = "Abastecido"
                    km_num = 0.0
                    
                    if str(km).strip() in ['N/A', 'N/I', '', 'nan']:
                        status_impor = "Não Abasteceu (N/A)"
                        qtde = "N/A"
                    else:
                        try:
                            km_num = float(str(km).replace(',', '.').strip())
                        except ValueError:
                            status_impor = "Não Abasteceu (N/A)"
                            
                    c.execute("""
                        INSERT OR IGNORE INTO abastecimentos 
                        (data, prefixo, motorista, km_atual, litros_liberados, horario, status_registro, observacao)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (data_str, prefixo, str(motorista).strip(), km_num, str(qtde).strip(), str(horario).strip(), status_impor, "Importado da Planilha"))
                    if c.rowcount > 0:
                        novos += 1
            conn.commit()
            if novos > 0:
                st.success(f"Foram importados/atualizados {novos} registros!")
            else:
                st.info("Nenhuma novidade encontrada (todos os registros já estavam gravados).")

conn.close()
