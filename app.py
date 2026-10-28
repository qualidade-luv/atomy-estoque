import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime

# ============================================================
# CONFIGURAÇÃO INICIAL
# ============================================================
st.set_page_config(page_title="ATOMY Estoque", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

SPREADSHEET_ID = "11pc1Ll_iB438BHMl9NCBc3EEYqXEQjKANe6wlQyhax0"
ABA_BASE = "Base"
ABA_LOGIN = "Login"

HEADERS_BASE = [
    "PRODUTO", "UNIDADE", "ESTOQUE", "ENTRADAS", "SAIDAS", "SALDO",
    "VALIDADE", "ESTOQUE_MINIMO"
]

# ============================================================
# CONEXÃO COM GOOGLE SHEETS
# ============================================================
@st.cache_resource
def get_gspread_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource
def get_spreadsheet():
    client = get_gspread_client()
    return client.open_by_key(SPREADSHEET_ID)

def garantir_aba_e_cabecalhos():
    """Cria a aba 'Base' e os cabeçalhos se não existirem."""
    sh = get_spreadsheet()
    try:
        ws = sh.worksheet(ABA_BASE)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=ABA_BASE, rows=1000, cols=len(HEADERS_BASE))
        ws.append_row(HEADERS_BASE)
        return ws
    
    # Verifica se a primeira linha está vazia ou sem cabeçalhos
    valores = ws.row_values(1)
    if not valores or valores[:len(HEADERS_BASE)] != HEADERS_BASE:
        # Se a linha 1 estiver vazia, insere os cabeçalhos
        if not valores:
            ws.insert_row(HEADERS_BASE, index=1)
        # Se houver dados mas sem cabeçalhos, você pode optar por sobrescrever ou avisar
        # Aqui assumimos que a planilha já tem os cabeçalhos corretos
    return ws

# ============================================================
# AUTENTICAÇÃO (CASE-INSENSITIVE)
# ============================================================
def autenticar(login, senha):
    sh = get_spreadsheet()
    try:
        ws_login = sh.worksheet(ABA_LOGIN)
    except gspread.WorksheetNotFound:
        return False, "Aba 'Login' não encontrada na planilha."
    
    registros = ws_login.get_all_records()
    login_lower = login.strip().lower()
    senha_lower = senha.strip().lower()
    
    for reg in registros:
        user = str(reg.get("Login", "")).strip().lower()
        pwd = str(reg.get("Senha", "")).strip().lower()
        if user == login_lower and pwd == senha_lower:
            return True, login
    return False, "Login ou senha inválidos."

def tela_login():
    st.title("🔐 ATOMY - Controle de Estoque")
    st.subheader("Acesso ao Sistema")
    
    with st.form("login_form"):
        login = st.text_input("Login")
        senha = st.text_input("Senha", type="password")
        submit = st.form_submit_button("Entrar")
        
        if submit:
            if not login or not senha:
                st.warning("Preencha todos os campos.")
            else:
                ok, msg = autenticar(login, senha)
                if ok:
                    st.session_state["autenticado"] = True
                    st.session_state["usuario"] = msg
                    st.rerun()
                else:
                    st.error(msg)

# ============================================================
# LEITURA E ESCRITA NA BASE
# ============================================================
def carregar_dados():
    ws = garantir_aba_e_cabecalhos()
    dados = ws.get_all_records()
    if not dados:
        return pd.DataFrame(columns=HEADERS_BASE)
    df = pd.DataFrame(dados)
    # Garante que todas as colunas existam
    for col in HEADERS_BASE:
        if col not in df.columns:
            df[col] = ""
    return df[HEADERS_BASE]

def salvar_dados(df):
    """Sobrescreve a aba Base com o DataFrame fornecido (após confirmação)."""
    ws = garantir_aba_e_cabecalhos()
    # Limpa a partir da linha 2 (mantém cabeçalhos)
    ws.batch_clear([f"A2:{gspread.utils.rowcol_to_a1(len(df)+1, len(HEADERS_BASE))}"])
    if not df.empty:
        # Converte para lista de listas, garantindo strings
        valores = df.astype(str).values.tolist()
        ws.update(valores, f"A2", value_input_option="USER_ENTERED")
    # Invalida cache
    st.cache_data.clear()
    st.rerun()

# ============================================================
# INTERFACE PRINCIPAL (APÓS LOGIN)
# ============================================================
def app_principal():
    st.sidebar.success(f"Usuário: {st.session_state.get('usuario', '')}")
    if st.sidebar.button("Sair"):
        st.session_state.clear()
        st.rerun()
    
    st.title("📦 ATOMY - Controle de Estoque")
    
    df = carregar_dados()
    
    # -------- Exibição com indicadores --------
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total de Produtos", len(df))
    with col2:
        if "ESTOQUE_MINIMO" in df.columns and "SALDO" in df.columns:
            try:
                abaixo_min = df[pd.to_numeric(df["SALDO"], errors="coerce") < pd.to_numeric(df["ESTOQUE_MINIMO"], errors="coerce")]
                st.metric("Abaixo do Mínimo", len(abaixo_min))
            except:
                st.metric("Abaixo do Mínimo", "—")
    with col3:
        st.metric("Última Atualização", datetime.now().strftime("%d/%m/%Y %H:%M"))
    
    st.dataframe(df, use_container_width=True)
    
    st.divider()
    
    # -------- Operações CRUD --------
    aba_operacao = st.selectbox(
        "Selecione a operação:",
        ["Adicionar Produto", "Editar Produto", "Excluir Produto"]
    )
    
    if aba_operacao == "Adicionar Produto":
        with st.form("form_add"):
            st.subheader("➕ Adicionar Novo Produto")
            cols = st.columns(4)
            novo = {}
            novo["PRODUTO"] = cols[0].text_input("Produto *")
            novo["UNIDADE"] = cols[1].text_input("Unidade")
            novo["ESTOQUE"] = cols[2].text_input("Estoque Inicial")
            novo["VALIDADE"] = cols[3].text_input("Validade (DD/MM/AAAA)")
            
            cols2 = st.columns(4)
            novo["ENTRADAS"] = cols2[0].text_input("Entradas")
            novo["SAIDAS"] = cols2[1].text_input("Saídas")
            novo["SALDO"] = cols2[2].text_input("Saldo")
            novo["ESTOQUE_MINIMO"] = cols2[3].text_input("Estoque Mínimo")
            
            confirmar = st.checkbox("Confirmo que desejo adicionar este produto.")
            submit_add = st.form_submit_button("Adicionar")
            
            if submit_add:
                if not confirmar:
                    st.warning("Você precisa confirmar a operação (Sim).")
                elif not novo["PRODUTO"].strip():
                    st.warning("O campo PRODUTO é obrigatório.")
                else:
                    novo_df = pd.DataFrame([novo])
                    df_final = pd.concat([df, novo_df], ignore_index=True)
                    salvar_dados(df_final)
    
    elif aba_operacao == "Editar Produto":
        st.subheader("✏️ Editar Produto")
        if df.empty:
            st.info("Nenhum produto cadastrado.")
        else:
            produto_sel = st.selectbox("Selecione o produto para editar:", df["PRODUTO"].tolist())
            idx = df.index[df["PRODUTO"] == produto_sel][0]
            linha = df.loc[idx].to_dict()
            
            with st.form("form_edit"):
                cols = st.columns(4)
                editado = {}
                editado["PRODUTO"] = cols[0].text_input("Produto *", value=str(linha.get("PRODUTO", "")))
                editado["UNIDADE"] = cols[1].text_input("Unidade", value=str(linha.get("UNIDADE", "")))
                editado["ESTOQUE"] = cols[2].text_input("Estoque", value=str(linha.get("ESTOQUE", "")))
                editado["VALIDADE"] = cols[3].text_input("Validade", value=str(linha.get("VALIDADE", "")))
                
                cols2 = st.columns(4)
                editado["ENTRADAS"] = cols2[0].text_input("Entradas", value=str(linha.get("ENTRADAS", "")))
                editado["SAIDAS"] = cols2[1].text_input("Saídas", value=str(linha.get("SAIDAS", "")))
                editado["SALDO"] = cols2[2].text_input("Saldo", value=str(linha.get("SALDO", "")))
                editado["ESTOQUE_MINIMO"] = cols2[3].text_input("Estoque Mínimo", value=str(linha.get("ESTOQUE_MINIMO", "")))
                
                confirmar = st.checkbox("Confirmo que desejo salvar as alterações.")
                submit_edit = st.form_submit_button("Salvar Alterações")
                
                if submit_edit:
                    if not confirmar:
                        st.warning("Você precisa confirmar a operação (Sim).")
                    elif not editado["PRODUTO"].strip():
                        st.warning("O campo PRODUTO é obrigatório.")
                    else:
                        for k, v in editado.items():
                            df.at[idx, k] = v
                        salvar_dados(df)
    
    elif aba_operacao == "Excluir Produto":
        st.subheader("🗑️ Excluir Produto")
        if df.empty:
            st.info("Nenhum produto cadastrado.")
        else:
            produto_del = st.selectbox("Selecione o produto para excluir:", df["PRODUTO"].tolist())
            st.warning(f"Você está prestes a excluir permanentemente **{produto_del}**.")
            confirmar = st.checkbox("Confirmo que desejo excluir este produto.")
            if st.button("Excluir"):
                if not confirmar:
                    st.warning("Você precisa confirmar a operação (Sim).")
                else:
                    df = df[df["PRODUTO"] != produto_del]
                    salvar_dados(df)

# ============================================================
# FLUXO PRINCIPAL
# ============================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    tela_login()
else:
    app_principal()
