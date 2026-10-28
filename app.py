import json
import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from datetime import datetime

# ============================================================
# CONFIGURAÇÃO INICIAL
# ============================================================
st.set_page_config(page_title="ATOMY Estoque", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_ID = "11pc1Ll_iB438BHMl9NCBc3EEYqXEQjKANe6wlQyhax0"
ABA_BASE = "Base"
ABA_LOGIN = "Login"

HEADERS_BASE = [
    "PRODUTO", "UNIDADE", "ESTOQUE", "ENTRADAS", "SAIDAS", "SALDO",
    "VALIDADE", "ESTOQUE_MINIMO",
]
HEADERS_LOGIN = ["Login", "Senha"]


# ============================================================
# CONEXÃO COM GOOGLE SHEETS
# ============================================================
def _normalizar_credenciais():
    """
    Lê st.secrets['gcp_service_account'] e devolve um dict pronto
    para Credentials.from_service_account_info.

    Aceita dois formatos no secrets.toml:
      1) Tabela TOML  ->  [gcp_service_account]  type = ...  private_key = "..."
      2) String JSON  ->  gcp_service_account = '''{ ...json... }'''
    """
    if "gcp_service_account" not in st.secrets:
        raise ValueError(
            "Secrets ausente: adicione a seção [gcp_service_account] "
            "ou a chave gcp_service_account com o JSON da conta de serviço."
        )

    raw = st.secrets["gcp_service_account"]

    # Caso 1: veio como string JSON
    if isinstance(raw, str):
        try:
            creds_dict = json.loads(raw)
        except json.JSONDecodeError as e:
            raise ValueError(f"JSON da conta de serviço inválido: {e}")
    else:
        creds_dict = dict(raw)

    # Validação mínima dos campos exigidos pelo google-auth
    for campo in ("client_email", "token_uri", "private_key"):
        if not creds_dict.get(campo):
            raise ValueError(f"Credencial incompleta: campo '{campo}' ausente.")

    # ---- Correção da private_key -----------------------------------------
    # No TOML/Streamlit Cloud a chave costuma vir com "\n" literais
    # (barra invertida + n) em vez de quebras de linha reais.
    pk = creds_dict["private_key"]
    # Substitui a sequência literal \n por quebra real
    pk = pk.replace("\\n", "\n")
    # Remove espaços/quebras extras que alguns editores inserem no início/fim
    pk = pk.strip()
    # Garante o header/footer PEM
    if not pk.startswith("-----BEGIN PRIVATE KEY-----"):
        raise ValueError(
            "private_key inválida: não começa com '-----BEGIN PRIVATE KEY-----'. "
            "Copie o valor exatamente do arquivo JSON da conta de serviço."
        )
    if not pk.endswith("-----END PRIVATE KEY-----"):
        # Alguns JSONs vêm com \n final; normalizamos
        pk = pk.rstrip() + "\n-----END PRIVATE KEY-----\n"
    creds_dict["private_key"] = pk

    return creds_dict


@st.cache_resource(show_spinner="Conectando ao Google Sheets...")
def get_gspread_client():
    creds_dict = _normalizar_credenciais()
    try:
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    except Exception as e:
        raise ValueError(
            "Falha ao criar credenciais. Verifique se a private_key foi colada "
            f"corretamente nos Secrets. Detalhe: {e}"
        )
    return gspread.authorize(creds)


@st.cache_resource(show_spinner="Abrindo planilha...")
def get_spreadsheet():
    client = get_gspread_client()
    try:
        return client.open_by_key(SPREADSHEET_ID)
    except gspread.SpreadsheetNotFound:
        raise ValueError(
            "Planilha não encontrada ou sem permissão. "
            "Compartilhe a planilha como Editor com o client_email da conta de serviço."
        )


def garantir_abas():
    """Cria as abas 'Base' e 'Login' com cabeçalhos se não existirem."""
    sh = get_spreadsheet()

    # ---- Aba Base ----
    try:
        ws_base = sh.worksheet(ABA_BASE)
    except gspread.WorksheetNotFound:
        ws_base = sh.add_worksheet(
            title=ABA_BASE, rows=1000, cols=len(HEADERS_BASE)
        )
        ws_base.append_row(HEADERS_BASE)

    # Se a aba existe mas está vazia, escreve cabeçalhos
    valores_base = ws_base.row_values(1)
    if not valores_base:
        ws_base.insert_row(HEADERS_BASE, index=1)

    # ---- Aba Login ----
    try:
        ws_login = sh.worksheet(ABA_LOGIN)
    except gspread.WorksheetNotFound:
        ws_login = sh.add_worksheet(
            title=ABA_LOGIN, rows=100, cols=len(HEADERS_LOGIN)
        )
        ws_login.append_row(HEADERS_LOGIN)
        # Usuário padrão para primeiro acesso
        ws_login.append_row(["admin", "admin123"])

    valores_login = ws_login.row_values(1)
    if not valores_login:
        ws_login.insert_row(HEADERS_LOGIN, index=1)
        ws_login.append_row(["admin", "admin123"])

    return ws_base, ws_login


# ============================================================
# AUTENTICAÇÃO (CASE-INSENSITIVE)
# ============================================================
def autenticar(login: str, senha: str) -> tuple[bool, str]:
    _, ws_login = garantir_abas()
    registros = ws_login.get_all_records()

    login_norm = (login or "").strip().lower()
    senha_norm = (senha or "").strip().lower()

    for reg in registros:
        user = str(reg.get("Login", "")).strip().lower()
        pwd = str(reg.get("Senha", "")).strip().lower()
        if user == login_norm and pwd == senha_norm and user != "":
            return True, login
    return False, "Login ou senha inválidos."


def tela_login():
    st.title("🔐 ATOMY — Controle de Estoque")
    st.caption("Acesso restrito. Faça login para continuar.")

    with st.form("login_form", clear_on_submit=False):
        login = st.text_input("Login")
        senha = st.text_input("Senha", type="password")
        submit = st.form_submit_button("Entrar", use_container_width=True)

        if submit:
            if not login or not senha:
                st.warning("Preencha login e senha.")
            else:
                try:
                    ok, msg = autenticar(login, senha)
                except Exception as e:
                    st.error(f"Erro ao autenticar: {e}")
                    st.stop()
                if ok:
                    st.session_state["autenticado"] = True
                    st.session_state["usuario"] = msg
                    st.rerun()
                else:
                    st.error(msg)


# ============================================================
# LEITURA E ESCRITA NA BASE
# ============================================================
def carregar_dados() -> pd.DataFrame:
    ws_base, _ = garantir_abas()
    registros = ws_base.get_all_records()
    if not registros:
        return pd.DataFrame(columns=HEADERS_BASE)

    df = pd.DataFrame(registros)
    for col in HEADERS_BASE:
        if col not in df.columns:
            df[col] = ""
    # Mantém apenas as colunas oficiais e na ordem
    return df[HEADERS_BASE].astype(str)


def salvar_dados(df: pd.DataFrame, mensagem_sucesso: str = "Operação concluída."):
    """Sobrescreve o conteúdo da aba Base (mantendo cabeçalhos)."""
    ws_base, _ = garantir_abas()

    # Limpa dados antigos a partir da linha 2
    if ws_base.row_count > 1:
        ws_base.batch_clear(
            [f"A2:{gspread.utils.rowcol_to_a1(ws_base.row_count, len(HEADERS_BASE))}"]
        )

    if not df.empty:
        valores = df[HEADERS_BASE].astype(str).values.tolist()
        ws_base.update(valores, "A2", value_input_option="USER_ENTERED")

    # Invalida caches e recarrega
    st.cache_data.clear()
    st.toast(mensagem_sucesso, icon="✅")
    st.rerun()


# ============================================================
# INTERFACE PRINCIPAL
# ============================================================
def app_principal():
    # ---- Sidebar ----
    st.sidebar.success(f"👤 {st.session_state.get('usuario', '')}")
    if st.sidebar.button("Sair", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    st.title("📦 ATOMY — Controle de Estoque")

    try:
        df = carregar_dados()
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        st.stop()

    # ---- Indicadores ----
    c1, c2, c3 = st.columns(3)
    c1.metric("Total de Produtos", len(df))
    try:
        saldo = pd.to_numeric(df["SALDO"], errors="coerce")
        minimo = pd.to_numeric(df["ESTOQUE_MINIMO"], errors="coerce")
        abaixo = df[(saldo < minimo) & saldo.notna() & minimo.notna()]
        c2.metric("Abaixo do Mínimo", len(abaixo))
    except Exception:
        c2.metric("Abaixo do Mínimo", "—")
    c3.metric("Atualizado em", datetime.now().strftime("%d/%m/%Y %H:%M"))

    st.dataframe(df, use_container_width=True)
    st.divider()

    # ---- CRUD ----
    operacao = st.selectbox(
        "Selecione a operação:",
        ["Adicionar Produto", "Editar Produto", "Excluir Produto"],
    )

    # ---------- ADICIONAR ----------
    if operacao == "Adicionar Produto":
        with st.form("form_add"):
            st.subheader("➕ Adicionar Novo Produto")
            c = st.columns(4)
            novo = {
                "PRODUTO": c[0].text_input("Produto *"),
                "UNIDADE": c[1].text_input("Unidade"),
                "ESTOQUE": c[2].text_input("Estoque inicial"),
                "VALIDADE": c[3].text_input("Validade (DD/MM/AAAA)"),
            }
            c2 = st.columns(4)
            novo["ENTRADAS"] = c2[0].text_input("Entradas")
            novo["SAIDAS"] = c2[1].text_input("Saídas")
            novo["SALDO"] = c2[2].text_input("Saldo")
            novo["ESTOQUE_MINIMO"] = c2[3].text_input("Estoque mínimo")

            confirmar = st.checkbox("Confirmo que desejo adicionar este produto.")
            submit = st.form_submit_button("Adicionar", use_container_width=True)

            if submit:
                if not confirmar:
                    st.warning("Confirmação obrigatória: marque o checkbox para continuar.")
                elif not novo["PRODUTO"].strip():
                    st.warning("O campo PRODUTO é obrigatório.")
                else:
                    df_novo = pd.concat([df, pd.DataFrame([novo])], ignore_index=True)
                    try:
                        salvar_dados(df_novo, "Produto adicionado com sucesso.")
                    except Exception as e:
                        st.error(f"Erro ao salvar: {e}")

    # ---------- EDITAR ----------
    elif operacao == "Editar Produto":
        st.subheader("✏️ Editar Produto")
        if df.empty:
            st.info("Nenhum produto cadastrado.")
        else:
            produto = st.selectbox("Produto a editar:", df["PRODUTO"].tolist())
            idx = df.index[df["PRODUTO"] == produto][0]
            linha = df.loc[idx].to_dict()

            with st.form("form_edit"):
                c = st.columns(4)
                editado = {
                    "PRODUTO": c[0].text_input("Produto *", value=linha["PRODUTO"]),
                    "UNIDADE": c[1].text_input("Unidade", value=linha["UNIDADE"]),
                    "ESTOQUE": c[2].text_input("Estoque", value=linha["ESTOQUE"]),
                    "VALIDADE": c[3].text_input("Validade", value=linha["VALIDADE"]),
                }
                c2 = st.columns(4)
                editado["ENTRADAS"] = c2[0].text_input("Entradas", value=linha["ENTRADAS"])
                editado["SAIDAS"] = c2[1].text_input("Saídas", value=linha["SAIDAS"])
                editado["SALDO"] = c2[2].text_input("Saldo", value=linha["SALDO"])
                editado["ESTOQUE_MINIMO"] = c2[3].text_input(
                    "Estoque mínimo", value=linha["ESTOQUE_MINIMO"]
                )

                confirmar = st.checkbox("Confirmo que desejo salvar as alterações.")
                submit = st.form_submit_button("Salvar Alterações", use_container_width=True)

                if submit:
                    if not confirmar:
                        st.warning("Confirmação obrigatória: marque o checkbox para continuar.")
                    elif not editado["PRODUTO"].strip():
                        st.warning("O campo PRODUTO é obrigatório.")
                    else:
                        for k, v in editado.items():
                            df.at[idx, k] = v
                        try:
                            salvar_dados(df, "Produto atualizado com sucesso.")
                        except Exception as e:
                            st.error(f"Erro ao salvar: {e}")

    # ---------- EXCLUIR ----------
    else:
        st.subheader("🗑️ Excluir Produto")
        if df.empty:
            st.info("Nenhum produto cadastrado.")
        else:
            produto = st.selectbox("Produto a excluir:", df["PRODUTO"].tolist())
            st.warning(f"Você está prestes a excluir permanentemente **{produto}**.")
            confirmar = st.checkbox("Confirmo que desejo excluir este produto.")
            if st.button("Excluir", type="primary", use_container_width=True):
                if not confirmar:
                    st.warning("Confirmação obrigatória: marque o checkbox para continuar.")
                else:
                    df = df[df["PRODUTO"] != produto]
                    try:
                        salvar_dados(df, "Produto excluído com sucesso.")
                    except Exception as e:
                        st.error(f"Erro ao excluir: {e}")


# ============================================================
# FLUXO PRINCIPAL
# ============================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    tela_login()
else:
    app_principal()
