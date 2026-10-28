import json
import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from datetime import datetime, date

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

MARKER_BEGIN = "-----BEGIN PRIVATE KEY-----"
MARKER_END = "-----END PRIVATE KEY-----"

DIAS_ALERTA_AMARELO = 15

# ============================================================
# SANEAMENTO DA PRIVATE_KEY
# ============================================================
def _sanitizar_private_key(pk: str) -> str:
    if not pk:
        raise ValueError("private_key vazia nos Secrets.")
    pk = pk.replace("\\n", "\n").replace("\r", "")
    idx_begin = pk.find(MARKER_BEGIN)
    if idx_begin == -1:
        raise ValueError("Cabeçalho '-----BEGIN PRIVATE KEY-----' não encontrado.")
    pk = pk[idx_begin:]
    idx_end = pk.find(MARKER_END)
    if idx_end == -1:
        raise ValueError("Rodapé '-----END PRIVATE KEY-----' não encontrado.")
    pk = pk[: idx_end + len(MARKER_END)]
    linhas = [l.strip() for l in pk.split("\n") if l.strip()]
    if linhas[0] != MARKER_BEGIN:
        linhas.insert(0, MARKER_BEGIN)
    if linhas[-1] != MARKER_END:
        linhas.append(MARKER_END)
    return "\n".join(linhas) + "\n"

# ============================================================
# CONEXÃO COM GOOGLE SHEETS
# ============================================================
def _normalizar_credenciais() -> dict:
    if "gcp_service_account" not in st.secrets:
        raise ValueError("Secrets ausente: adicione 'gcp_service_account'.")
    raw = st.secrets["gcp_service_account"]
    if isinstance(raw, str):
        creds_dict = json.loads(raw)
    else:
        creds_dict = dict(raw)
    for campo in ("client_email", "token_uri", "private_key"):
        if not creds_dict.get(campo):
            raise ValueError(f"Credencial incompleta: '{campo}' ausente.")
    creds_dict["private_key"] = _sanitizar_private_key(creds_dict["private_key"])
    return creds_dict

@st.cache_resource(show_spinner="Conectando ao Google Sheets...")
def get_gspread_client():
    creds_dict = _normalizar_credenciais()
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner="Abrindo planilha...")
def get_spreadsheet():
    client = get_gspread_client()
    try:
        return client.open_by_key(SPREADSHEET_ID)
    except gspread.SpreadsheetNotFound:
        raise ValueError("Planilha não encontrada ou sem permissão.")

def garantir_abas():
    sh = get_spreadsheet()
    try:
        ws_base = sh.worksheet(ABA_BASE)
    except gspread.WorksheetNotFound:
        ws_base = sh.add_worksheet(title=ABA_BASE, rows=1000, cols=len(HEADERS_BASE))
        ws_base.append_row(HEADERS_BASE)
    if not ws_base.row_values(1):
        ws_base.insert_row(HEADERS_BASE, index=1)

    try:
        ws_login = sh.worksheet(ABA_LOGIN)
    except gspread.WorksheetNotFound:
        ws_login = sh.add_worksheet(title=ABA_LOGIN, rows=100, cols=len(HEADERS_LOGIN))
        ws_login.append_row(HEADERS_LOGIN)
        ws_login.append_row(["admin", "admin123"])
    if not ws_login.row_values(1):
        ws_login.insert_row(HEADERS_LOGIN, index=1)
        ws_login.append_row(["admin", "admin123"])

    return ws_base, ws_login

# ============================================================
# AUTENTICAÇÃO
# ============================================================
def autenticar(login: str, senha: str) -> tuple[bool, str]:
    _, ws_login = garantir_abas()
    registros = ws_login.get_all_records()
    login_norm = (login or "").strip().lower()
    senha_norm = (senha or "").strip().lower()
    for reg in registros:
        user = str(reg.get("Login", "")).strip().lower()
        pwd = str(reg.get("Senha", "")).strip().lower()
        if user and user == login_norm and pwd == senha_norm:
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
# HELPERS NUMÉRICOS / DATA
# ============================================================
def _to_num(serie) -> float:
    if isinstance(serie, pd.Series):
        s = serie.astype(str).str.strip().str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        s = s.replace({"": None, "nan": None, "None": None})
        return pd.to_numeric(s, errors="coerce").fillna(0)
    try:
        return float(str(serie).replace(".", "").replace(",", "."))
    except Exception:
        return 0.0

def _parse_data(txt) -> date | None:
    if txt is None: return None
    s = str(txt).strip()
    if not s or s.lower() in ("nan", "none", ""): return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

def _calc_dias_validade(txt) -> int | None:
    d = _parse_data(txt)
    if d is None: return None
    return (d - date.today()).days

# ============================================================
# LEITURA E ESCRITA
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
    
    # AUTO-CORREÇÃO: Se o SALDO estiver vazio na planilha antiga, calcula e preenche
    for i in df.index:
        if str(df.at[i, "SALDO"]).strip() in ("", "nan", "None", "0.0"):
            e = _to_num(df.at[i, "ESTOQUE"])
            ent = _to_num(df.at[i, "ENTRADAS"])
            sai = _to_num(df.at[i, "SAIDAS"])
            df.at[i, "SALDO"] = str(int(e + ent - sai))

    return df[HEADERS_BASE].astype(str)

def salvar_dados(df: pd.DataFrame, mensagem_sucesso: str = "Operação concluída."):
    ws_base, _ = garantir_abas()
    if ws_base.row_count > 1:
        ws_base.batch_clear([f"A2:{gspread.utils.rowcol_to_a1(ws_base.row_count, len(HEADERS_BASE))}"])
    if not df.empty:
        valores = df[HEADERS_BASE].astype(str).values.tolist()
        ws_base.update(valores, "A2", value_input_option="USER_ENTERED")
    st.cache_data.clear()
    st.toast(mensagem_sucesso, icon="✅")
    st.rerun()

# ============================================================
# ESTILIZAÇÃO CONDICIONAL
# ============================================================
def _linha_abaixo_minimo(row) -> bool:
    saldo = _to_num(row.get("SALDO", 0))
    minimo = _to_num(row.get("ESTOQUE_MINIMO", 0))
    return minimo > 0 and saldo < minimo

def _linha_vencida(row) -> bool:
    dias = row.get("_DIAS_VALIDADE")
    return dias is not None and not pd.isna(dias) and dias < 0

def _linha_amarela(row) -> bool:
    dias = row.get("_DIAS_VALIDADE")
    return dias is not None and not pd.isna(dias) and 0 <= dias <= DIAS_ALERTA_AMARELO

def _estilo_linha(row):
    if _linha_vencida(row):
        return ["background-color: #ff4d4d; color: white;"] * len(row)
    if _linha_amarela(row):
        return ["background-color: #fff3b0; color: #333;"] * len(row)
    if _linha_abaixo_minimo(row):
        return ["background-color: #ffb347; color: #222;"] * len(row)
    return [""] * len(row)

def preparar_df_exibicao(df: pd.DataFrame) -> pd.DataFrame:
    df_exib = df.copy()
    df_exib["_DIAS_VALIDADE"] = df_exib["VALIDADE"].apply(_calc_dias_validade)
    def _label(dias):
        if dias is None or pd.isna(dias): return "—"
        if dias < 0: return f"VENCIDO ({abs(dias)}d)"
        return f"{dias} dias"
    df_exib["VALIDADE_RESTANTE"] = df_exib["_DIAS_VALIDADE"].apply(_label)
    return df_exib

# ============================================================
# INTERFACE PRINCIPAL
# ============================================================
def app_principal():
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

    # ---------- Indicadores ----------
    df_calc = df.copy()
    saldo = _to_num(df_calc["SALDO"])
    minimo = _to_num(df_calc["ESTOQUE_MINIMO"])
    abaixo_min = ((minimo > 0) & (saldo < minimo)).sum()
    df_calc["_DIAS_VALIDADE"] = df_calc["VALIDADE"].apply(_calc_dias_validade)
    vencidos = df_calc["_DIAS_VALIDADE"].apply(lambda d: d is not None and not pd.isna(d) and d < 0).sum()
    a_vencer = df_calc["_DIAS_VALIDADE"].apply(lambda d: d is not None and not pd.isna(d) and 0 <= d <= DIAS_ALERTA_AMARELO).sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de Produtos", len(df))
    c2.metric("Abaixo do Mínimo", int(abaixo_min))
    c3.metric("Vencidos", int(vencidos))
    c4.metric("A vencer (≤15 dias)", int(a_vencer))

    # ---------- Tabela estilizada ----------
    st.subheader("📋 Estoque atual")
    if df.empty:
        st.info("Nenhum produto cadastrado ainda.")
    else:
        df_exib = preparar_df_exibicao(df)
        colunas_exibir = HEADERS_BASE + ["VALIDADE_RESTANTE"]
        df_view = df_exib[colunas_exibir + ["_DIAS_VALIDADE"]]

        styled = (
            df_view.style
            .apply(_estilo_linha, axis=1)
            .format({"_DIAS_VALIDADE": lambda v: "" if pd.isna(v) else int(v)})
            .hide(axis="columns", subset=["_DIAS_VALIDADE"])
        )
        st.dataframe(styled, use_container_width=True)
        st.caption("🟧 **Laranja**: saldo abaixo do mínimo | 🟨 **Amarelo**: ≤ 15 dias | 🟥 **Vermelho**: vencido")

    st.divider()

    # ---------- CRUD ----------
    operacao = st.selectbox(
        "Selecione a operação:",
        ["Adicionar Produto", "Editar Produto", "Excluir Produto", "Registrar Entrada", "Registrar Saída"],
    )

    # ---------- ADICIONAR ----------
    if operacao == "Adicionar Produto":
        with st.form("form_add"):
            st.subheader("➕ Adicionar Novo Produto")
            c = st.columns(4)
            novo = {
                "PRODUTO": c[0].text_input("Produto *"),
                "UNIDADE": c[1].text_input("Unidade"),
                "ESTOQUE": c[2].text_input("Estoque inicial", value="0"),
                "VALIDADE": c[3].text_input("Validade (DD/MM/AAAA)"),
            }
            c2 = st.columns(4)
            novo["ENTRADAS"] = c2[0].text_input("Entradas", value="0")
            novo["SAIDAS"] = c2[1].text_input("Saídas", value="0")
            novo["ESTOQUE_MINIMO"] = c2[3].text_input("Estoque mínimo", value="0")

            estoque_ini = _to_num(novo["ESTOQUE"])
            entradas = _to_num(novo["ENTRADAS"])
            saidas = _to_num(novo["SAIDAS"])
            novo["SALDO"] = str(int(estoque_ini + entradas - saidas))
            st.info(f"Saldo calculado: **{novo['SALDO']}**")

            confirmar = st.checkbox("Confirmo que desejo adicionar este produto.")
            submit = st.form_submit_button("Adicionar", use_container_width=True)

            if submit:
                if not confirmar:
                    st.warning("Confirmação obrigatória.")
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
                editado["ESTOQUE_MINIMO"] = c2[3].text_input("Estoque mínimo", value=linha["ESTOQUE_MINIMO"])

                estoque_e = _to_num(editado["ESTOQUE"])
                entradas_e = _to_num(editado["ENTRADAS"])
                saidas_e = _to_num(editado["SAIDAS"])
                editado["SALDO"] = str(int(estoque_e + entradas_e - saidas_e))
                st.info(f"Saldo recalculado: **{editado['SALDO']}**")

                confirmar = st.checkbox("Confirmo que desejo salvar as alterações.")
                submit = st.form_submit_button("Salvar Alterações", use_container_width=True)

                if submit:
                    if not confirmar:
                        st.warning("Confirmação obrigatória.")
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
    elif operacao == "Excluir Produto":
        st.subheader("🗑️ Excluir Produto")
        if df.empty:
            st.info("Nenhum produto cadastrado.")
        else:
            produto = st.selectbox("Produto a excluir:", df["PRODUTO"].tolist())
            st.warning(f"Você está prestes a excluir permanentemente **{produto}**.")
            confirmar = st.checkbox("Confirmo que desejo excluir este produto.")
            if st.button("Excluir", type="primary", use_container_width=True):
                if not confirmar:
                    st.warning("Confirmação obrigatória.")
                else:
                    df = df[df["PRODUTO"] != produto]
                    try:
                        salvar_dados(df, "Produto excluído com sucesso.")
                    except Exception as e:
                        st.error(f"Erro ao excluir: {e}")

    # ---------- REGISTRAR ENTRADA ----------
    elif operacao == "Registrar Entrada":
        st.subheader("📥 Registrar Entrada de Estoque")
        if df.empty:
            st.info("Cadastre um produto antes de registrar entradas.")
        else:
            produto = st.selectbox("Produto:", df["PRODUTO"].tolist())
            idx = df.index[df["PRODUTO"] == produto][0]
            saldo_atual = _to_num(df.at[idx, "SALDO"])

            with st.form("form_entrada"):
                qtd = st.number_input("Quantidade a acrescentar:", min_value=1, step=1, value=1)
                st.write(f"Saldo atual: **{int(saldo_atual)}** → Saldo após entrada: **{int(saldo_atual + qtd)}**")
                confirmar = st.checkbox("Confirmo o registro desta ENTRADA.")
                submit = st.form_submit_button("Registrar Entrada", use_container_width=True)

                if submit:
                    if not confirmar:
                        st.warning("Confirmação obrigatória.")
                    else:
                        entradas_atuais = _to_num(df.at[idx, "ENTRADAS"])
                        estoque_atual = _to_num(df.at[idx, "ESTOQUE"])
                        saidas_atuais = _to_num(df.at[idx, "SAIDAS"])

                        df.at[idx, "ENTRADAS"] = str(int(entradas_atuais + qtd))
                        df.at[idx, "ESTOQUE"] = str(int(estoque_atual + qtd))
                        df.at[idx, "SALDO"] = str(int(_to_num(df.at[idx, "ESTOQUE"]) + _to_num(df.at[idx, "ENTRADAS"]) - _to_num(df.at[idx, "SAIDAS"])))
                        try:
                            salvar_dados(df, f"Entrada de {int(qtd)} registrada.")
                        except Exception as e:
                            st.error(f"Erro ao salvar: {e}")

    # ---------- REGISTRAR SAÍDA ----------
    elif operacao == "Registrar Saída":
        st.subheader("📤 Registrar Saída de Estoque")
        if df.empty:
            st.info("Cadastre um produto antes de registrar saídas.")
        else:
            produto = st.selectbox("Produto:", df["PRODUTO"].tolist())
            idx = df.index[df["PRODUTO"] == produto][0]
            saldo_atual = _to_num(df.at[idx, "SALDO"])

            with st.form("form_saida"):
                qtd = st.number_input("Quantidade a retirar:", min_value=1, step=1, value=1)
                novo_saldo = saldo_atual - qtd
                if novo_saldo < 0:
                    st.error(f"Saída maior que o saldo disponível ({int(saldo_atual)}). Ajuste a quantidade.")
                st.write(f"Saldo atual: **{int(saldo_atual)}** → Saldo após saída: **{int(novo_saldo)}**")
                confirmar = st.checkbox("Confirmo o registro desta SAÍDA.")
                submit = st.form_submit_button("Registrar Saída", use_container_width=True)

                if submit:
                    if not confirmar:
                        st.warning("Confirmação obrigatória.")
                    elif novo_saldo < 0:
                        st.error("Saldo insuficiente. Operação cancelada.")
                    else:
                        saidas_atuais = _to_num(df.at[idx, "SAIDAS"])
                        estoque_atual = _to_num(df.at[idx, "ESTOQUE"])

                        df.at[idx, "SAIDAS"] = str(int(saidas_atuais + qtd))
                        df.at[idx, "ESTOQUE"] = str(int(estoque_atual - qtd))
                        df.at[idx, "SALDO"] = str(int(_to_num(df.at[idx, "ESTOQUE"]) + _to_num(df.at[idx, "ENTRADAS"]) - _to_num(df.at[idx, "SAIDAS"])))
                        try:
                            salvar_dados(df, f"Saída de {int(qtd)} registrada.")
                        except Exception as e:
                            st.error(f"Erro ao salvar: {e}")

# ============================================================
# FLUXO PRINCIPAL
# ============================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    tela_login()
else:
    app_principal()
