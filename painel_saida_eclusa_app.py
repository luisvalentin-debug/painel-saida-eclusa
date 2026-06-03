import io
import html
import re
import textwrap
import unicodedata
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from streamlit_autorefresh import st_autorefresh

# =========================
# CONFIGURAÇÃO DA PÁGINA
# =========================
# App atualizado: V3 - correção robusta de datas
st.set_page_config(
    page_title="Painel Saída Eclusa",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="collapsed",
)

TIMEZONE = "America/Sao_Paulo"
AUTOREFRESH_SECONDS = 60
HEADER_ROW_EXCEL = 2  # títulos começam na linha 3 do Excel/Google Sheets
APP_VERSION = "V33 - link fixo Google Sheets e filtro data status"
DEFAULT_GOOGLE_SHEETS_URL = "https://docs.google.com/spreadsheets/d/12fMDx0ih2P8LNTRa1EL987GYyPsZv2dDP4iH54Qjun8/edit?gid=2029130478#gid=2029130478"


def get_query_value(nome: str, default: str = "") -> str:
    """Lê parâmetro da URL. Ex.: ?turno=1%C2%BA&modal=VENDA"""
    try:
        valor = st.query_params.get(nome, default)
        if isinstance(valor, list):
            valor = valor[0] if valor else default
        return str(valor).strip()
    except Exception:
        return default


def normalizar_url_value(valor: str) -> str:
    return str(valor or "").strip().lower().replace("_", " ").replace("-", " ")


def escolher_por_query(options, query_name: str, default_label: str):
    """Retorna index de selectbox baseado em query param, aceitando texto sem acento/case."""
    raw = get_query_value(query_name, "")
    if not raw:
        return options.index(default_label) if default_label in options else 0
    alvo = normalizar_url_value(raw)
    aliases = {
        "todos": "Todos",
        "todas": "Todas",
        "claro": "Claro",
        "escuro": "Escuro",
        "venda": "VENDA",
        "transferencia": "TRANSFERÊNCIA",
        "transferência": "TRANSFERÊNCIA",
        "transf": "TRANSFERÊNCIA",
        "aba": "ABA",
        "pendentes": "Somente pendentes",
        "somente pendentes": "Somente pendentes",
        "com saida": "Somente com saída",
        "com saída": "Somente com saída",
        "somente com saida": "Somente com saída",
        "somente com saída": "Somente com saída",
        "2h": "Até 2h para frente",
        "ate 2h": "Até 2h para frente",
        "até 2h": "Até 2h para frente",
        "hoje": "Hoje",
        "amanha": "Amanhã",
        "amanhã": "Amanhã",
        "dias anteriores": "Dias anteriores",
        "todos ate amanha": "Todos até amanhã",
        "todos até amanhã": "Todos até amanhã",
    }
    candidato = aliases.get(alvo, raw)
    for i, opt in enumerate(options):
        if normalizar_url_value(opt) == normalizar_url_value(candidato):
            return i
    return options.index(default_label) if default_label in options else 0

# =========================
# FUNÇÕES AUXILIARES
# =========================
def norm_txt(value: str) -> str:
    """Normaliza texto para comparação de colunas/status."""
    if value is None:
        return ""
    value = str(value).strip()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"\s+", " ", value)
    return value.upper()


def find_col(df: pd.DataFrame, aliases: list[str]) -> str | None:
    """Localiza uma coluna por uma lista de possíveis nomes."""
    normalized_cols = {norm_txt(c): c for c in df.columns}
    for alias in aliases:
        key = norm_txt(alias)
        if key in normalized_cols:
            return normalized_cols[key]
    return None


def excel_serial_to_datetime(series: pd.Series) -> pd.Series:
    """
    Converte datas do Excel/Google Sheets com proteção contra overflow.

    Regras:
    - se a coluna já veio como datetime, apenas mantém;
    - se veio como texto, tenta converter como data brasileira;
    - se veio como número serial do Excel, só converte números plausíveis
      entre 1 e 100000. Isso evita o erro OutOfBoundsTimedelta.
    """
    if series is None:
        return pd.Series(dtype="datetime64[ns]")

    s = series.copy()

    # Quando o Excel já trouxe datetime64, não tente converter como número.
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")

    # Primeira tentativa: texto/data comum.
    text_dt = pd.to_datetime(s, errors="coerce", dayfirst=True)

    # Segunda tentativa: serial numérico real do Excel.
    numeric = pd.to_numeric(s, errors="coerce")
    is_excel_serial = numeric.between(1, 100000)

    serial_dt = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    if bool(is_excel_serial.any()):
        serial_dt.loc[is_excel_serial] = (
            pd.Timestamp("1899-12-30")
            + pd.to_timedelta(numeric.loc[is_excel_serial], unit="D")
        )

    # Usa serial onde for serial válido; no restante, usa texto/data comum.
    return serial_dt.combine_first(text_dt)


def empty_dt(value) -> bool:
    return pd.isna(value)


def empty_value(value) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def format_dt(value) -> str:
    if pd.isna(value):
        return "-"
    return pd.to_datetime(value).strftime("%d/%m %H:%M")


def format_tempo(minutos: float, prefixo: str = "") -> str:
    if pd.isna(minutos):
        return "-"
    minutos = int(round(abs(minutos)))
    h = minutos // 60
    m = minutos % 60
    if h > 0:
        txt = f"{h}h {m:02d}m"
    else:
        txt = f"{m}m"
    return f"{prefixo}{txt}" if prefixo else txt


def clean_modal(value: str) -> str:
    v = norm_txt(value)
    if "VENDA" in v or "VEND" in v:
        return "VENDA"
    if "TRANS" in v or "TRANSFER" in v:
        return "TRANSFERÊNCIA"
    if "ABA" in v or "ABAST" in v:
        return "ABA"
    return str(value).strip().upper() if not empty_value(value) else "-"


def modal_prioritario(values) -> str:
    vals = [clean_modal(v) for v in values if not empty_value(v)]
    # Prioridade operacional: se tiver venda em qualquer linha da carga, destacar como VENDA
    if "VENDA" in vals:
        return "VENDA"
    if "ABA" in vals:
        return "ABA"
    if "TRANSFERÊNCIA" in vals:
        return "TRANSFERÊNCIA"
    return vals[0] if vals else "-"


def first_valid(values):
    for v in values:
        if not empty_value(v):
            return v
    return None


def parse_numero_br(series: pd.Series) -> pd.Series:
    """
    Converte números vindos do Excel/Google Sheets em formato BR ou US.
    Ex.: 12,5 -> 12.5 | 1.234,56 -> 1234.56 | 12.5 -> 12.5.
    Evita que valores de M³ vindos como texto com vírgula virem zero.
    """
    if series is None:
        return pd.Series(dtype="float64")
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0)

    s = series.astype(str).str.strip()
    s = s.replace({"": "0", "nan": "0", "None": "0", "-": "0"})
    # remove espaços e símbolos, mantendo dígitos, vírgula, ponto e sinal
    s = s.str.replace(r"[^0-9,.-]", "", regex=True)

    def conv(x: str):
        if not x or x in {"-", ",", "."}:
            return 0
        # Quando tem vírgula, assume vírgula como decimal e ponto como milhar.
        if "," in x:
            x = x.replace(".", "").replace(",", ".")
        return x

    return pd.to_numeric(s.map(conv), errors="coerce").fillna(0)


# =========================
# LEITURA DA BASE
# =========================
@st.cache_data(ttl=AUTOREFRESH_SECONDS)
def load_excel(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), header=HEADER_ROW_EXCEL)


@st.cache_data(ttl=AUTOREFRESH_SECONDS)
def load_google_csv(csv_url: str) -> pd.DataFrame:
    return pd.read_csv(csv_url, header=HEADER_ROW_EXCEL)


def google_sheets_to_csv_url(url: str) -> str:
    """
    Aceita link normal do Google Sheets ou link CSV publicado e devolve uma URL CSV.
    Funciona melhor quando a planilha está compartilhada/publicada para leitura.
    """
    url = str(url or "").strip()
    if not url:
        return url
    if "format=csv" in url or "output=csv" in url:
        return url
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not m:
        return url
    sheet_id = m.group(1)
    gid_match = re.search(r"gid=([0-9]+)", url)
    gid = gid_match.group(1) if gid_match else "0"
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


def preparar_base(df_raw: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    df = df_raw.copy()
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    col = {
        "id_carga": find_col(df, ["ID CARGA", "ID Carga", "ID"]),
        "plano": find_col(df, ["Plano de Transporte", "Plano"]),
        "transportador": find_col(df, ["Transportadora", "Transportador", "Transp"]),
        "turno": find_col(df, ["Turno"]),
        "janela": find_col(df, ["Janela"]),
        "chegada": find_col(df, ["Chegada Portaria", "Chegada"]),
        "saida": find_col(df, ["Saída Eclusa", "Saida Eclusa", "Saida Veiculo", "Saída Veículo"]),
        "limite_indicador": find_col(df, ["Limite Indicador"]),
        "limite_janela": find_col(df, ["Limite Janela"]),
        "modal": find_col(df, ["Modal", "Modalidade"]),
        "doca": find_col(df, ["Doca"]),
        "pedidos": find_col(df, ["Pedidos", "Ped Venda"]),
        "pecas": find_col(df, ["Peças", "Pecas"]),
        "m3": find_col(df, ["M³", "M3", "M³ Total", "M3 Total"]),
    }

    obrigatorias = [
        "id_carga", "plano", "transportador", "turno", "janela",
        "chegada", "saida", "limite_indicador", "limite_janela", "modal"
    ]
    faltantes = [k for k in obrigatorias if col[k] is None]
    if faltantes:
        return pd.DataFrame(), faltantes

    # Conversões de datas
    for key in ["janela", "chegada", "saida", "limite_indicador", "limite_janela"]:
        df[col[key]] = excel_serial_to_datetime(df[col[key]])

    # Valores numéricos para somar, quando existirem
    # No Google Sheets, M³ e Pedidos podem vir como texto em formato brasileiro (ex.: "0,5").
    # Por isso usamos parse_numero_br para não transformar os valores em zero.
    for key in ["pedidos", "pecas", "m3"]:
        if col[key]:
            df[col[key]] = parse_numero_br(df[col[key]])

    # Consolida uma linha por ID Carga
    agg_dict = {
        col["plano"]: first_valid,
        col["transportador"]: first_valid,
        col["turno"]: first_valid,
        col["janela"]: first_valid,
        col["chegada"]: first_valid,
        col["saida"]: first_valid,
        col["limite_indicador"]: first_valid,
        col["limite_janela"]: first_valid,
        col["modal"]: modal_prioritario,
    }
    if col["doca"]:
        agg_dict[col["doca"]] = first_valid
    if col["pedidos"]:
        agg_dict[col["pedidos"]] = "sum"
    if col["pecas"]:
        agg_dict[col["pecas"]] = "sum"
    if col["m3"]:
        agg_dict[col["m3"]] = "sum"

    base = df.groupby(col["id_carga"], dropna=True).agg(agg_dict).reset_index()

    rename = {
        col["id_carga"]: "ID Carga",
        col["plano"]: "Plano de Transporte",
        col["transportador"]: "Transportador",
        col["turno"]: "Turno",
        col["janela"]: "Janela",
        col["chegada"]: "Chegada Portaria",
        col["saida"]: "Saida Eclusa",
        col["limite_indicador"]: "Limite Indicador",
        col["limite_janela"]: "Limite Janela",
        col["modal"]: "Modal",
    }
    if col["doca"]:
        rename[col["doca"]] = "Doca"
    if col["pedidos"]:
        rename[col["pedidos"]] = "Pedidos"
    if col["pecas"]:
        rename[col["pecas"]] = "Peças"
    if col["m3"]:
        rename[col["m3"]] = "M³"

    base = base.rename(columns=rename)
    base["Modal"] = base["Modal"].apply(clean_modal)
    # Se Transportador vier vazio/nulo, exibe "-" no painel e filtros
    base["Transportador"] = base["Transportador"].apply(lambda x: "-" if empty_value(x) else str(x).strip())

    return base, []


# =========================
# REGRAS OPERACIONAIS
# =========================
def aplicar_regras(df: pd.DataFrame, agora: datetime, alerta_chegada_min: int = 60, alerta_indicador_min: int = 30, alerta_janela_min: int = 30) -> pd.DataFrame:
    base = df.copy()

    def status_chegada(row):
        janela = row["Janela"]
        chegada = row["Chegada Portaria"]
        if pd.isna(janela):
            return "-"
        if not pd.isna(chegada):
            return "ATRASO CHEGADA" if chegada > janela else "NO PRAZO"
        minutos = (janela - agora).total_seconds() / 60
        if minutos < 0:
            return "ATRASO CHEGADA"
        if minutos <= alerta_chegada_min:
            return "PONTO DE ATENÇÃO"
        return "AGUARDANDO"

    def status_saida(row, limite_col, alerta_min):
        limite = row[limite_col]
        saida = row["Saida Eclusa"]
        if pd.isna(limite):
            return "-"
        if not pd.isna(saida):
            return "ATRASO SAÍDA" if saida > limite else "NO PRAZO"
        minutos = (limite - agora).total_seconds() / 60
        if minutos < 0:
            return "ATRASO SAÍDA"
        if minutos <= alerta_min:
            return "PONTO DE ATENÇÃO"
        return "AGUARDANDO"

    base["Status Chegada Monitor"] = base.apply(status_chegada, axis=1)
    base["Status Saída Indicador Monitor"] = base.apply(lambda r: status_saida(r, "Limite Indicador", alerta_indicador_min), axis=1)
    base["Status Saída Janela Monitor"] = base.apply(lambda r: status_saida(r, "Limite Janela", alerta_janela_min), axis=1)

    # Minutos para cards ativos
    base["Minutos Chegada"] = (base["Janela"] - agora).dt.total_seconds() / 60
    base["Minutos Indicador"] = (base["Limite Indicador"] - agora).dt.total_seconds() / 60
    base["Minutos Janela"] = (base["Limite Janela"] - agora).dt.total_seconds() / 60

    # Atraso real para análises históricas/detalhadas:
    # quando já existe Saida Eclusa, calcula Saida Eclusa - Limite;
    # quando ainda não saiu, usa Agora - Limite para manter o monitoramento ativo.
    base["Atraso Indicador Real"] = (base["Saida Eclusa"] - base["Limite Indicador"]).dt.total_seconds() / 60
    pend_ind = base["Atraso Indicador Real"].isna()
    base.loc[pend_ind, "Atraso Indicador Real"] = (agora - base.loc[pend_ind, "Limite Indicador"]).dt.total_seconds() / 60

    base["Atraso Janela Real"] = (base["Saida Eclusa"] - base["Limite Janela"]).dt.total_seconds() / 60
    pend_jan = base["Atraso Janela Real"].isna()
    base.loc[pend_jan, "Atraso Janela Real"] = (agora - base.loc[pend_jan, "Limite Janela"]).dt.total_seconds() / 60

    return base


def tipo_periodo(janela, hoje):
    if pd.isna(janela):
        return ""
    data = pd.to_datetime(janela).date()
    if data == hoje + timedelta(days=1):
        return "AMANHÃ"
    if data == hoje:
        return "HOJE"
    if data < hoje:
        return "DIAS ANTERIORES"
    return "FUTURO"


def filtrar_periodo_operacional(df: pd.DataFrame, hoje) -> pd.DataFrame:
    # Exibe dias anteriores + hoje + amanhã
    limite = hoje + timedelta(days=1)
    return df[df["Janela"].dt.date <= limite].copy()


def itens_card(df, status_col, minutos_col, saida_ou_chegada_col, limite_minutos_atencao):
    # Card superior: só mostra itens ativos, ou seja, sem chegada/saída preenchida.
    ativos = df[df[saida_ou_chegada_col].isna()].copy()
    ativos = ativos[
        ativos[status_col].isin(["ATRASO CHEGADA", "ATRASO SAÍDA", "PONTO DE ATENÇÃO"])
    ].copy()
    ativos["Prioridade"] = ativos[status_col].map({"ATRASO CHEGADA": 0, "ATRASO SAÍDA": 0, "PONTO DE ATENÇÃO": 1}).fillna(2)
    ativos["Min_abs"] = ativos[minutos_col].abs()
    return ativos.sort_values(["Prioridade", "Min_abs"], ascending=[True, False])


# =========================
# HTML / CSS
# =========================
def get_theme_css(theme: str) -> str:
    dark = theme == "Escuro"
    if dark:
        return """
        <style>
        :root{
            --bg:#07111f; --card:#0d1b2f; --card2:#10243d; --text:#eef5ff; --muted:#a7b7cf;
            --border:#24466f; --blue:#00A5FF; --blue2:#0076CE; --soft:#132b47;
            --shadow:0 8px 22px rgba(0,0,0,.35);
        }
        .stApp {background: var(--bg); color: var(--text);}
        div[data-testid="stToolbar"], header {visibility:hidden; height:0%;}
        .block-container{padding: 1.0rem 1.2rem 0.7rem 1.2rem; max-width: 100%;}
        </style>
        """
    return """
    <style>
    :root{
        --bg:#f6fbff; --card:#ffffff; --card2:#f9fcff; --text:#09203f; --muted:#4c6688;
        --border:#b7d7ff; --blue:#0086ff; --blue2:#0054a6; --soft:#edf6ff;
        --shadow:0 6px 18px rgba(0,81,171,.09);
    }
    .stApp {background: var(--bg); color: var(--text);}
    div[data-testid="stToolbar"], header {visibility:hidden; height:0%;}
    .block-container{padding: 1.0rem 1.2rem 0.7rem 1.2rem; max-width: 100%;}
    </style>
    """


def base_css() -> str:
    return """
    <style>
    .block-container{padding-top:.1rem !important; padding-bottom:.4rem !important; max-width:100% !important;}
    header[data-testid="stHeader"]{height:0rem;}
    div[data-testid="stVerticalBlock"]{gap:.28rem;}
    .topbar{display:flex; justify-content:space-between; align-items:flex-start; gap:12px; margin-bottom:6px;}
    .brand{display:flex; align-items:center; gap:16px;}
    .logo{font-size:34px; font-weight:900; color:var(--blue); letter-spacing:-1px; line-height:1;}
    .logo:after{content:""; display:block; width:100%; height:5px; margin-top:4px; border-radius:99px; background:linear-gradient(90deg,#ff2fa0,#ff8b00,#ffdc00,#00c853,#0086ff);}
    .title{font-size:30px; font-weight:900; color:var(--blue2); letter-spacing:.5px; line-height:1.05;}
    .subtitle{font-size:16px; color:var(--muted); margin-top:3px;}
    .clock{background:var(--card); border:1px solid var(--border); border-radius:14px; padding:8px 18px; min-width:178px; text-align:center; box-shadow:var(--shadow); margin-top:-18px;}
    .clock .time{font-size:26px; font-weight:900; color:var(--blue2);}
    .clock .date{font-size:14px; color:var(--blue); font-weight:800;}

    .filters{display:grid; grid-template-columns: 1.1fr 1.3fr 1fr 1fr 1fr 1.3fr .95fr .95fr; gap:10px; margin:8px 0 10px 0;}
    .filter-card{background:var(--card); border:1px solid var(--border); border-radius:12px; padding:9px 12px; box-shadow:var(--shadow); min-height:54px;}
    .filter-label{font-size:12px; color:var(--blue2); font-weight:900; margin-bottom:3px;}
    .filter-value{font-size:15px; color:var(--text); font-weight:700; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}

    .info{display:flex; justify-content:space-between; align-items:center; gap:10px; background:var(--soft); border:1px solid var(--border); border-radius:12px; padding:7px 14px; margin:6px 0 8px 0; color:var(--text); font-size:14px;}
    .refresh-pill{background:var(--card); border:1px solid var(--border); border-radius:10px; padding:6px 10px; font-weight:800; color:var(--blue2); white-space:nowrap;}

    .cards{display:grid; grid-template-columns: 1fr 1fr 1fr; gap:10px; margin:5px 0 6px 0; position:relative; z-index:5;}
    .alert-card{background:var(--card); border:1px solid var(--border); border-radius:14px; padding:6px 9px; box-shadow:var(--shadow); overflow:hidden;}
    .alert-card.chegada{border-top:5px solid #ffb000;}
    .alert-card.indicador{border-top:5px solid #f04438;}
    .alert-card.janela{border-top:5px solid #ff7a00;}
    .card-head{display:flex; align-items:center; justify-content:space-between; gap:6px; margin-bottom:4px;}
    .card-title{font-size:13px; font-weight:950; display:flex; align-items:center; gap:5px; white-space:nowrap; min-width:0; overflow:hidden; text-overflow:ellipsis; flex:1 1 auto;}
    .chegada .card-title{color:#f2a100}.indicador .card-title{color:#ef233c}.janela .card-title{color:#ff7a00}
    .metrics-inline{display:flex; align-items:center; justify-content:flex-end; gap:4px; flex:0 0 auto; flex-wrap:nowrap;}
    .metric-mini{background:var(--card2); border:1px solid var(--border); border-radius:10px; padding:6px 8px; min-width:66px; text-align:center; font-size:10.5px; font-weight:850; color:var(--muted); white-space:nowrap;}
    .metric-mini:last-child{min-width:82px;}
    .metric-mini b{font-size:14px; color:var(--text); margin-left:3px;}
    .mini-table{width:100%; border-collapse:collapse; font-size:12px;}
    .mini-scroll{max-height:172px; overflow-y:auto; padding-right:2px;}
    .mini-scroll::-webkit-scrollbar,.table-card::-webkit-scrollbar{width:10px;height:10px;}
    .mini-scroll::-webkit-scrollbar-thumb,.table-card::-webkit-scrollbar-thumb{background:#9ec7ff;border-radius:99px;}
    .mini-scroll::-webkit-scrollbar-track,.table-card::-webkit-scrollbar-track{background:rgba(128,128,128,.12);border-radius:99px;}
    .mini-table th{font-size:10px; color:var(--blue2); text-align:left; padding:4px 6px; border-bottom:1px solid var(--border);}
    .mini-table td{padding:4px 6px; border-bottom:1px solid rgba(128,128,128,.16); color:var(--text); font-weight:650;}
    .mini-table tr:last-child td{border-bottom:none;}

    .table-card{background:var(--card); border:1px solid var(--border); border-radius:14px; box-shadow:var(--shadow); padding:0; overflow:auto; height:72vh; position:relative; z-index:1; contain:paint; margin-top:5px; margin-bottom:22px; clear:both;}
    div[data-testid="stExpander"]{margin-top:14px !important; clear:both !important;}
    table.main-table{width:100%; border-collapse:separate; border-spacing:0; font-size:13px;}
    .main-table thead{position:static; background:var(--soft);}
    .main-table th{position:sticky; top:0; z-index:20; background:var(--soft) !important; color:var(--blue2); padding:8px 7px; text-align:left; font-size:12px; font-weight:950; border-bottom:1px solid var(--border); box-shadow:0 3px 7px rgba(0,0,0,.08);}
    .main-table td{padding:5px 7px; border-bottom:1px solid rgba(128,128,128,.14); color:var(--text); font-weight:650; white-space:nowrap;}
    .main-table th:nth-child(n+4), .main-table td:nth-child(n+4){text-align:center;}
    .main-table td:nth-child(1), .main-table th:nth-child(1), .main-table td:nth-child(2), .main-table th:nth-child(2), .main-table td:nth-child(3), .main-table th:nth-child(3){text-align:left;}
    .main-table tr:nth-child(even) td{background:rgba(128,128,128,.035);}
    .main-table tbody tr{background:var(--card);}
    .dot{display:inline-block; width:11px; height:11px; border-radius:50%; margin-right:5px;}
    .dot.amanha{background:#2478ff}.dot.hoje{background:#16a34a}.dot.anterior{background:#9ca3af}.dot.futuro{background:#7c3aed}

    .badge{display:inline-block; padding:3px 8px; border-radius:8px; font-size:11px; font-weight:950; line-height:1; border:1px solid transparent;}
    .modal-venda{color:#e11d48; background:#ffe8ee; border-color:#ffb6c7;}
    .modal-transferencia{color:#0969da; background:#e7f1ff; border-color:#aed0ff;}
    .modal-aba{color:#048a45; background:#e5f7ed; border-color:#a8e3c0;}
    .st-ok{color:#039855; background:#e8fff2; border-color:#9ee7bd;}
    .st-aguardando{color:#b77900; background:#fff7df; border-color:#ffd970;}
    .st-atencao{color:#e57a00; background:#fff4df; border-color:#ffbd5b;}
    .st-chegada{color:#c55300; background:#fff0e5; border-color:#ffb47a;}
    .st-atraso{color:#c1121f; background:#ffe7e9; border-color:#ff9aa2;}

    .after-main-table-spacer{height:8px; clear:both;}
    .daily-impact{background:var(--card); border:1px solid var(--border); border-radius:14px; box-shadow:var(--shadow); padding:9px 12px; margin-top:9px; clear:both;}
    .daily-impact-title{font-size:15px; font-weight:950; color:var(--blue2); margin-bottom:2px;}
    .daily-impact-subtitle{font-size:12px; color:var(--muted); font-weight:700; margin-bottom:6px;}
    .daily-impact-scroll{overflow-x:auto; width:100%;}
    .daily-impact-table{width:100%; border-collapse:separate; border-spacing:0; font-size:11.5px; min-width:900px;}
    .daily-impact-table th{background:var(--soft); color:var(--blue2); padding:5px; text-align:center; font-weight:950; border-bottom:1px solid var(--border);}
    .daily-impact-table td{padding:6px 5px; border-bottom:1px solid rgba(128,128,128,.14); text-align:center; font-weight:900; color:#0b2341;}
    .daily-impact-table .data-cell{background:var(--soft); color:var(--blue2); font-weight:950; text-align:center; min-width:62px; font-size:11px;}
    .daily-impact-table .total-cell{background:#e7f1ff !important; color:var(--blue2) !important; font-weight:950; border-left:1px solid var(--border);}

    .status-kpis{display:grid; grid-template-columns:repeat(5,1fr); gap:8px; margin:8px 0 8px 0;}
    .status-kpi{border:1px solid var(--border); border-radius:10px; padding:8px 10px; background:rgba(255,255,255,.55); display:flex; align-items:center; justify-content:space-between; font-weight:900;}
    .status-kpi span{font-size:11px; color:var(--muted);}
    .status-kpi b{font-size:18px; color:var(--blue2);}
    .status-kpi.aguardando b{color:#d97706;}
    .status-kpi.carregando b{color:#0b74de;}
    .status-kpi.liberadas b{color:#00a650;}
    .status-kpi.pct b{color:#00a650;}
    .status-day-table .pct-cell{font-weight:950; color:#00a650; background:#effdf5;}
    .heat-0{background:#e9f8ef !important;}
    .heat-1{background:#f2f8d8 !important;}
    .heat-2{background:#fff3b0 !important;}
    .heat-3{background:#ffd27a !important;}
    .heat-4{background:#ff8b8b !important; color:#7f1d1d !important;}
    .heat-legend{display:flex; align-items:center; gap:8px; margin-top:7px; font-size:12px; color:var(--muted); font-weight:750; flex-wrap:wrap;}
    .heat-dot{display:inline-block; width:20px; height:12px; border-radius:4px; border:1px solid rgba(0,0,0,.08);}
    .heat-desc{margin-left:8px;}

    .indicator-summary{background:var(--card); border:1px solid var(--border); border-radius:14px; box-shadow:var(--shadow); padding:9px 12px; margin-top:9px;}
    .indicator-summary-title{font-size:15px; font-weight:950; color:var(--blue2); margin-bottom:2px;}
    .indicator-summary-subtitle{font-size:12px; color:var(--muted); font-weight:700; margin-bottom:6px;}
    .indicator-summary-table{width:100%; border-collapse:separate; border-spacing:0; font-size:12.5px;}
    .indicator-summary-table th{background:var(--soft); color:var(--blue2); padding:7px; text-align:center; font-weight:950; border-bottom:1px solid var(--border);}
    .indicator-summary-table td{padding:7px; border-bottom:1px solid rgba(128,128,128,.14); text-align:center; font-weight:850; color:var(--text);}
    .indicator-summary-table .data-cell{background:var(--soft); color:var(--blue2); font-weight:950; min-width:70px;}
    .indicator-summary-table .num-strong{color:var(--blue2); font-weight:950;}
    .indicator-summary-table .m3-cell{color:#7c3aed; font-weight:950;}
    .indicator-summary-table .ped-cell{color:#039855; font-weight:950;}
    .indicator-summary-table .atraso-cell{color:#ef4444; font-weight:950;}

    .legends{display:grid; grid-template-columns: .9fr 1.6fr .8fr; gap:8px; margin-top:7px;}
    .legend-card{background:var(--card); border:1px solid var(--border); border-radius:12px; padding:6px 10px; box-shadow:var(--shadow); font-size:11px;}
    .legend-title{color:var(--blue2); font-weight:950; margin-right:10px;}
    .legend-inline{display:flex; align-items:center; flex-wrap:wrap; gap:8px;}
    .footer{display:flex; justify-content:space-between; align-items:center; gap:10px; background:var(--card); border:1px solid var(--border); border-radius:12px; padding:7px 12px; margin-top:7px; box-shadow:var(--shadow); font-size:13px; color:var(--muted);}
    .footer b{color:var(--blue2);}

    /* Labels e menus dos filtros no tema claro/escuro */
    div[data-testid="stSelectbox"] label,
    div[data-testid="stMultiSelect"] label,
    div[data-testid="stRadio"] label{
        color:var(--blue2) !important;
        font-weight:900 !important;
    }
    div[data-testid="stSelectbox"], div[data-testid="stMultiSelect"]{margin-bottom:0 !important;}
    div[data-testid="stMarkdownContainer"] p{margin-bottom:.2rem;}

    /* Controles do Streamlit: força contraste legível nos filtros */
    div[data-baseweb="select"] > div{
        background:#2b2d36 !important;
        border:1px solid #2b2d36 !important;
        color:#ffffff !important;
    }
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] input,
    div[data-baseweb="select"] div,
    div[data-baseweb="select"] svg{
        color:#ffffff !important;
        -webkit-text-fill-color:#ffffff !important;
        fill:#ffffff !important;
    }
    div[data-baseweb="select"] input::placeholder{
        color:#ffffff !important;
        opacity:1 !important;
        -webkit-text-fill-color:#ffffff !important;
    }
    div[data-baseweb="popover"],
    div[data-baseweb="popover"] * ,
    div[role="listbox"],
    div[role="listbox"] *,
    ul[role="listbox"],
    ul[role="listbox"] *,
    div[role="option"],
    div[role="option"] *{
        background:#1f2430 !important;
        color:#ffffff !important;
        -webkit-text-fill-color:#ffffff !important;
    }
    div[role="option"]:hover, div[role="option"]:hover *{
        background:#334155 !important;
        color:#ffffff !important;
        -webkit-text-fill-color:#ffffff !important;
    }

    /* Campo para colar o link do Google Sheets: legível nos dois temas */
    div[data-testid="stTextInput"] input{
        background:var(--card) !important;
        color:var(--text) !important;
        -webkit-text-fill-color:var(--text) !important;
        border:1px solid var(--border) !important;
        border-radius:10px !important;
        font-weight:700 !important;
    }
    div[data-testid="stTextInput"] input::placeholder{
        color:var(--muted) !important;
        -webkit-text-fill-color:var(--muted) !important;
        opacity:1 !important;
    }
    div[data-testid="stTextInput"] label, div[data-testid="stTextInput"] label *{
        color:var(--blue2) !important;
        -webkit-text-fill-color:var(--blue2) !important;
        font-weight:900 !important;
    }


    /* Expander e campos numéricos: contraste correto no tema claro e escuro */
    div[data-testid="stExpander"] details{
        background:var(--card) !important;
        border:1px solid var(--border) !important;
        border-radius:10px !important;
        overflow:hidden !important;
    }
    div[data-testid="stExpander"] summary{
        background:var(--card) !important;
        color:var(--text) !important;
        font-weight:900 !important;
    }
    div[data-testid="stExpander"] summary *{
        color:var(--text) !important;
        -webkit-text-fill-color:var(--text) !important;
        fill:var(--text) !important;
    }
    div[data-testid="stExpander"] label,
    div[data-testid="stExpander"] label *,
    div[data-testid="stNumberInput"] label,
    div[data-testid="stNumberInput"] label *{
        color:var(--text) !important;
        -webkit-text-fill-color:var(--text) !important;
        font-weight:850 !important;
    }
    div[data-testid="stNumberInput"] input{
        background:#2b2d36 !important;
        color:#ffffff !important;
        -webkit-text-fill-color:#ffffff !important;
        border:1px solid #2b2d36 !important;
        font-weight:800 !important;
    }
    div[data-testid="stNumberInput"] button{
        background:#2b2d36 !important;
        color:#ffffff !important;
        border-color:#2b2d36 !important;
    }
    div[data-testid="stNumberInput"] button *{
        color:#ffffff !important;
        fill:#ffffff !important;
    }


    /* Campo de link / inputs: força contraste em tema claro e escuro */
    div[data-baseweb="input"] > div,
    div[data-testid="stTextInput"] div[data-baseweb="input"] > div{
        background:var(--card) !important;
        border:1px solid var(--border) !important;
    }
    div[data-baseweb="input"] input,
    div[data-testid="stTextInput"] input{
        background:var(--card) !important;
        color:var(--text) !important;
        -webkit-text-fill-color:var(--text) !important;
        caret-color:var(--blue2) !important;
        font-weight:800 !important;
    }
    div[data-baseweb="input"] input::placeholder,
    div[data-testid="stTextInput"] input::placeholder{
        color:var(--muted) !important;
        -webkit-text-fill-color:var(--muted) !important;
        opacity:1 !important;
    }
    .daily-impact-table .total-janela{background:#fff3b0 !important; color:#7a4b00 !important; font-weight:950; border-left:1px solid var(--border);}
    .daily-impact-table .total-indicador{background:#e7f1ff !important; color:var(--blue2) !important; font-weight:950;}


    @media (max-width: 1500px){
        .card-title{font-size:12px;}
        .metric-mini{min-width:58px; padding:5px 6px; font-size:10px;}
        .metric-mini:last-child{min-width:76px;}
        .metric-mini b{font-size:13px;}
        .badge{font-size:10px; padding:3px 7px;}
    }

    @media (min-width: 1800px){
        .title{font-size:38px}.subtitle{font-size:20px}.filter-value{font-size:18px}.card-title{font-size:16px}
        table.main-table{font-size:16px}.main-table th{font-size:14px}.mini-table{font-size:15px}.badge{font-size:13px;padding:5px 10px;}
    }
    </style>
    """


def badge_modal(modal: str) -> str:
    m = clean_modal(modal)
    cls = {"VENDA": "modal-venda", "TRANSFERÊNCIA": "modal-transferencia", "ABA": "modal-aba"}.get(m, "")
    label = "TRANSF" if m == "TRANSFERÊNCIA" else m
    return f'<span class="badge {cls}">{label}</span>'


def badge_status(status: str, tempo: str | None = None) -> str:
    s = norm_txt(status)
    label = status if tempo is None else f"{status} {tempo}"
    if "PONTO" in s:
        cls = "st-atencao"
    elif "ATRASO CHEG" in s:
        cls = "st-chegada"
    elif "ATRASO" in s:
        cls = "st-atraso"
    elif "AGUARD" in s:
        cls = "st-aguardando"
    elif "PRAZO" in s:
        cls = "st-ok"
    else:
        cls = ""
    return f'<span class="badge {cls}">{label}</span>'


def render_filter(label, value, icon=""):
    return textwrap.dedent(f"""
    <div class="filter-card">
        <div class="filter-label">{icon} {label}</div>
        <div class="filter-value">{value}</div>
    </div>
    """).strip()


def html_join(parts):
    """Junta pedaços HTML sem indentação inicial para evitar que o Streamlit renderize como bloco de código."""
    return "".join(str(x) for x in parts)


def render_card(title, icon, df_card, tipo, minutos_col, status_col, color_class, limite_col=None):
    total = len(df_card)
    venda = int((df_card["Modal"] == "VENDA").sum()) if total else 0
    if total:
        atrasos = df_card[df_card[minutos_col] < 0][minutos_col].abs()
        maior = format_tempo(atrasos.max()) if len(atrasos) else format_tempo(df_card[minutos_col].abs().min())
    else:
        maior = "-"

    body_rows = []
    for _, r in df_card.iterrows():
        minutos = r[minutos_col]
        tempo = format_tempo(minutos)
        body_rows.append(
            "<tr>"
            f"<td>{r['ID Carga']}</td>"
            f"<td>{str(r['Plano de Transporte'])[:24]}</td>"
            f"<td>{badge_modal(r['Modal'])}</td>"
            f"<td>{badge_status(r[status_col], tempo)}</td>"
            "</tr>"
        )

    if not body_rows:
        body_rows.append("<tr><td colspan='4' style='text-align:center;color:var(--muted);padding:18px;'>Sem cargas em atenção neste momento</td></tr>")

    return html_join([
        f"<div class='alert-card {color_class}'>",
        "<div class='card-head'>",
        f"<div class='card-title'>{icon} {title}</div>",
        "<div class='metrics-inline'>",
        f"<div class='metric-mini'>Total <b>{total}</b></div>",
        f"<div class='metric-mini'>Venda <b>{venda}</b></div>",
        f"<div class='metric-mini'>Maior <b>{maior}</b></div>",
        "</div></div>",
        "<div class='mini-scroll'><table class='mini-table'>",
        "<thead><tr><th>ID</th><th>PLANO</th><th>MODAL</th><th>STATUS</th></tr></thead>",
        "<tbody>",
        *body_rows,
        "</tbody></table></div></div>",
    ])

def render_main_table(df: pd.DataFrame) -> str:
    body_rows = []
    for _, r in df.iterrows():
        periodo = r.get("Período", "")
        dot_cls = "amanha" if periodo == "AMANHÃ" else "hoje" if periodo == "HOJE" else "anterior" if periodo == "DIAS ANTERIORES" else "futuro"
        body_rows.append(
            "<tr>"
            f"<td><span class='dot {dot_cls}'></span>{r['ID Carga']}</td>"
            f"<td>{str(r['Plano de Transporte'])[:28]}</td>"
            f"<td>{str(r['Transportador'])[:18]}</td>"
            f"<td>{r['Turno'] if not empty_value(r['Turno']) else '-'}</td>"
            f"<td>{format_dt(r['Janela'])}</td>"
            f"<td>{badge_modal(r['Modal'])}</td>"
            f"<td>{format_dt(r['Chegada Portaria'])}</td>"
            f"<td>{format_dt(r['Limite Indicador'])}</td>"
            f"<td>{format_dt(r['Limite Janela'])}</td>"
            f"<td>{badge_status(r['Status Chegada Monitor'])}</td>"
            f"<td>{badge_status(r['Status Saída Indicador Monitor'])}</td>"
            f"<td>{badge_status(r['Status Saída Janela Monitor'])}</td>"
            "</tr>"
        )

    return html_join([
        "<div class='table-card'><table class='main-table'>",
        "<thead><tr>",
        "<th>ID CARGA</th><th>PLANO DE TRANSPORTE</th><th>TRANSPORTADOR</th><th>TURNO</th><th>JANELA ↓</th><th>MODAL</th>",
        "<th>CHEGADA PORTARIA</th><th>LIMITE INDICADOR</th><th>LIMITE JANELA</th><th>STATUS CHEGADA</th><th>STATUS SAÍDA INDICADOR</th><th>STATUS SAÍDA JANELA</th>",
        "</tr></thead><tbody>",
        *body_rows,
        "</tbody></table></div>",
    ])


def heat_class(v: int) -> str:
    if v <= 0:
        return "heat-0"
    if v == 1:
        return "heat-1"
    if v == 2:
        return "heat-2"
    if v == 3:
        return "heat-3"
    return "heat-4"



def render_status_cargas_dia(df: pd.DataFrame, modal_filtro: str = "Todos", turno_filtro: str = "Todos", datas_filtro=None) -> str:
    """Painel independente: status operacional por dia, com filtros próprios de Modal e Turno."""
    base_status = df.copy()
    base_status = base_status[base_status["Janela"].notna()].copy()
    if datas_filtro:
        datas_norm = set(pd.to_datetime(datas_filtro).date)
        base_status = base_status[base_status["Janela"].dt.date.isin(datas_norm)].copy()

    if modal_filtro != "Todos":
        base_status = base_status[base_status["Modal"].astype(str).eq(modal_filtro)]
    if turno_filtro != "Todos":
        base_status = base_status[base_status["Turno"].astype(str).eq(turno_filtro)]

    if base_status.empty:
        return html_join([
            "<div class='daily-impact'>",
            "<div class='daily-impact-title'>📦 STATUS DAS CARGAS POR DIA</div>",
            "<div class='daily-impact-subtitle'>Aguardando veículo, carregando e liberadas por Data Janela.</div>",
            "<div style='color:var(--muted);padding:14px;text-align:center;font-weight:750;'>Sem cargas para os filtros selecionados</div>",
            "</div>",
        ])

    base_status["Data Janela"] = base_status["Janela"].dt.date
    base_status["Aguardando Veículo"] = base_status["Chegada Portaria"].isna() & base_status["Saida Eclusa"].isna()
    base_status["Carregando"] = base_status["Chegada Portaria"].notna() & base_status["Saida Eclusa"].isna()
    base_status["Liberadas"] = base_status["Saida Eclusa"].notna()

    resumo = (
        base_status.groupby("Data Janela", dropna=True)
        .agg(
            AguardandoVeiculo=("Aguardando Veículo", "sum"),
            Carregando=("Carregando", "sum"),
            Liberadas=("Liberadas", "sum"),
            Total=("ID Carga", "count"),
        )
        .reset_index()
        .sort_values("Data Janela", ascending=False)
    )
    resumo["PercConcluido"] = resumo.apply(lambda r: (r["Liberadas"] / r["Total"] * 100) if r["Total"] else 0, axis=1)

    total_aguardando = int(resumo["AguardandoVeiculo"].sum())
    total_carregando = int(resumo["Carregando"].sum())
    total_liberadas = int(resumo["Liberadas"].sum())
    total_cargas = int(resumo["Total"].sum())
    pct_total = (total_liberadas / total_cargas * 100) if total_cargas else 0

    rows = []
    for _, r in resumo.iterrows():
        data_txt = pd.Timestamp(r["Data Janela"]).strftime("%d/%m")
        agu = int(r["AguardandoVeiculo"])
        carreg = int(r["Carregando"])
        lib = int(r["Liberadas"])
        total = int(r["Total"])
        pct = float(r["PercConcluido"])
        rows.append(
            "<tr>"
            f"<td class='data-cell'>{data_txt}</td>"
            f"<td class='heat-2'>{agu}</td>"
            f"<td class='heat-1'>{carreg}</td>"
            f"<td class='heat-0'>{lib}</td>"
            f"<td class='total-indicador'>{total}</td>"
            f"<td class='pct-cell'>{pct:.1f}%</td>"
            "</tr>"
        )

    rows.append(
        "<tr>"
        "<td class='data-cell'>TOTAL</td>"
        f"<td class='heat-2'>{total_aguardando}</td>"
        f"<td class='heat-1'>{total_carregando}</td>"
        f"<td class='heat-0'>{total_liberadas}</td>"
        f"<td class='total-indicador'>{total_cargas}</td>"
        f"<td class='pct-cell'>{pct_total:.1f}%</td>"
        "</tr>"
    )

    return html_join([
        "<div class='daily-impact'>",
        "<div class='daily-impact-title'>📦 STATUS DAS CARGAS POR DIA</div>",
        "<div class='daily-impact-subtitle'>Painel independente. Aguardando veículo = sem chegada e sem saída; Carregando = com chegada e sem saída; Liberadas = com saída eclusa.</div>",
        "<div class='status-kpis'>",
        f"<div class='status-kpi aguardando'><span>Aguardando veículo</span><b>{total_aguardando}</b></div>",
        f"<div class='status-kpi carregando'><span>Carregando</span><b>{total_carregando}</b></div>",
        f"<div class='status-kpi liberadas'><span>Liberadas</span><b>{total_liberadas}</b></div>",
        f"<div class='status-kpi total'><span>Total</span><b>{total_cargas}</b></div>",
        f"<div class='status-kpi pct'><span>Concluído</span><b>{pct_total:.1f}%</b></div>",
        "</div>",
        "<div class='daily-impact-scroll'>",
        "<table class='daily-impact-table status-day-table'>",
        "<thead><tr><th>DATA</th><th>AGUARDANDO VEÍCULO</th><th>CARREGANDO</th><th>LIBERADAS</th><th>TOTAL</th><th>% CONCLUÍDO</th></tr></thead><tbody>",
        *rows,
        "</tbody></table>",
        "</div>",
        "</div>",
    ])

def render_venda_turno_impact(df: pd.DataFrame) -> str:
    """Matriz: somente VENDA, linhas por data da Janela e colunas por Turno x Janela/Indicador."""
    venda = df[df["Modal"].eq("VENDA")].copy()
    venda = venda[venda["Janela"].notna()].copy()

    if venda.empty:
        return html_join([
            "<div class='daily-impact'>",
            "<div class='daily-impact-title'>📊 VENDA | ATRASOS POR DIA E TURNO</div>",
            "<div style='color:var(--muted);padding:14px;text-align:center;font-weight:750;'>Sem cargas de venda no período filtrado</div>",
            "</div>",
        ])

    venda["Data Janela"] = venda["Janela"].dt.date
    venda["Turno Texto"] = venda["Turno"].astype(str).fillna("-")

    datas = sorted(venda["Data Janela"].dropna().unique(), reverse=True)
    turnos = sorted(venda["Turno Texto"].dropna().unique().tolist())[:6]

    rows = []
    for data in datas:
        cells = [f"<td class='data-cell'>{pd.Timestamp(data).strftime('%d/%m')}</td>"]
        total_janela = 0
        total_ind = 0
        for turno in turnos:
            m = (venda["Data Janela"].eq(data)) & (venda["Turno Texto"].eq(turno))
            subset = venda[m]
            qtd_janela = int(subset["Status Saída Janela Monitor"].eq("ATRASO SAÍDA").sum())
            qtd_ind = int(subset["Status Saída Indicador Monitor"].eq("ATRASO SAÍDA").sum())
            total_janela += qtd_janela
            total_ind += qtd_ind
            cells.append(f"<td class='{heat_class(qtd_janela)}'>{qtd_janela}</td>")
            cells.append(f"<td class='{heat_class(qtd_ind)}'>{qtd_ind}</td>")
        cells.append(f"<td class='total-janela'>{total_janela}</td>")
        cells.append(f"<td class='total-indicador'>{total_ind}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")

    turno_headers = "".join([f"<th colspan='2'>{t}</th>" for t in turnos]) + "<th rowspan='2'>TOTAL<br>JANELA</th><th rowspan='2'>TOTAL<br>INDICADOR</th>"
    sub_headers = "".join(["<th>Janela</th><th>Indicador</th>" for _ in turnos])

    return html_join([
        "<div class='daily-impact'>",
        "<div class='daily-impact-title'>📊 VENDA | ATRASOS POR DIA E TURNO</div>",
        "<div class='daily-impact-subtitle'>Contagem de cargas de venda que impactaram Saída Janela e Saída Indicador, agrupadas por Data Janela e Turno.</div>",
        "<div class='daily-impact-scroll'>",
        "<table class='daily-impact-table'>",
        "<thead>",
        f"<tr><th rowspan='2'>DATA</th>{turno_headers}</tr>",
        f"<tr>{sub_headers}</tr>",
        "</thead><tbody>",
        *rows,
        "</tbody></table>",
        "</div>",
        "<div class='heat-legend'>",
        "<span class='legend-title'>Escala</span>",
        "<span class='heat-dot heat-0'></span>0",
        "<span class='heat-dot heat-1'></span>1",
        "<span class='heat-dot heat-2'></span>2",
        "<span class='heat-dot heat-3'></span>3",
        "<span class='heat-dot heat-4'></span>4+",
        "<span class='heat-desc'>Somente cargas com Modal = VENDA. Totais separados por Janela e Indicador.</span>",
        "</div>",
        "</div>",
    ])


def render_venda_indicador_dia(df: pd.DataFrame) -> str:
    """Detalhe diário: somente VENDA que impactou Saída Indicador, por carga/plano, com M³ e pedidos."""
    venda = df[df["Modal"].eq("VENDA")].copy()
    venda = venda[venda["Janela"].notna()].copy()
    if venda.empty:
        return html_join([
            "<div class='indicator-summary'>",
            "<div class='indicator-summary-title'>📦 VENDA | IMPACTO SAÍDA INDICADOR POR DIA</div>",
            "<div style='color:var(--muted);padding:12px;text-align:center;font-weight:750;'>Sem cargas de venda no período filtrado</div>",
            "</div>",
        ])

    venda["Data Janela"] = venda["Janela"].dt.date
    impact = venda[venda["Status Saída Indicador Monitor"].eq("ATRASO SAÍDA")].copy()

    if impact.empty:
        return html_join([
            "<div class='indicator-summary'>",
            "<div class='indicator-summary-title'>📦 VENDA | IMPACTO SAÍDA INDICADOR POR DIA</div>",
            "<div class='indicator-summary-subtitle'>Cargas de venda que impactaram o indicador, detalhadas por plano de transporte.</div>",
            "<div style='color:var(--muted);padding:12px;text-align:center;font-weight:750;'>Sem impacto de saída indicador no período filtrado</div>",
            "</div>",
        ])

    # Garante que as colunas opcionais existam mesmo se a base não trouxer M³/Pedidos.
    if "M³" not in impact.columns:
        impact["M³"] = 0
    if "Pedidos" not in impact.columns:
        impact["Pedidos"] = 0

    detalhe = (
        impact.groupby(["Data Janela", "ID Carga", "Plano de Transporte"], dropna=True)
        .agg(
            LimiteIndicador=("Limite Indicador", first_valid),
            SaidaEclusa=("Saida Eclusa", first_valid),
            MinutosIndicador=("Atraso Indicador Real", first_valid),
            M3=("M³", "sum"),
            Pedidos=("Pedidos", "sum"),
        )
        .reset_index()
        .sort_values(["Data Janela", "ID Carga"], ascending=[False, True])
    )

    total_cargas = int(detalhe["ID Carga"].nunique())
    total_m3 = float(detalhe["M3"].sum())
    total_pedidos = int(round(float(detalhe["Pedidos"].sum())))

    rows = []
    last_date = None
    for _, r in detalhe.iterrows():
        data_txt = pd.Timestamp(r["Data Janela"]).strftime("%d/%m")
        id_carga = html.escape(str(r["ID Carga"]))
        plano = html.escape(str(r["Plano de Transporte"]))
        limite_ind = format_dt(r["LimiteIndicador"])
        saida_eclusa = format_dt(r["SaidaEclusa"])
        atraso = format_tempo(r["MinutosIndicador"])
        m3 = float(r["M3"])
        pedidos = int(round(float(r["Pedidos"])))
        date_cell = data_txt if data_txt != last_date else ""
        last_date = data_txt
        rows.append(
            "<tr>"
            f"<td class='data-cell'>{date_cell}</td>"
            f"<td class='num-strong'>{id_carga}</td>"
            f"<td style='text-align:left;font-weight:900;'>{plano}</td>"
            f"<td>{limite_ind}</td>"
            f"<td>{saida_eclusa}</td>"
            f"<td class='atraso-cell'>{atraso}</td>"
            f"<td class='m3-cell'>{m3:,.1f}</td>"
            f"<td class='ped-cell'>{pedidos}</td>"
            "</tr>"
        )

    rows.append(
        "<tr>"
        "<td class='data-cell'>TOTAL</td>"
        f"<td class='num-strong'>{total_cargas}</td>"
        "<td style='text-align:left;font-weight:950;'>Cargas de venda com impacto no indicador</td>"
        "<td>-</td>"
        "<td>-</td>"
        "<td>-</td>"
        f"<td class='m3-cell'>{total_m3:,.1f}</td>"
        f"<td class='ped-cell'>{total_pedidos}</td>"
        "</tr>"
    )

    return html_join([
        "<div class='indicator-summary'>",
        "<div class='indicator-summary-title'>📦 VENDA | IMPACTO SAÍDA INDICADOR POR DIA</div>",
        "<div class='indicator-summary-subtitle'>Somente cargas de venda com Status Saída Indicador = ATRASO SAÍDA. Mostra cada carga/plano que impactou, com limite indicador, saída eclusa, tempo de atraso, M³ e pedidos.</div>",
        "<table class='indicator-summary-table'>",
        "<thead><tr><th>DATA JANELA</th><th>ID CARGA</th><th>PLANO DE TRANSPORTE</th><th>LIMITE INDICADOR</th><th>SAÍDA ECLUSA</th><th>ATRASO</th><th>M³</th><th>PEDIDOS</th></tr></thead><tbody>",
        *rows,
        "</tbody></table>",
        "</div>",
    ])


# =========================
# APP
# =========================
# Atualização automática real
st.session_state.setdefault("auto_refresh_seconds_v29", 60)
AUTOREFRESH_SECONDS = int(st.session_state.get("auto_refresh_seconds_v29", 60))
st_autorefresh(interval=AUTOREFRESH_SECONDS * 1000, key="auto_refresh_saida_eclusa")

# CSS inicial para a tela de carga
st.markdown(get_theme_css(st.session_state.get("tema_select", "Claro")) + base_css(), unsafe_allow_html=True)

# Entrada de dados
# Depois que uma base é carregada, ela fica salva na sessão. Assim o painel começa no topo,
# sem o bloco de upload ocupando espaço na TV.
df_raw = None
# Link padrão fixo: ao abrir o app publicado, qualquer pessoa já acessa a base correta.
# Se algum dia a planilha mudar, basta alterar esse link no código ou usar o menu lateral para trocar a base.
if "excel_bytes" in st.session_state:
    df_raw = load_excel(st.session_state["excel_bytes"])
else:
    st.session_state.setdefault("csv_url", DEFAULT_GOOGLE_SHEETS_URL)
    try:
        df_raw = load_google_csv(google_sheets_to_csv_url(st.session_state["csv_url"]))
    except Exception as e:
        st.markdown("""
        <div class="topbar" style="margin-top:24px;">
            <div class="brand">
                <div class="logo">Magalu</div>
                <div>
                    <div class="title">PAINEL DE ACOMPANHAMENTO | SAÍDA ECLUSA</div>
                    <div class="subtitle">Não foi possível carregar o Google Sheets padrão</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.error(f"Erro ao carregar a base: {e}")
        with st.container(border=True):
            fonte = st.radio("Fonte de dados", ["Upload Excel", "Google Sheets CSV"], horizontal=True)
            if fonte == "Upload Excel":
                arquivo = st.file_uploader("Envie o arquivo Excel", type=["xlsx", "xls"])
                if arquivo is not None:
                    st.session_state["excel_bytes"] = arquivo.getvalue()
                    st.rerun()
            else:
                csv_url = st.text_input(
                    "Cole o link da aba do Google Sheets",
                    value=st.session_state.get("csv_url", DEFAULT_GOOGLE_SHEETS_URL),
                    placeholder="https://docs.google.com/spreadsheets/d/.../edit?gid=...",
                    help="Pode ser o link normal da aba. A planilha precisa estar compartilhada para leitura ou publicada na web."
                )
                st.caption("Dica: use o link da aba onde os títulos começam na linha 3.")
                if csv_url:
                    st.session_state["csv_url"] = csv_url
                    st.session_state.pop("excel_bytes", None)
                    st.rerun()
        st.stop()

base, faltantes = preparar_base(df_raw)
if faltantes:
    st.error(f"Não encontrei as colunas obrigatórias: {', '.join(faltantes)}. Confira se os títulos estão na linha 3.")
    st.stop()

agora = datetime.now(ZoneInfo(TIMEZONE)).replace(tzinfo=None)
hoje = agora.date()

# Tempos configuráveis de ponto de atenção. A alteração no painel atualiza na próxima execução automática/manual.
st.session_state.setdefault("alerta_chegada_min", 60)
st.session_state.setdefault("alerta_indicador_min", 30)
st.session_state.setdefault("alerta_janela_min", 30)
alerta_chegada_min = int(st.session_state.get("alerta_chegada_min", 60))
alerta_indicador_min = int(st.session_state.get("alerta_indicador_min", 30))
alerta_janela_min = int(st.session_state.get("alerta_janela_min", 30))

base = aplicar_regras(base, agora, alerta_chegada_min, alerta_indicador_min, alerta_janela_min)
base["Período"] = base["Janela"].apply(lambda x: tipo_periodo(x, hoje))
base = filtrar_periodo_operacional(base, hoje)
base = base.sort_values("Janela", ascending=False)

# CSS e tema
# Usa o valor salvo na sessão antes de desenhar a tela, evitando que HTML apareça como texto.
tema_atual = st.session_state.get("tema_select", "Claro")
st.markdown(get_theme_css(tema_atual) + base_css(), unsafe_allow_html=True)

# Header visual
st.markdown(f"""
<div class="topbar">
    <div class="brand">
        <div class="logo">Magalu</div>
        <div>
            <div class="title">PAINEL DE ACOMPANHAMENTO | SAÍDA ECLUSA</div>
            <div class="subtitle">Monitoramento operacional em tempo real</div>
        </div>
    </div>
    <div class="clock">
        <div class="time">{agora.strftime('%H:%M:%S')}</div>
        <div class="date">{agora.strftime('%d/%m/%Y')}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# Filtros superiores reais do Streamlit
data_options = sorted(base["Janela"].dropna().dt.date.unique().tolist(), reverse=True)
data_labels = [d.strftime("%d/%m/%Y") for d in data_options]

ctrl1, ctrl2, ctrl3, ctrl4, ctrl5, ctrl6, ctrl7, ctrl8, ctrl9, ctrl10 = st.columns([.85, 1.05, 1.15, .68, .82, .66, 1.12, .72, .86, .62])
periodo_options = ["Todos até amanhã", "Até 2h para frente", "Hoje", "Amanhã", "Dias anteriores"]
transp_options = ["Todas"] + sorted(base["Transportador"].dropna().astype(str).unique().tolist())
turno_options = ["Todos"] + sorted(base["Turno"].dropna().astype(str).unique().tolist())
status_options = ["Todos", "NO PRAZO", "AGUARDANDO", "PONTO DE ATENÇÃO", "ATRASO CHEGADA", "ATRASO SAÍDA"]
doca_options = ["Todas"]
if "Doca" in base.columns:
    doca_options += sorted(base["Doca"].dropna().astype(str).unique().tolist())
plano_options = ["Todos"] + sorted(base["Plano de Transporte"].dropna().astype(str).unique().tolist())
modal_options = ["Todos"] + sorted(base["Modal"].dropna().astype(str).unique().tolist())
saida_options = ["Todas", "Somente pendentes", "Somente com saída"]
tema_options = ["Claro", "Escuro"]

with ctrl1:
    periodo_sel = st.selectbox("Período", periodo_options, index=escolher_por_query(periodo_options, "periodo", "Todos até amanhã"), label_visibility="visible")
with ctrl2:
    datas_janela_sel = st.multiselect("Data Janela", data_labels, placeholder="Todas")
with ctrl3:
    transp_sel = st.selectbox("Transportadora", transp_options, index=escolher_por_query(transp_options, "transportadora", "Todas"), label_visibility="visible")
with ctrl4:
    turno_sel = st.selectbox("Turno", turno_options, index=escolher_por_query(turno_options, "turno", "Todos"), label_visibility="visible")
with ctrl5:
    status_sel = st.selectbox("Status", status_options, index=escolher_por_query(status_options, "status", "Todos"), label_visibility="visible")
with ctrl6:
    doca_sel = st.selectbox("Doca", doca_options, index=escolher_por_query(doca_options, "doca", "Todas"), label_visibility="visible")
with ctrl7:
    plano_sel = st.selectbox("Plano de Transporte", plano_options, index=escolher_por_query(plano_options, "plano", "Todos"), label_visibility="visible")
with ctrl8:
    modal_sel = st.selectbox("Modal", modal_options, index=escolher_por_query(modal_options, "modal", "Todos"), label_visibility="visible")
with ctrl9:
    saida_sel = st.selectbox("Saída", saida_options, index=escolher_por_query(saida_options, "saida", "Todas"), label_visibility="visible")
with ctrl10:
    tema = st.selectbox("Tema", tema_options, index=escolher_por_query(tema_options, "tema", st.session_state.get("tema_select", "Claro")), key="tema_select", label_visibility="visible")

with st.expander("⚙️ Configurar tempos de alerta e atualização", expanded=False):
    ac1, ac2, ac3, ac4 = st.columns(4)
    with ac1:
        st.number_input("Chegada entra em atenção quando faltar até (min)", min_value=5, max_value=240, step=5, key="alerta_chegada_min")
    with ac2:
        st.number_input("Saída Indicador entra em atenção quando faltar até (min)", min_value=5, max_value=180, step=5, key="alerta_indicador_min")
    with ac3:
        st.number_input("Saída Janela entra em atenção quando faltar até (min)", min_value=5, max_value=180, step=5, key="alerta_janela_min")
    with ac4:
        st.number_input("Atualização automática a cada (seg)", min_value=10, max_value=300, step=5, key="auto_refresh_seconds_v29")

f = base.copy()
if periodo_sel == "Hoje":
    f = f[f["Período"] == "HOJE"]
elif periodo_sel == "Amanhã":
    f = f[f["Período"] == "AMANHÃ"]
elif periodo_sel == "Dias anteriores":
    f = f[f["Período"] == "DIAS ANTERIORES"]
elif periodo_sel == "Até 2h para frente":
    limite_futuro = agora + timedelta(hours=2)
    # Mantém todas as cargas do passado e mostra somente as próximas 2h para frente
    f = f[f["Janela"].notna() & (f["Janela"] <= limite_futuro)]
if datas_janela_sel:
    datas_filtro = [datetime.strptime(x, "%d/%m/%Y").date() for x in datas_janela_sel]
    f = f[f["Janela"].dt.date.isin(datas_filtro)]
if transp_sel != "Todas":
    f = f[f["Transportador"].astype(str) == transp_sel]
if turno_sel != "Todos":
    f = f[f["Turno"].astype(str) == turno_sel]
if "Doca" in f.columns and doca_sel != "Todas":
    f = f[f["Doca"].astype(str) == doca_sel]
if plano_sel != "Todos":
    f = f[f["Plano de Transporte"].astype(str) == plano_sel]
if modal_sel != "Todos":
    f = f[f["Modal"].astype(str) == modal_sel]
if saida_sel == "Somente pendentes":
    f = f[f["Saida Eclusa"].isna()]
elif saida_sel == "Somente com saída":
    f = f[f["Saida Eclusa"].notna()]
if status_sel != "Todos":
    mask_status = (
        f["Status Chegada Monitor"].eq(status_sel)
        | f["Status Saída Indicador Monitor"].eq(status_sel)
        | f["Status Saída Janela Monitor"].eq(status_sel)
    )
    f = f[mask_status]

st.markdown(f"""
<div class="info">
    <div>ℹ️ <b>Ponto de atenção:</b> chegada aparece quando falta até <b>{alerta_chegada_min} min</b> para a Janela; saída indicador até <b>{alerta_indicador_min} min</b>; saída janela até <b>{alerta_janela_min} min</b>.</div>
    <div class="refresh-pill">🔄 Atualização automática ativa | {AUTOREFRESH_SECONDS}s</div>
</div>
""", unsafe_allow_html=True)

# Cards superiores
card_chegada = itens_card(f, "Status Chegada Monitor", "Minutos Chegada", "Chegada Portaria", 60)
card_ind = itens_card(f, "Status Saída Indicador Monitor", "Minutos Indicador", "Saida Eclusa", 30)
card_jan = itens_card(f, "Status Saída Janela Monitor", "Minutos Janela", "Saida Eclusa", 30)

st.markdown(f"""
<div class="cards">
{render_card('ATRASO / ATENÇÃO DE CHEGADA', '🕒', card_chegada, 'chegada', 'Minutos Chegada', 'Status Chegada Monitor', 'chegada', 'Janela')}
{render_card('IMPACTO SAÍDA INDICADOR', '📈', card_ind, 'indicador', 'Minutos Indicador', 'Status Saída Indicador Monitor', 'indicador')}
{render_card('IMPACTO SAÍDA JANELA', '🚪', card_jan, 'janela', 'Minutos Janela', 'Status Saída Janela Monitor', 'janela')}
</div>
""", unsafe_allow_html=True)

# Tabela principal
st.markdown(render_main_table(f), unsafe_allow_html=True)
st.markdown("<div class='after-main-table-spacer'></div>", unsafe_allow_html=True)

# Painel independente de status das cargas por dia
with st.expander("📦 Status das cargas por dia", expanded=True):
    ps1, ps2, ps3 = st.columns([1, 1, 1.4])
    with ps1:
        modal_status_options = ["Todos"] + sorted(base["Modal"].dropna().astype(str).unique().tolist())
        modal_default = "VENDA" if "VENDA" in modal_status_options else "Todos"
        modal_status_sel = st.selectbox("Modal do painel status", modal_status_options, index=escolher_por_query(modal_status_options, "modal_status", modal_default), key="modal_status_dia")
    with ps2:
        turno_status_options = ["Todos"] + sorted(base["Turno"].dropna().astype(str).unique().tolist())
        turno_status_sel = st.selectbox("Turno do painel status", turno_status_options, index=escolher_por_query(turno_status_options, "turno_status", "Todos"), key="turno_status_dia")
    with ps3:
        datas_status = sorted(base.loc[base["Janela"].notna(), "Janela"].dt.date.unique(), reverse=True)
        datas_status_options = [pd.Timestamp(d).strftime("%d/%m/%Y") for d in datas_status]
        datas_status_sel_txt = st.multiselect(
            "Data Janela do painel status",
            datas_status_options,
            default=[],
            placeholder="Todas",
            key="data_status_dia"
        )
        # Quando não selecionar nenhuma data, o painel mostra todas.
        datas_status_sel = [pd.to_datetime(x, dayfirst=True).date() for x in datas_status_sel_txt] if datas_status_sel_txt else None
    st.markdown(render_status_cargas_dia(base, modal_status_sel, turno_status_sel, datas_status_sel), unsafe_allow_html=True)

# Matriz adicional abaixo da tabela, sem reduzir os quadros principais
st.markdown(render_venda_turno_impact(f), unsafe_allow_html=True)

# Resumo diário de cargas VENDA que impactaram o indicador, com M³ e pedidos
st.markdown(render_venda_indicador_dia(f), unsafe_allow_html=True)

# Legendas e rodapé compactos
st.markdown(f"""
<div class="legends">
    <div class="legend-card"><div class="legend-inline"><span class="legend-title">MODAL</span> <span class="badge modal-venda">VENDA</span> <span class="badge modal-transferencia">TRANSF</span> <span class="badge modal-aba">ABA</span></div></div>
    <div class="legend-card"><div class="legend-inline"><span class="legend-title">STATUS</span> <span class="badge st-ok">NO PRAZO</span> <span class="badge st-aguardando">AGUARDANDO</span> <span class="badge st-atencao">PONTO DE ATENÇÃO</span> <span class="badge st-chegada">ATRASO CHEGADA</span> <span class="badge st-atraso">ATRASO SAÍDA</span></div></div>
    <div class="legend-card"><div class="legend-inline"><span class="legend-title">PERÍODO</span> <span><span class="dot amanha"></span>AMANHÃ</span> <span><span class="dot hoje"></span>HOJE</span> <span><span class="dot anterior"></span>ANTERIORES</span></div></div>
</div>
<div class="footer">
    <div>🟢 Atualização automática ativa</div>
    <div>⏱️ Intervalo: <b>{AUTOREFRESH_SECONDS}s</b></div>
    <div>⚙️ Alertas: chegada <b>{alerta_chegada_min}min</b> | indicador <b>{alerta_indicador_min}min</b> | janela <b>{alerta_janela_min}min</b></div>
    <div>↧ Ordenado por <b>Janela decrescente</b></div>
    <div>📅 Exibindo dias anteriores, hoje e amanhã</div>
    <div><b>Dados atualizados em tempo real</b></div>
</div>
""", unsafe_allow_html=True)


with st.sidebar.expander("⚙️ Fonte de dados / trocar base", expanded=False):
    origem = "Excel carregado" if "excel_bytes" in st.session_state else "Google Sheets" if "csv_url" in st.session_state else "Nenhuma"
    st.caption(f"Fonte atual: {origem}")
    novo_link = st.text_input(
        "Link do Google Sheets",
        value=st.session_state.get("csv_url", DEFAULT_GOOGLE_SHEETS_URL),
        placeholder="https://docs.google.com/spreadsheets/d/.../edit?gid=...",
        help="Cole o link da aba da planilha. A aba precisa estar compartilhada/publicada para leitura."
    )
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Usar Google Sheets") and novo_link:
            st.session_state["csv_url"] = novo_link
            st.session_state.pop("excel_bytes", None)
            st.rerun()
    with c2:
        if st.button("Trocar base"):
            st.session_state.pop("excel_bytes", None)
            st.session_state.pop("csv_url", None)
            st.rerun()
