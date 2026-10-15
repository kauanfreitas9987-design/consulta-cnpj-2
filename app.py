
import io
import re
import time

import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment

# ==========================================
# CONFIGURAÇÕES DO SITE
# ==========================================

st.set_page_config(
    page_title="Consulta de CNPJs",
    page_icon="🔎",
    layout="centered",
)

st.title("Consulta de CNPJs")
st.caption("Consulte os status cadastrais e baixe os resultados em Excel.")

OPENCNPJ_URL = "https://kitana.opencnpj.com/cnpj/{}"
BRASILAPI_URL = "https://brasilapi.com.br/api/cnpj/v1/{}"

MAX_CNPJS = 10_000
TIMEOUT = 20
HEADERS = {"User-Agent": "ConsultaCNPJ/1.0"}

COLUNAS = ["CNPJ", "Status"]

STATUS_VALIDOS = [
    "ATIVO",
    "BAIXADO",
    "NULO",
    "INAPTO",
    "SUSPENSO",
]


# ==========================================
# NORMALIZAR CNPJ
# ==========================================

def normalizar_cnpj(valor):
    if valor is None or pd.isna(valor):
        return ""

    digitos = re.sub(r"\D", "", str(valor).strip())

    if not digitos:
        return ""

    return digitos.zfill(14) if len(digitos) <= 14 else digitos


# ==========================================
# LER ARQUIVO EXCEL OU CSV
# ==========================================

def ler_arquivo(arquivo):
    if arquivo is None:
        return []

    if arquivo.name.lower().endswith(".xlsx"):
        df = pd.read_excel(arquivo, dtype=str)
    else:
        conteudo = arquivo.getvalue()

        try:
            df = pd.read_csv(
                io.BytesIO(conteudo),
                dtype=str,
                sep=None,
                engine="python",
                encoding="utf-8-sig",
            )
        except Exception:
            df = pd.read_csv(
                io.BytesIO(conteudo),
                dtype=str,
                encoding="latin-1",
            )

    if df.empty:
        return []

    coluna = next(
        (
            nome for nome in df.columns
            if str(nome).strip().lower()
            in ("cnpj", "cnpjs", "documento")
        ),
        None,
    )

    valores = (
        df[coluna].tolist()
        if coluna is not None
        else df.iloc[:, 0].tolist()
    )

    return [
        normalizar_cnpj(valor)
        for valor in valores
        if normalizar_cnpj(valor)
    ]


# ==========================================
# NORMALIZAR SITUAÇÃO CADASTRAL
# ==========================================

def normalizar_status(situacao):
    if situacao is None:
        return None

    texto = str(situacao).strip().upper()

    # Códigos de situação cadastral.
    codigos = {
        "01": "NULO",
        "02": "ATIVO",
        "03": "SUSPENSO",
        "04": "INAPTO",
        "08": "BAIXADO",
    }

    # Reconhece códigos numéricos, quando fornecidos.
    if re.fullmatch(r"\d{1,2}", texto):
        return codigos.get(texto.zfill(2))

    # Reconhece descrições por extenso.
    if "ATIVA" in texto or "ATIVO" in texto:
        return "ATIVO"

    if "BAIXADA" in texto or "BAIXADO" in texto:
        return "BAIXADO"

    if "INAPTA" in texto or "INAPTO" in texto:
        return "INAPTO"

    if "SUSPENSA" in texto or "SUSPENSO" in texto:
        return "SUSPENSO"

    if "NULA" in texto or "NULO" in texto:
        return "NULO"

    return None


# ==========================================
# CONSULTAR UMA FONTE
# ==========================================

def consultar_fonte(url, cnpj, fonte):
    try:
        resposta = requests.get(
            url.format(cnpj),
            headers=HEADERS,
            timeout=TIMEOUT,
        )

        if resposta.status_code != 200:
            return None

        dados = resposta.json()

        if not isinstance(dados, dict):
            return None

        if fonte == "OpenCNPJ":
            if not dados.get("success"):
                return None

            dados = dados.get("data") or {}

            if not isinstance(dados, dict):
                return None

            situacao = dados.get("situacaoCadastral")

        else:
            situacao = (
                dados.get("descricao_situacao_cadastral")
                or dados.get("situacao_cadastral")
            )

        return normalizar_status(situacao)

    except (
        requests.RequestException,
        ValueError,
        TypeError,
    ):
        return None


# ==========================================
# CONSULTAR CNPJ NAS DUAS APIS
# ==========================================

def consultar_cnpj(cnpj):
    fontes = [
        (OPENCNPJ_URL, "OpenCNPJ"),
        (BRASILAPI_URL, "BrasilAPI"),
    ]

    # Tenta as fontes duas vezes em caso de falha.
    for tentativa in range(2):
        for url, fonte in fontes:
            status = consultar_fonte(url, cnpj, fonte)

            if status in STATUS_VALIDOS:
                return {
                    "CNPJ": cnpj,
                    "Status": status,
                }

        if tentativa == 0:
            time.sleep(2)

    # Não inventa uma situação quando não há confirmação.
    return {
        "CNPJ": cnpj,
        "Status": "ERRO NA CONSULTA",
    }


# ==========================================
# GERAR EXCEL SIMPLES
# ==========================================

def gerar_excel(resultados):
    df = pd.DataFrame(resultados, columns=COLUNAS)
    arquivo = io.BytesIO()

    with pd.ExcelWriter(
        arquivo,
        engine="openpyxl",
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="CNPJs",
        )

        planilha = writer.sheets["CNPJs"]

        planilha.freeze_panes = "A2"
        planilha.auto_filter.ref = planilha.dimensions

        for celula in planilha[1]:
            celula.font = Font(
                bold=True,
                color="FFFFFF",
            )
            celula.fill = PatternFill(
                "solid",
                fgColor="1F4E78",
            )
            celula.alignment = Alignment(
                horizontal="center",
            )

        planilha.column_dimensions["A"].width = 22
        planilha.column_dimensions["B"].width = 25

    arquivo.seek(0)
    return arquivo.getvalue()


# ==========================================
# ENTRADA DOS CNPJS
# ==========================================

arquivo = st.file_uploader(
    "Enviar arquivo Excel ou CSV (opcional)",
    type=["xlsx", "csv"],
)

texto = st.text_area(
    "Cole os CNPJs abaixo",
    placeholder="Um CNPJ por linha",
    height=150,
)

try:
    cnpjs = ler_arquivo(arquivo)
except Exception:
    st.error("Não foi possível ler o arquivo enviado.")
    cnpjs = []

# Aceita CNPJs colados, separados por linhas,
# vírgulas, ponto e vírgula ou espaços.
cnpjs += [
    normalizar_cnpj(valor)
    for valor in re.split(r"[\s,;]+", texto or "")
    if normalizar_cnpj(valor)
]

# Remove duplicados sem alterar a ordem.
cnpjs = list(dict.fromkeys(cnpjs))

validos = [
    cnpj for cnpj in cnpjs
    if len(cnpj) == 14
]

invalidos = [
    cnpj for cnpj in cnpjs
    if len(cnpj) != 14
]

st.caption(
    f"CNPJs únicos: {len(cnpjs):,}".replace(",", ".")
)

if len(cnpjs) > MAX_CNPJS:
    st.error(
        "O limite é de 10.000 CNPJs únicos por consulta. "
        "Divida a lista em lotes menores."
    )
    st.stop()

if invalidos:
    with st.expander(
        f"Ver {len(invalidos)} entrada(s) com formato inválido"
    ):
        st.write(invalidos[:200])


# ==========================================
# GUARDAR RESULTADOS
# ==========================================

if "resultados_cnpjs" not in st.session_state:
    st.session_state.resultados_cnpjs = None


# ==========================================
# BOTÃO DE CONSULTA
# ==========================================

if st.button(
    "Consultar CNPJs",
    type="primary",
    use_container_width=True,
    disabled=not cnpjs,
):

    resultados = []

    # Registra entradas inválidas separadamente.
    for cnpj in invalidos:
        resultados.append({
            "CNPJ": cnpj,
            "Status": "FORMATO INVÁLIDO",
        })

    barra = st.progress(
        0,
        text="Iniciando consulta...",
    )

    total_validos = len(validos)

    for indice, cnpj in enumerate(validos, start=1):
        resultado = consultar_cnpj(cnpj)
        resultados.append(resultado)

        barra.progress(
            indice / max(total_validos, 1),
            text=(
                f"Consultando {indice:,} de "
                f"{total_validos:,} CNPJs"
            ).replace(",", "."),
        )

        if indice < total_validos:
            time.sleep(0.3)

    # Mantém a mesma ordem da lista original.
    ordem = {
        cnpj: indice
        for indice, cnpj in enumerate(cnpjs)
    }

    resultados.sort(
        key=lambda item: ordem.get(
            item["CNPJ"],
            10**9,
        )
    )

    st.session_state.resultados_cnpjs = resultados

    barra.progress(
        1.0,
        text="Consulta concluída!",
    )


# ==========================================
# PAINEL DE RESULTADOS
# ==========================================

if st.session_state.resultados_cnpjs is not None:

    df = pd.DataFrame(
        st.session_state.resultados_cnpjs,
        columns=COLUNAS,
    )

    # Padroniza os textos para contar corretamente.
    contagens = (
        df["Status"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .value_counts()
    )

    total = len(df)

    ativos = int(contagens.get("ATIVO", 0))
    baixados = int(contagens.get("BAIXADO", 0))
    nulos = int(contagens.get("NULO", 0))
    inaptos = int(contagens.get("INAPTO", 0))
    suspensos = int(contagens.get("SUSPENSO", 0))
    erros = int(contagens.get("ERRO NA CONSULTA", 0))
    invalidos = int(contagens.get("FORMATO INVÁLIDO", 0))

    st.divider()
    st.subheader("Painel de resultados")

    # Total geral.
    st.metric(
        "Total processado",
        f"{total:,}".replace(",", "."),
    )

    # Primeira linha.
    col1, col2 = st.columns(2)

    col1.metric(
        "Ativos",
        f"{ativos:,}".replace(",", "."),
    )

    col2.metric(
        "Baixados",
        f"{baixados:,}".replace(",", "."),
    )

    # Segunda linha.
    col3, col4 = st.columns(2)

    col3.metric(
        "Nulos",
        f"{nulos:,}".replace(",", "."),
    )

    col4.metric(
        "Inaptos",
        f"{inaptos:,}".replace(",", "."),
    )

    # Terceira linha.
    col5, col6 = st.columns(2)

    col5.metric(
        "Suspensos",
        f"{suspensos:,}".replace(",", "."),
    )

    col6.metric(
        "Erros na consulta",
        f"{erros:,}".replace(",", "."),
    )

    # Formatos inválidos não são situações cadastrais.
    if invalidos:
        st.metric(
            "Formato inválido",
            f"{invalidos:,}".replace(",", "."),
        )

    st.divider()
    st.subheader("Resultados individuais")

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Baixar Excel",
        data=gerar_excel(
            st.session_state.resultados_cnpjs
        ),
        file_name="consulta_cnpjs.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
    )

st.divider()

st.caption(
    "Os resultados dependem da disponibilidade e da atualização "
    "das fontes consultadas. Para confirmação oficial, consulte "
    "a Receita Federal."
)
