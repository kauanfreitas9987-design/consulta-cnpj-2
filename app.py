
import io
import re
import time

import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Consulta de CNPJs",
    page_icon="🔎",
    layout="centered",
)

st.title("Consulta de CNPJs")
st.caption("Consulte o status cadastral e baixe o resultado em Excel.")

OPENCNPJ_URL = "https://kitana.opencnpj.com/cnpj/{}"
BRASILAPI_URL = "https://brasilapi.com.br/api/cnpj/v1/{}"

TIMEOUT = 20
MAX_CNPJS = 10_000
HEADERS = {"User-Agent": "ConsultaCNPJ/1.0"}

COLUNAS = ["CNPJ", "Status"]


def normalizar_cnpj(valor):
    if valor is None or pd.isna(valor):
        return ""

    digitos = re.sub(r"\D", "", str(valor).strip())

    if not digitos:
        return ""

    return digitos.zfill(14) if len(digitos) <= 14 else digitos


def ler_arquivo(arquivo):
    if arquivo is None:
        return []

    if arquivo.name.lower().endswith(".xlsx"):
        df = pd.read_excel(arquivo, dtype=str)
    else:
        df = pd.read_csv(
            arquivo,
            dtype=str,
            sep=None,
            engine="python",
            encoding="utf-8-sig",
        )

    if df.empty:
        return []

    coluna = next(
        (
            c for c in df.columns
            if str(c).strip().lower() in (
                "cnpj", "cnpjs", "documento"
            )
        ),
        None,
    )

    valores = (
        df[coluna].tolist()
        if coluna is not None
        else df.iloc[:, 0].tolist()
    )

    return [
        normalizar_cnpj(v)
        for v in valores
        if normalizar_cnpj(v)
    ]


def normalizar_status(situacao):
    if situacao is None:
        return None

    texto = str(situacao).strip().upper()

    # Códigos oficiais de situação cadastral da Receita Federal.
    codigos = {
        "01": "NULO",
        "02": "ATIVO",
        "03": "SUSPENSO",
        "04": "INAPTO",
        "08": "BAIXADO",
    }

    # Primeiro, procura a descrição da situação.
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

    # Se vier um código numérico, tenta identificá-lo.
    codigo = re.match(r"^\s*(\d{1,2})(?:\s*[-–:.]\s*|\s*$)", texto)

    if codigo:
        numero = codigo.group(1).zfill(2)
        return codigos.get(numero)

    return None


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

        if fonte == "OpenCNPJ":
            if not dados.get("success"):
                return None

            dados = dados.get("data") or {}
            situacao = dados.get("situacaoCadastral")
        else:
            situacao = (
                dados.get("descricao_situacao_cadastral")
                or dados.get("situacao_cadastral")
            )

        return normalizar_status(situacao)

    except (requests.RequestException, ValueError, TypeError):
        return None


def consultar_cnpj(cnpj):
    for tentativa in range(2):
        for url, fonte in (
            (OPENCNPJ_URL, "OpenCNPJ"),
            (BRASILAPI_URL, "BrasilAPI"),
        ):
            status = consultar_fonte(url, cnpj, fonte)

            if status:
                return {"CNPJ": cnpj, "Status": status}

        if tentativa == 0:
            time.sleep(2)

    return {"CNPJ": cnpj, "Status": "ERRO NA CONSULTA"}


def gerar_excel(resultados):
    df = pd.DataFrame(resultados, columns=COLUNAS)
    arquivo = io.BytesIO()

    with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="CNPJs")

        planilha = writer.sheets["CNPJs"]
        planilha.freeze_panes = "A2"
        planilha.auto_filter.ref = planilha.dimensions

        for celula in planilha[1]:
            celula.font = Font(bold=True, color="FFFFFF")
            celula.fill = PatternFill(
                "solid", fgColor="1F4E78"
            )
            celula.alignment = Alignment(horizontal="center")

        planilha.column_dimensions["A"].width = 22
        planilha.column_dimensions["B"].width = 24

    arquivo.seek(0)
    return arquivo.getvalue()


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

cnpjs += [
    normalizar_cnpj(v)
    for v in re.split(r"[\s,;]+", texto or "")
    if normalizar_cnpj(v)
]

# Remove duplicados, mantendo a ordem original.
cnpjs = list(dict.fromkeys(cnpjs))

invalidos = [c for c in cnpjs if len(c) != 14]
validos = [c for c in cnpjs if len(c) == 14]

st.caption(f"{len(cnpjs)} CNPJs recebidos • {len(validos)} prontos para consulta")

if len(cnpjs) > MAX_CNPJS:
    st.error("O limite é de 10.000 CNPJs por consulta.")
    st.stop()

if invalidos:
    with st.expander(f"Ver {len(invalidos)} entradas com formato inválido"):
        st.write(invalidos[:200])

if "resultados" not in st.session_state:
    st.session_state.resultados = None

if st.button(
    "Consultar CNPJs",
    type="primary",
    disabled=not cnpjs,
    use_container_width=True,
):
    resultados = [
        {"CNPJ": c, "Status": "ERRO NA CONSULTA"}
        for c in invalidos
    ]

    barra = st.progress(0, text="Iniciando consulta...")

    for i, cnpj in enumerate(validos, start=1):
        resultados.append(consultar_cnpj(cnpj))

        barra.progress(
            i / max(len(validos), 1),
            text=f"Consultando {i} de {len(validos)} CNPJs",
        )

        if i < len(validos):
            time.sleep(0.3)

    ordem = {c: i for i, c in enumerate(cnpjs)}
    resultados.sort(key=lambda r: ordem.get(r["CNPJ"], 10**9))

    st.session_state.resultados = resultados
    barra.progress(1.0, text="Consulta concluída!")

if st.session_state.resultados is not None:
    df = pd.DataFrame(
        st.session_state.resultados,
        columns=COLUNAS,
    )

    st.subheader("Resultados")
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Baixar Excel",
        data=gerar_excel(st.session_state.resultados),
        file_name="consulta_cnpjs.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

st.caption(
    "Os resultados dependem da disponibilidade e atualização "
    "das fontes consultadas. Confirme situações duvidosas "
    "nos canais oficiais da Receita Federal."
)
