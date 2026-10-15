
import io
import re
import time

import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Consulta de CNPJs 2.0",
    page_icon="🔎",
    layout="wide",
)

st.title("Consulta de CNPJs 2.0")
st.caption("Consulte listas de CNPJs e exporte os resultados.")

OPENCNPJ_URL = "https://kitana.opencnpj.com/cnpj/{}"
BRASILAPI_URL = "https://brasilapi.com.br/api/cnpj/v1/{}"

MAX_CNPJS = 10_000
TIMEOUT = 20
HEADERS = {"User-Agent": "ConsultaCNPJ-2.0/1.0"}
COLUNAS = [
    "CNPJ",
    "Razão social",
    "Situação cadastral",
    "Município",
    "UF",
    "Resultado da consulta",
    "Detalhe",
]


with st.expander("Limites e observações importantes"):
    st.write(
        "O limite do site é de 10.000 CNPJs únicos por lote. "
        "As APIs externas também têm limites próprios e podem ficar "
        "temporariamente indisponíveis. Listas grandes podem demorar."
    )
    st.write(
        "O site tenta consultar o OpenCNPJ primeiro e usa a BrasilAPI "
        "como alternativa. Erros técnicos não significam que um CNPJ "
        "esteja baixado ou inválido."
    )
    st.write(
        "Os dados podem estar desatualizados. Para decisões oficiais, "
        "confirme a situação na Receita Federal."
    )


def normalize_cnpj(value):
    if value is None or pd.isna(value):
        return ""

    digits = re.sub(r"\D", "", str(value).strip())

    if not digits:
        return ""

    return digits.zfill(14) if len(digits) <= 14 else digits


def read_file(file):
    if not file:
        return []

    if file.name.lower().endswith(".xlsx"):
        df = pd.read_excel(file, dtype=str)
    else:
        raw = file.getvalue()

        try:
            df = pd.read_csv(
                io.BytesIO(raw),
                dtype=str,
                sep=None,
                engine="python",
                encoding="utf-8-sig",
            )
        except Exception:
            df = pd.read_csv(
                io.BytesIO(raw),
                dtype=str,
                encoding="latin-1",
            )

    if df.empty:
        return []

    col = next(
        (
            name for name in df.columns
            if str(name).strip().lower()
            in ("cnpj", "cnpjs", "documento")
        ),
        None,
    )

    values = df[col].tolist() if col is not None else df.iloc[:, 0].tolist()

    return [
        normalize_cnpj(value)
        for value in values
        if normalize_cnpj(value)
    ]


def result_row(
    cnpj,
    razao="",
    situacao="",
    municipio="",
    uf="",
    resultado="ERRO",
    detalhe="",
):
    return {
        "CNPJ": cnpj,
        "Razão social": razao,
        "Situação cadastral": situacao,
        "Município": municipio,
        "UF": uf,
        "Resultado da consulta": resultado,
        "Detalhe": detalhe,
    }


def consulta_fonte(url, cnpj, fonte):
    ultimo_erro = ""

    for tentativa in range(3):
        try:
            resposta = requests.get(
                url.format(cnpj),
                headers=HEADERS,
                timeout=TIMEOUT,
            )

            if resposta.status_code == 404:
                return "nao_encontrado", None, "HTTP 404"

            if resposta.status_code == 200:
                json_data = resposta.json()

                if fonte == "OpenCNPJ":
                    if not json_data.get("success"):
                        ultimo_erro = (
                            json_data.get("message")
                            or "A fonte não confirmou a consulta."
                        )
                        return "erro", None, str(ultimo_erro)

                    data = json_data.get("data") or {}

                    situacao = data.get("situacaoCadastral")
                    razao = (
                        data.get("razaoSocial")
                        or data.get("nomeFantasia")
                        or ""
                    )
                    municipio = data.get("municipio") or ""
                    uf = data.get("uf") or ""

                else:
                    data = json_data

                    situacao = (
                        data.get("descricao_situacao_cadastral")
                        or data.get("situacao_cadastral")
                    )
                    razao = (
                        data.get("razao_social")
                        or data.get("nome_fantasia")
                        or ""
                    )
                    municipio = data.get("municipio") or ""
                    uf = data.get("uf") or ""

                if not situacao:
                    return (
                        "erro",
                        None,
                        f"{fonte} respondeu sem informar a situação cadastral.",
                    )

                return (
                    "ok",
                    result_row(
                        cnpj=cnpj,
                        razao=razao,
                        situacao=str(situacao).strip().upper(),
                        municipio=municipio,
                        uf=uf,
                        resultado="OK",
                        detalhe=f"Fonte: {fonte}",
                    ),
                    "",
                )

            ultimo_erro = f"HTTP {resposta.status_code}"

            if resposta.status_code not in (429, 500, 502, 503, 504):
                return "erro", None, ultimo_erro

        except (requests.RequestException, ValueError) as exc:
            ultimo_erro = str(exc)[:250]

        if tentativa < 2:
            time.sleep(3 if tentativa == 0 else 8)

    return "erro", None, ultimo_erro or f"Falha na fonte {fonte}"


def lookup(cnpj):
    erros = []
    fontes = [
        (OPENCNPJ_URL, "OpenCNPJ"),
        (BRASILAPI_URL, "BrasilAPI"),
    ]

    encontrou_404 = False

    for url, fonte in fontes:
        estado, linha, detalhe = consulta_fonte(url, cnpj, fonte)

        if estado == "ok":
            return linha

        if estado == "nao_encontrado":
            encontrou_404 = True
        else:
            erros.append(f"{fonte}: {detalhe}")

    if encontrou_404 and not erros:
        return result_row(
            cnpj,
            situacao="NÃO ENCONTRADO",
            resultado="NÃO ENCONTRADO",
            detalhe="As duas fontes não encontraram esse CNPJ.",
        )

    return result_row(
        cnpj,
        situacao="PENDENTE / ERRO",
        resultado="ERRO",
        detalhe=" | ".join(erros) or "Não foi possível concluir a consulta.",
    )


def make_excel(rows):
    df = pd.DataFrame(rows, columns=COLUNAS)

    resumo = (
        df["Situação cadastral"]
        .fillna("NÃO INFORMADA")
        .value_counts()
        .rename_axis("Situação cadastral")
        .reset_index(name="Quantidade")
    )

    out = io.BytesIO()

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Resultados", index=False)
        resumo.to_excel(writer, sheet_name="Resumo", index=False)

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(horizontal="center")

            for cells in ws.columns:
                width = min(
                    max(
                        max(len(str(cell.value or "")) for cell in cells) + 2,
                        12,
                    ),
                    42,
                )
                ws.column_dimensions[
                    get_column_letter(cells[0].column)
                ].width = width

    out.seek(0)
    return out.getvalue()


uploaded = st.file_uploader(
    "Enviar Excel (.xlsx) ou CSV (.csv)",
    type=["xlsx", "csv"],
)

pasted = st.text_area(
    "Ou cole os CNPJs (um por linha, ou separados por vírgulas/espaços)",
    height=180,
)

try:
    values = read_file(uploaded)
except Exception as exc:
    st.error(f"Não consegui ler o arquivo enviado: {exc}")
    values = []

values += [
    normalize_cnpj(value)
    for value in re.split(r"[\s,;]+", pasted or "")
    if normalize_cnpj(value)
]

seen = set()
cnpjs = []

for value in values:
    if value not in seen:
        seen.add(value)
        cnpjs.append(value)

invalid = [value for value in cnpjs if len(value) != 14]
valid = [value for value in cnpjs if len(value) == 14]

a, b, c = st.columns(3)
a.metric("CNPJs únicos", f"{len(cnpjs):,}".replace(",", "."))
b.metric("Formato de 14 dígitos", f"{len(valid):,}".replace(",", "."))
c.metric("Formato inválido", f"{len(invalid):,}".replace(",", "."))

if len(cnpjs) > MAX_CNPJS:
    st.error(
        "O limite é 10.000 CNPJs únicos por lote. Divida a lista em lotes menores."
    )
    st.stop()

if invalid:
    with st.expander(f"Ver {len(invalid)} item(ns) com formato inválido"):
        st.write(invalid[:200])


if "results" not in st.session_state:
    st.session_state.results = None

st.subheader("Consultar")
st.write(f"Prontos para consulta: **{len(valid)}** CNPJs com 14 dígitos.")

if st.button("Iniciar consulta", type="primary", disabled=not valid):
    rows = [
        result_row(
            value,
            situacao="FORMATO INVÁLIDO",
            resultado="INVÁLIDO",
            detalhe="O número não tem 14 dígitos.",
        )
        for value in invalid
    ]

    progress = st.progress(0, text="Iniciando consulta...")
    status = st.empty()
    total = len(valid)

    for indice, cnpj in enumerate(valid, start=1):
        linha = lookup(cnpj)
        rows.append(linha)

        progress.progress(
            indice / total,
            text=f"Consultados {indice:,} de {total:,}".replace(",", "."),
        )
        status.write(f"Concluídos: {indice:,} / {total:,}".replace(",", "."))

        # Pequena pausa entre CNPJs para reduzir a pressão sobre as APIs.
        if indice < total:
            time.sleep(0.35)

    ordem = {value: indice for indice, value in enumerate(cnpjs)}
    rows.sort(key=lambda row: ordem.get(row["CNPJ"], 10**9))

    st.session_state.results = rows
    progress.progress(1.0, text="Consulta finalizada.")
    status.success("Consulta concluída.")


if st.session_state.results is not None:
    df = pd.DataFrame(st.session_state.results, columns=COLUNAS)

    st.subheader("Resultados")

    x, y, z = st.columns(3)
    x.metric("Processados", len(df))
    y.metric(
        "ATIVAS",
        int((df["Situação cadastral"].str.upper() == "ATIVA").sum()),
    )
    z.metric(
        "Erros",
        int((df["Resultado da consulta"] == "ERRO").sum()),
    )

    st.dataframe(df, use_container_width=True, hide_index=True)

    st.download_button(
        "Baixar Excel",
        make_excel(st.session_state.results),
        "resultado_consulta_cnpjs.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.download_button(
        "Baixar CSV",
        df.to_csv(index=False).encode("utf-8-sig"),
        "resultado_consulta_cnpjs.csv",
        "text/csv",
    )

st.divider()
st.caption(
    "ConsultaCNPJ 2.0 • Os resultados dependem da disponibilidade "
    "e da atualização das fontes consultadas."
)


st.divider()
st.caption("ConsultaCNPJ 2.0 • Os resultados dependem da disponibilidade e atualização da fonte consultada.")
