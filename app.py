import io
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Consulta de CNPJs 2.0", page_icon="🔎", layout="wide")
st.title("Consulta de CNPJs 2.0")
st.caption("Consulte listas de CNPJs e exporte os resultados.")

API_URL = "https://brasilapi.com.br/api/cnpj/v1/{}"
MAX_CNPJS = 10_000
WORKERS = 5
TIMEOUT = 20
HEADERS = {"User-Agent": "ConsultaCNPJ-2.0/1.0"}

with st.expander("Limites e observações importantes"):
    st.write("O site aceita até 10.000 CNPJs únicos por lote, mas a BrasilAPI pode impor limites próprios ou ficar temporariamente indisponível. O processamento pode demorar. Alguns resultados podem ficar pendentes/erro.")
    st.write("A situação depende dos dados retornados pela BrasilAPI e pode estar desatualizada. Para decisões oficiais, confirme na Receita Federal.")

def normalize_cnpj(value):
    if value is None or pd.isna(value):
        return ""
    digits = re.sub(r"\D", "", str(value).strip())
    if not digits:
        return ""
    return digits.zfill(14) if len(digits) < 14 else digits

def read_file(file):
    if not file:
        return []
    if file.name.lower().endswith(".xlsx"):
        df = pd.read_excel(file, dtype=str)
    else:
        raw = file.getvalue()
        try:
            df = pd.read_csv(io.BytesIO(raw), dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(io.BytesIO(raw), dtype=str, encoding="latin-1")
    if df.empty:
        return []
    col = next((c for c in df.columns if str(c).strip().lower() in ("cnpj", "cnpjs", "documento")), None)
    values = df[col].tolist() if col is not None else df.iloc[:, 0].tolist()
    return [normalize_cnpj(v) for v in values if normalize_cnpj(v)]

uploaded = st.file_uploader("Enviar Excel (.xlsx) ou CSV (.csv)", type=["xlsx", "csv"])
pasted = st.text_area("Ou cole os CNPJs (um por linha, ou separados por vírgulas/espaços)", height=180)
values = read_file(uploaded) + [normalize_cnpj(x) for x in re.split(r"[\s,;]+", pasted or "") if normalize_cnpj(x)]
seen, cnpjs = set(), []
for value in values:
    if value not in seen:
        seen.add(value)
        cnpjs.append(value)

invalid = [c for c in cnpjs if len(c) != 14]
valid = [c for c in cnpjs if len(c) == 14]
a, b, c = st.columns(3)
a.metric("CNPJs únicos", f"{len(cnpjs):,}".replace(",", "."))
b.metric("Formato de 14 dígitos", f"{len(valid):,}".replace(",", "."))
c.metric("Formato inválido", f"{len(invalid):,}".replace(",", "."))

if len(cnpjs) > MAX_CNPJS:
    st.error("O limite é 10.000 CNPJs únicos por lote. Divida a lista em lotes menores.")
    st.stop()

if invalid:
    with st.expander(f"Ver {len(invalid)} item(ns) com formato inválido"):
        st.write(invalid[:200])

def lookup(cnpj):
    last_error = ""
    for attempt in range(3):
        try:
            r = requests.get(API_URL.format(cnpj), headers=HEADERS, timeout=TIMEOUT)
            if r.status_code == 200:
                data = r.json()
                status = data.get("descricao_situacao_cadastral") or data.get("situacao_cadastral") or "SITUAÇÃO NÃO INFORMADA"
                return {"CNPJ": cnpj, "Razão social": data.get("razao_social") or data.get("nome_fantasia") or "",
                        "Situação cadastral": str(status).strip().upper(), "Município": data.get("municipio") or "",
                        "UF": data.get("uf") or "", "Resultado da consulta": "OK", "Detalhe": ""}
            if r.status_code == 404:
                return {"CNPJ": cnpj, "Razão social": "", "Situação cadastral": "NÃO ENCONTRADO",
                        "Município": "", "UF": "", "Resultado da consulta": "NÃO ENCONTRADO",
                        "Detalhe": "A API não encontrou esse CNPJ."}
            last_error = f"HTTP {r.status_code}"
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(2 ** attempt)
                continue
            break
        except requests.RequestException as exc:
            last_error = str(exc)[:250]
            time.sleep(2 ** attempt)
    return {"CNPJ": cnpj, "Razão social": "", "Situação cadastral": "PENDENTE / ERRO",
            "Município": "", "UF": "", "Resultado da consulta": "ERRO", "Detalhe": last_error}

def make_excel(rows):
    df = pd.DataFrame(rows)
    summary = df["Situação cadastral"].fillna("NÃO INFORMADA").value_counts().rename_axis("Situação cadastral").reset_index(name="Quantidade")
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Resultados", index=False)
        summary.to_excel(writer, sheet_name="Resumo", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(horizontal="center")
            for cells in ws.columns:
                width = min(max(max(len(str(x.value or "")) for x in cells) + 2, 12), 42)
                ws.column_dimensions[get_column_letter(cells[0].column)].width = width
    out.seek(0)
    return out.getvalue()

if "results" not in st.session_state:
    st.session_state.results = None

st.subheader("Consultar")
st.write(f"Prontos para consulta: **{len(valid)}** CNPJs com 14 dígitos.")
if st.button("Iniciar consulta", type="primary", disabled=not valid):
    rows = [{"CNPJ": x, "Razão social": "", "Situação cadastral": "FORMATO INVÁLIDO",
             "Município": "", "UF": "", "Resultado da consulta": "INVÁLIDO",
             "Detalhe": "O número não tem 14 dígitos."} for x in invalid]
    progress = st.progress(0, text="Iniciando...")
    status = st.empty()
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(lookup, x): x for x in valid}
        total, done = len(futures), 0
        for future in as_completed(futures):
            try:
                rows.append(future.result())
            except Exception as exc:
                rows.append({"CNPJ": futures[future], "Razão social": "", "Situação cadastral": "PENDENTE / ERRO",
                             "Município": "", "UF": "", "Resultado da consulta": "ERRO", "Detalhe": str(exc)[:250]})
            done += 1
            progress.progress(done / total, text=f"Consultados {done:,} de {total:,}".replace(",", "."))
            status.write(f"Concluídos: {done:,} / {total:,}".replace(",", "."))
    order = {x: i for i, x in enumerate(cnpjs)}
    rows.sort(key=lambda row: order.get(row["CNPJ"], 10**9))
    st.session_state.results = rows
    progress.progress(1.0, text="Consulta finalizada.")
    status.success("Consulta concluída.")

if st.session_state.results is not None:
    df = pd.DataFrame(st.session_state.results)
    st.subheader("Resultados")
    x, y, z = st.columns(3)
    x.metric("Processados", len(df))
    y.metric("ATIVOS", int((df["Situação cadastral"].str.upper() == "ATIVA").sum()))
    z.metric("Erros/Inválidos", int(df["Resultado da consulta"].isin(["ERRO", "INVÁLIDO"]).sum()))
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button("Baixar Excel", make_excel(st.session_state.results),
        "resultado_consulta_cnpjs.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    st.download_button("Baixar CSV", df.to_csv(index=False).encode("utf-8-sig"),
        "resultado_consulta_cnpjs.csv", "text/csv")

st.divider()
st.caption("ConsultaCNPJ 2.0 • Os resultados dependem da disponibilidade e atualização da fonte consultada.")
