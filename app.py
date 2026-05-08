import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

st.set_page_config(page_title="Classificador de Pautas")

st.title("Classificador de Pautas ALMG")

arquivo = st.file_uploader(
    "Envie a pauta em PDF",
    type=["pdf"]
)

def extrair_texto_pdf(arquivo_pdf):

    texto = ""

    with pdfplumber.open(arquivo_pdf) as pdf:

        for pagina in pdf.pages:

            conteudo = pagina.extract_text()

            if conteudo:
                texto += conteudo + "\n"

    return texto

def separar_cargo_orgao(descricao):

    partes = descricao.split(" - ")

    cargo = partes[0] if len(partes) > 0 else ""
    orgao = partes[1] if len(partes) > 1 else ""
    sigla = partes[2] if len(partes) > 2 else ""

    return cargo, orgao, sigla

def extrair_dados(texto):

    convidados = []

    comissao = ""
    data = ""
    horario = ""
    local = ""

    m = re.search(r"\n([A-Za-zÀ-ÿ ]+)\nDia:", texto)

    if m:
        comissao = m.group(1).strip()

    m = re.search(r"Dia:\s*(\d{2}/\d{2}/\d{4})", texto)

    if m:
        data = m.group(1)

    m = re.search(r"Horário:([0-9:]+)", texto)

    if m:
        horario = m.group(1)

    m = re.search(r"Local:\s*(.*?)\s*Tel:", texto)

    if m:
        local = m.group(1)

    if "Convidados:" in texto:

        trecho = texto.split("Convidados:")[1]

        if "2ª FASE" in trecho:
            trecho = trecho.split("2ª FASE")[0]

        linhas = [
            l.strip()
            for l in trecho.split("\n")
            if l.strip()
        ]

        i = 0

        while i < len(linhas) - 1:

            nome = linhas[i]
            descricao = linhas[i + 1]

            if (
                "Representante(s)" in nome
                or "Representante(s)" in descricao
            ):
                i += 1
                continue

            presenca = "Não"

            if "presença confirmada" in nome.lower():

                presenca = "Sim"

                nome = (
                    nome
                    .replace("- presença confirmada", "")
                    .strip()
                )

            cargo, orgao, sigla = separar_cargo_orgao(descricao)

            convidados.append({
                "Comissão": comissao,
                "Data": data,
                "Horário": horario,
                "Local": local,
                "Nome": nome,
                "Cargo": cargo,
                "Órgão": orgao,
                "Sigla": sigla,
                "Presença Confirmada": presenca
            })

            i += 2

    return pd.DataFrame(convidados)

def gerar_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Convidados"
        )

    return output.getvalue()

if arquivo:

    texto = extrair_texto_pdf(arquivo)

    st.subheader("Texto extraído")

    st.text_area(
        "Conteúdo da pauta",
        texto,
        height=250
    )

    if st.button("Gerar planilha Excel"):

        df = extrair_dados(texto)

        st.subheader("Dados encontrados")

        st.dataframe(
            df,
            use_container_width=True
        )

        excel = gerar_excel(df)

        st.download_button(
            "Baixar Excel",
            data=excel,
            file_name="pauta_almg.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
