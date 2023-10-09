import streamlit as st
import pandas as pd
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

st.session_state['arquivo_para_baixar'] = False


@st.cache_data
def baixarPlanilha(dataframe, index=False):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    dataframe.to_excel(writer, index=index, sheet_name="Sheet1")
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0'})
    worksheet.set_column('A:A', None, format1)
    writer.save()
    processed_data = output.getvalue()
    return processed_data


arquivos_turmas = st.file_uploader(
    "Arquivos das Turmas",
    accept_multiple_files=True,
    type=['XLS', "XLSX"],
    key="ARQUIVOS_TURMAS"
)

if st.session_state["ARQUIVOS_TURMAS"]:
    ver_turmas, baixar_turmas = st.tabs(["Ver Turmas", "Baixar Lista das Turmas"])
    with ver_turmas:
        st.radio(
            label="Selecione um arquivo",
            options=[arquivo.name for arquivo in arquivos_turmas],
            horizontal=True,
            key="ARQUIVO_SELECIONADO"
        )
        for arquivo in arquivos_turmas:
            if arquivo.name == st.session_state['ARQUIVO_SELECIONADO']:
                st.radio(
                    label="Selecione a Turma",
                    options=pd.read_excel(arquivo, None).keys(),
                    horizontal=True,
                    key="TURMA_SELECIONADA"
                )
                turma = pd.DataFrame(
                    pd.read_excel(
                        arquivo,
                        sheet_name=st.session_state['TURMA_SELECIONADA'],
                        skiprows=st.number_input(
                            "Ajustar Cabecalho", min_value=0, max_value=10, key=f"{arquivo}"
                        )
                    )
                ).dropna()
                st.dataframe(turma, use_container_width=True, hide_index=True)
    with baixar_turmas:
        turmas_para_baixar = []
        padrao = st.number_input("Digite um valor padrão para o cabecalho", min_value=0)
        st.multiselect(
            label="Selecione um arquivo",
            options=[arquivo.name for arquivo in arquivos_turmas],
            key="ARQUIVO_SELECIONADO_BAIXAR"
        )
        for arquivo_baixar in arquivos_turmas:
            with st.expander(f"Arquivo: {arquivo_baixar.name}"):
                if arquivo_baixar.name in st.session_state['ARQUIVO_SELECIONADO_BAIXAR']:
                    turmas_baixar = st.multiselect(
                        label=f"Selecione a Turma-{arquivo_baixar.name}",
                        options=pd.read_excel(arquivo_baixar, None).keys(),
                        key=f"TURMA_SELECIONADA_BAIXAR-{arquivo_baixar.name}"
                    )
                    for turma_baixar in turmas_baixar:
                        if turma_baixar in turmas_baixar:
                            turma_ver = pd.DataFrame(
                                pd.read_excel(
                                    arquivo_baixar,
                                    sheet_name=turma_baixar,
                                    skiprows=st.number_input(
                                        f"Ajustar Cabecalho para turma: {turma_baixar}", min_value=0,
                                        key=f"{arquivo_baixar.name}-{turma_baixar}",
                                        value=padrao
                                    )
                                )
                            ).dropna()
                            turma_ver['Turma'] = f'{turma_baixar[-4]}º ANO - {turma_baixar[-1]}'
                            turma_ver['Turno'] = "TARDE" if f'{turma_baixar[-3]}' == "T" else (
                                "MANHÃ" if f'{turma_baixar[-3]}' == "M" else "OUTRO")
                            st.dataframe(turma_ver, use_container_width=True, hide_index=True)
                            turmas_para_baixar.append(turma_ver)
                    st.session_state['arquivo_para_baixar'] = True
if st.session_state['arquivo_para_baixar']:
    st.divider()
    turmas_totais = pd.DataFrame(pd.concat(turmas_para_baixar))
    st.write(f"Total de Alunos: {len(turmas_totais)}")
    colunas_remover = st.multiselect("Selecione as colunas que deseja remover", options=turmas_totais.keys())
    turmas_totais = turmas_totais.drop(colunas_remover, axis=1)
    turmas_totais["Nº"] = [n for n in range(1, len(turmas_totais) + 1)]
    turmas_totais = turmas_totais.set_index('Nº')
    with st.expander("Metodos de contagem"):
        st.multiselect("Agrupar por", turmas_totais.keys(), key='AGRUPAR_POR_COLUNAS')
        if st.session_state['AGRUPAR_POR_COLUNAS']:
            df = turmas_totais.groupby(
                by=st.session_state['AGRUPAR_POR_COLUNAS']
            ).count()
            st.dataframe(df, use_container_width=True)
            planilha = baixarPlanilha(df, True)
            st.download_button("📥 Baixar Lista de contagem dos alunos", data=planilha,
                               file_name="Lista de contagem dos alunos completa.xlsx")
        else:
            st.warning("Selecione as colunas adequadas")

    st.divider()
    st.dataframe(turmas_totais, use_container_width=True)
    planilha = baixarPlanilha(turmas_totais)
    st.download_button("📥 Baixar Lista de alunos", data=planilha, file_name="Lista do Alunos Completa.xlsx")
