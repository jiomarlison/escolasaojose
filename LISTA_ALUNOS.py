import streamlit as st
import pandas as pd
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import datetime as dttm
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
    writer.close()
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
        nome_escola = st.text_input("Digite o nome da escola")
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
    st.header(f"**Total de Alunos: {len(turmas_totais)}**")
    colunas_remover = st.multiselect(":red[**Selecione a(s) coluna(s) que deseja remover**]",
                                     options=turmas_totais.keys())
    turmas_totais = turmas_totais.drop(colunas_remover, axis=1)

    with st.expander(":red[BAIXAR PLANILHA DE CONTAGEM]"):
        st.multiselect("Agrupar por", turmas_totais.keys(), key='AGRUPAR_POR_COLUNAS', max_selections=3)
        if st.session_state['AGRUPAR_POR_COLUNAS']:
            df = turmas_totais
            df['Quantidade Alunos'] = ''
            df = df.groupby(
                by=st.session_state['AGRUPAR_POR_COLUNAS']
            ).count()
            colunas = list(df.columns).pop()
            df = df[colunas]
            st.dataframe(df, use_container_width=True)
            planilha = baixarPlanilha(df, True)
            st.download_button("📥 Baixar Lista de contagem dos alunos", data=planilha,
                               file_name="Lista de contagem dos alunos completa.xlsx")
        else:
            st.warning("Selecione as colunas adequadas")

    st.divider()
    colunas_adicionar = st.multiselect(
        ":green[**Selecione coluna(s) a ser(em) adicionada(s)**]",
        options=['OBSERVAÇÃO', 'ASSINATURA', 'NOTA 1', 'NOTA 2', 'NOTA 3', 'NOTA 4', 'PROVA', 'MÉDIA']
    )
    ordenar_por = st.multiselect(
        "Ordenar por",
        options=reversed(turmas_totais.keys()),
        max_selections=3
    )

    if colunas_adicionar:
        turmas_totais[colunas_adicionar] = ''
    if ordenar_por:
        turmas_totais = turmas_totais.sort_values(by=ordenar_por)

    cabecalho_planilha = list(turmas_totais.keys())
    cabecalho_planilha = pd.MultiIndex.from_product([[nome_escola], cabecalho_planilha])
    turmas_totais = pd.DataFrame(columns=cabecalho_planilha, data=turmas_totais.values)

    st.data_editor(
        turmas_totais,
        use_container_width=True,
        column_config={
            'Matrícula': st.column_config.NumberColumn("Matrícula", format="%d")
        },
        disabled=True
    )
    planilha = baixarPlanilha(turmas_totais, True)
    st.download_button("📥 Baixar Lista de alunos das turmas selecionadas", data=planilha,
                       file_name=f"Lista do Alunos Completa{dttm.datetime.today().strftime('%d.%m.%Y')}.xlsx")
