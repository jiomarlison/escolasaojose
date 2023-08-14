import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Escola São José",
    layout="wide",
    initial_sidebar_state="expanded", )

arquivos = st.file_uploader("Arquivos das Turmas", type=["xlsx", "xls"], accept_multiple_files=True)
arqSelecionado = st.radio("Arquivos", options=reversed([arquivo.name for arquivo in arquivos]), horizontal=True)
if arquivos is not None:
    for arquivo in arquivos:
        if arquivo.name == arqSelecionado:
            Turmas = pd.read_excel(arquivo, None).keys()
            turmaSelecionada = st.radio("Selecione uma turma", options=Turmas, horizontal=True)
            pandas = pd.DataFrame(
                pd.read_excel(
                    arquivo, header=st.number_input("Ajuste ate o cabecalho estar certo", min_value=0,
                                                    max_value=10)
                )).dropna()
            colunasRemovidas = st.multiselect("Colunas a remover", options=pandas.keys())
            pandas = pandas.drop(colunasRemovidas, axis=1)
            st.dataframe(pandas, use_container_width=True, hide_index=True)
            # for Turma in Turmas:
            #     st.dataframe(pd.read_excel(arquivo, sheet_name=Turma).dropna(), use_container_width=True)
    # for arquivo in arquivos:
    #     Turmas = pd.read_excel(arquivo, None).keys()
    #     st.header(arquivo.name)
    #     st.subheader(Turmas)
    #     for Turma in Turmas:
    #         st.text(f"Turma: {Turma[-4]}º Ano - {Turma[-1] }, Turno: {'Manhã' if Turma[-3] == 'M' else 'Tarde'}")
    #         st.dataframe(pd.read_excel(arquivo, sheet_name=Turma).dropna(), use_container_width=True)
