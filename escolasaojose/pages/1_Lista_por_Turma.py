import pandas as pd
import streamlit as st

arquivos = st.file_uploader("*:blue[Arquivos das Turmas]*", type=["xlsx", "xls"], accept_multiple_files=True, help='Faça o Upload dos arquivos, Limite de 200Mb Total')
arqSelecionado = st.multiselect("***:green[Arquivos Disponiveis]***", options=reversed([arquivo.name for arquivo in arquivos]))


if arquivos is not None:
    for i, arquivo in enumerate(arquivos):
        if arquivo.name in arqSelecionado:
            with st.expander(f"**Turmas arquivo: :red[***{arquivo.name}***]**"):
                turmas = pd.read_excel(arquivo, None).keys()
                ajusteCabecalho = st.number_input(f"**:blue[Ajuste o Cabeçalho para {arquivo.name}]**", min_value=0, max_value=10)
                colunasRemovidas = st.multiselect(f"*:blue[Colunas a remover de {arquivo.name}]*", options=pd.DataFrame(pd.read_excel(arquivo, header=ajusteCabecalho, sheet_name=list(turmas)[0])).dropna().keys())
                for turma in turmas:
                    pandas = pd.DataFrame(
                        pd.read_excel(
                            arquivo,
                            header=ajusteCabecalho,
                            sheet_name=turma
                        )
                    ).dropna().drop(colunasRemovidas, axis=1)

                    st.dataframe(pandas, use_container_width=True, hide_index=True)
