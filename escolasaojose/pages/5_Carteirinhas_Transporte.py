import streamlit as st
import pandas as pd
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

arquivo = st.file_uploader("PLANILHA DE TRANSPORTE ESCOLAR", type=["xlsx", "xls"])

if arquivo is not None:
    pandas = pd.DataFrame(pd.read_excel(
        arquivo, header=st.number_input("Ajuste ate o cabecalho estar certo", min_value=0,
                                        max_value=10))).dropna()
    pandas["Turma/Série"] = pandas["Série"] + " - " + pandas["Turma"]
    colunasRemover = st.multiselect("Remover Colunas",
                                    options=pandas.keys(),
                                    help="Escolha as colunas a remover")
    colunasAdicionadas = st.multiselect("Adicionar Coluna",
                                        options=["Observação", "Assinatura"],
                                        help="Escolha as colunas a adicionar")
    if colunasAdicionadas:
        pandas[colunasAdicionadas] = ""
    pandas = pandas.drop(colunasRemover, axis=1)


    @st.cache_data
    def baixarPlanilha(df, index=False):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine="xlsxwriter")
        df.to_excel(writer, index=index, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0'})
        worksheet.set_column('A:A', None, format1)
        writer.save()
        processed_data = output.getvalue()
        return processed_data


    planilha = baixarPlanilha(pandas)

    st.dataframe(pandas, use_container_width=True)
    st.download_button(label="📥 Baixar Lista", data=planilha, file_name='Lista Alunos Transporte.xlsx')
    st.dataframe(pandas.groupby(["Turma/Série"]).count(), use_container_width=True)
