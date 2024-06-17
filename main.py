import streamlit as st
import pandas as pd
from io import BytesIO
import datetime as dttm
import plotly.express as px
from plotly.offline import plot
from plotly.subplots import make_subplots

st.set_page_config(
    page_title="Format-Plan.",
    page_icon=":material/school:",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.title("Mesclar planilhas escolares")

arquivos = None
planilhas = []
with st.expander("Informações Iniciais"):
    nome_escola = st.text_input(
        ":school: **Digite o nome da escola**",
        help="""
        Digite o nome da escola para comecar, este será usado 
        como cabeçalho no(s) arquivo(s) a serem baixados.\n
        :red[O nome da escola será posto em maiusculo nos arquivos.]
        """,
        key="nome_escola",
        max_chars=100,
        placeholder="Nome completo da escola",
    )
    nome_escola = nome_escola.upper()
    if nome_escola:
        arquivos = st.file_uploader(
            label="**Planilhas de alunos**",
            accept_multiple_files=True,
            type=["XLSX", "XLS", "CSV"],
            help="""
            Aqui você poderá colocar planilha(s) contendo varias abas, 
            então você poderá escolher abaixo os ajustes delas para depois baixar 
            um arquivo final.\n
            :red[**TODAS DEVEM CONTER AS MESMAS COLUNAS E AJUSTES, 
             CASO CONTRARIO O RESULTADO FINAL NÃO SERÁ UTIL!**]
            """
        )

if arquivos:
    ajuste_plan, download_plan, graficos = st.tabs(["**Ajuste de planilha**", "**Baixar planilha**", "**Graficos**"])
    with ajuste_plan:
        st.subheader("Realize os ajustes na(s) planilha(s)")
        for arquivo in arquivos:
            st.subheader(
                f"Ajustes arquivo - :green[**{arquivo.name}**]",
                divider="green",
            )
            if arquivo.name.endswith(".csv"):
                chunks = []
                for chunk in pd.read_csv(arquivo, chunksize=10_000, sep=","):
                    chunks.append(
                        chunk
                    )
                planilhas.append(pd.DataFrame(pd.concat(chunks)))
            if arquivo.name.endswith(".xls") or arquivo.name.endswith(".xlsx"):
                abas_planilha_selecionadas = st.multiselect(
                    label=f"**Abas arquivo - :orange[{arquivo.name}]**",
                    placeholder="Selecione as abas que deseja mesclar",
                    options=list(pd.read_excel(arquivo, sheet_name=None).keys())
                )
                pular_linhas = st.number_input(
                    label=f"**Linhas a pular - :orange[{arquivo.name}]**",
                    placeholder="Ajuste o número até o cabeçalho da tabela ficar certo",
                    value=0,
                )
                for aba in abas_planilha_selecionadas:
                    df_temp = pd.DataFrame(
                        pd.read_excel(
                            arquivo,
                            sheet_name=aba,
                            skiprows=pular_linhas,
                        ),
                    ).dropna()

                    turma = str(aba).strip().replace("-", "").replace(" ", "")[1:]
                    df_temp["Turma"] = f"{turma[-3]}º Ano - {turma[-1]}"
                    df_temp["Turno"] = "Manhã" if f"{turma[-2]}" == "M" else "Tarde"
                    df_temp["Etapa"] = f"{turma[:-3]}"
                    df_temp["Etapa"] = pd.Categorical(
                        df_temp["Etapa"], categories=["EF9A", "NEM"]
                    )
                    df_temp[["Turma", "Turno", "Etapa"]] = df_temp[["Turma", "Turno", "Etapa"]].astype("category")
                    if "Matrícula" in df_temp.columns:
                        df_temp["Matrícula"] = pd.to_numeric(df_temp["Matrícula"].astype("int"), downcast="unsigned")
                    if "Nome" in df_temp.columns:
                        df_temp["Nome"] = df_temp["Nome"].astype(pd.StringDtype())
                    if "Data de nascimento" in df_temp.columns:
                        df_temp["Data de nascimento"] = df_temp["Data de nascimento"].astype(pd.StringDtype())
                    if "Nº de classe" in df_temp.columns:
                        df_temp["Nº de classe"] = pd.to_numeric(df_temp["Nº de classe"].astype("int"),
                                                                downcast="unsigned")
                    if "Sexo" in df_temp.columns:
                        df_temp["Sexo"] = df_temp["Sexo"].astype("category")
                    if "Situação" in df_temp.columns:
                        df_temp["Situação"] = df_temp["Situação"].astype("category")
                    if "Raça/Cor" in df_temp.columns:
                        df_temp["Raça/Cor"] = df_temp["Raça/Cor"].astype("category")
                    if "Procedência modalidade/curso" in df_temp.columns:
                        df_temp["Procedência modalidade/curso"] = df_temp["Procedência modalidade/curso"].astype(
                            "category")
                    planilhas.append(df_temp)
        st.divider()
        df = pd.concat(planilhas) if len(planilhas) > 0 else pd.DataFrame(data={"N/A": ["Ajuste as informações"]})
        if "Etapa" in df.columns:
            if "Turma" in df.columns:
                if "Nome" in df.columns:
                    df = df.sort_values(by=["Etapa", "Turma", "Nome"])
        if "Matrícula" in df.columns:
            df = df.set_index("Matrícula")
        df2 = df.copy()

        st.subheader("Adição e Remoção de colunas")
        adicionar, remover = st.columns([0.5, 0.5])
        with adicionar:
            colunas_adicionar = st.multiselect(
                label=":green-background[**Colunas para adicionar**]",
                options=[
                    "Observação", "Assinatura",
                    "Nota - 1", "Nota - 2", "Nota - 3", "Nota - 4",
                    "Prova", "Média",
                ],
                placeholder="Escolhas aquelas que serão necessarias",
            )
            for col_add in colunas_adicionar:
                df[col_add] = ["" for x in range(len(df))]
        with remover:
            colunas_remover = st.multiselect(
                label=":red-background[**Colunas para remover**]",
                options=df.columns,
                placeholder="Escolhas aquelas que não serão necessarias",
                max_selections=len(df.columns) - 1
            )
            df = df.drop(columns=colunas_remover, axis=1)
        st.subheader("Visualizar Tabela Final")
        with st.expander("Tabela", expanded=False):
            st.dataframe(
                df,
                use_container_width=True,
            )
    with download_plan:
        st.write("Escolhas o que deseja baixar")
        if "Turma" in df2.columns:
            alunos_por_turma = df2.groupby(by=["Turma"]).count()["Nome"]
        else:
            alunos_por_turma = df2


        @st.cache_data
        def baixarPlanilha(dataframe):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine="xlsxwriter")
            multiindex = pd.MultiIndex.from_product(
                [[nome_escola], list(dataframe.columns)],
            )
            dataframe = pd.DataFrame(
                columns=multiindex,
                data=dataframe.values,
            )
            dataframe.to_excel(writer, sheet_name="Inicio", index=True)
            alunos_por_turma.to_excel(writer, sheet_name="Alunos por Turma")
            writer.close()
            processed_data = output.getvalue()
            return processed_data


        planilha = baixarPlanilha(df)
        st.download_button("📥 Baixar Lista de alunos das turmas selecionadas", data=planilha,
                           file_name=f"Lista do Alunos Completa - {dttm.datetime.today().strftime('%d-%m-%Y')}.xlsx")
    with graficos:
        try:
            st.subheader(f"Total de alunos: {df2['Nome'].count()}")
            df_graficos = df2
            df_graficos["Quantidade"] = ''

            st.plotly_chart(
                FIG1 := px.bar(
                    data_frame=df_graficos.groupby(["Turma"]).count()["Quantidade"].reset_index(),
                    x="Turma",
                    y="Quantidade",
                    text_auto=True,
                    title="Distribuição de alunos por Turma",
                    color_discrete_sequence=[
                        '#636EFA', '#EF553B', '#00CC96', '#AB63FA',
                        '#FFA15A', '#19D3F3', '#FF6692', '#B6E880',
                        '#FF97FF', '#FECB52'
                    ],
                ),
            )
            st.download_button(
                "Baixar imagem do grafico acima",
                data=FIG1.to_image(format="png"),
                file_name="Total alunos por Turma.png",
                mime="image/png"
            )

            st.plotly_chart(
                FIG2 := px.bar(
                    data_frame=df_graficos.groupby(["Turma", "Sexo"]).count()["Quantidade"].reset_index(),
                    x="Turma",
                    y="Quantidade",
                    color="Sexo",
                    text_auto=True,
                    title="Distribuição de alunos por Turma e Sexo",
                    color_discrete_sequence=[
                        '#636EFA', '#EF553B', '#00CC96', '#AB63FA',
                        '#FFA15A', '#19D3F3', '#FF6692', '#B6E880',
                        '#FF97FF', '#FECB52'
                    ]
                ),
            )
            st.download_button(
                "Baixar imagem do grafico acima",
                data=FIG2.to_image(format="png"),
                file_name="Total alunos por Turma e Sexo.png",
                mime="image/png"
            )

            st.plotly_chart(
                FIG3 := px.bar(
                    data_frame=df_graficos.groupby(["Raça/Cor"]).count()["Quantidade"].reset_index(),
                    x="Raça/Cor",
                    y="Quantidade",
                    text_auto=True,
                    title="Distribuição de alunos por Raça/Cor",
                    color_discrete_sequence=[
                        '#636EFA', '#EF553B', '#00CC96', '#AB63FA',
                        '#FFA15A', '#19D3F3', '#FF6692', '#B6E880',
                        '#FF97FF', '#FECB52'
                    ]
                ),
            )
            st.download_button(
                "Baixar imagem do grafico acima",
                data=FIG3.to_image(format="png"),
                file_name="Total alunos por Raça/Cor.png",
                mime="image/png"
            )

            st.plotly_chart(
                FIG4 := px.bar(
                    data_frame=df_graficos.groupby(["Raça/Cor", "Turma"]).count()["Quantidade"].reset_index(),
                    x="Turma",
                    y="Quantidade",
                    color="Raça/Cor",
                    text_auto=True,
                    title="Distribuição de alunos por Turma e Raça/Cor",
                    color_discrete_sequence=[
                        '#636EFA', '#EF553B', '#00CC96', '#AB63FA',
                        '#FFA15A', '#19D3F3', '#FF6692', '#B6E880',
                        '#FF97FF', '#FECB52'
                    ]
                ),
            )
            st.download_button(
                "Baixar imagem do grafico acima",
                data=FIG4.to_image(format="png"),
                file_name="Total alunos por Turma e Raça/Cor.png",
                mime="image/png"
            )

            st.plotly_chart(
                FIG5 := px.bar(
                    data_frame=df_graficos.groupby(["Turno"]).count()["Quantidade"].reset_index(),
                    x="Turno",
                    y="Quantidade",
                    text_auto=True,
                    title="Distribuição de alunos por Turno",
                    color_discrete_sequence=[
                        '#636EFA', '#EF553B', '#00CC96', '#AB63FA',
                        '#FFA15A', '#19D3F3', '#FF6692', '#B6E880',
                        '#FF97FF', '#FECB52'
                    ]
                ),
            )
            st.download_button(
                "Baixar imagem do grafico acima",
                data=FIG5.to_image(format="png"),
                file_name="Total alunos por Turno.png",
                mime="image/png"
            )
            st.divider()
            figuras = [FIG1, FIG2, FIG3, FIG4, FIG5]
            FIG = make_subplots(
                rows=len(figuras),
                cols=1,
                subplot_titles=("Turma", "Turma e Sexo", "Raça/Cor", "Turma e Raça/Cor", "Turno"),

            )
            for i, figure in enumerate(figuras):
                for trace in range(len(figure["data"])):
                    FIG.append_trace(figure["data"][trace], row=i + 1, col=1)
            FIG.update_layout(
                height=2500,
                width=1250,
                title_text="Quantidade de alunos",
                showlegend=False,
                coloraxis=dict(colorscale='Bluered_r')
            )

            st.download_button(
                "Baixar Imagem dos Graficos",
                data=FIG.to_image(format="png"),
                file_name="Grafico subplot.png",
                mime="image/pdf"
            )
            st.download_button(
                "Baixar PDF dos Graficos",
                data=FIG.to_image(format="pdf"),
                file_name="Grafico subplot.pdf",
                mime="application/pdf"
            )

        except Exception as erro:
            st.write("Algo deu errado, realize os ajustes na tabela.")
