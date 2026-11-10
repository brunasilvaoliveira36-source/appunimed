import streamlit as st
import pandas as pd
import openpyxl

from datetime import datetime
from pathlib import Path

# =====================================================
# CONFIGURAÇÕES
# =====================================================

st.set_page_config(
    page_title="Hotelaria Unimed",
    page_icon="🏥",
    layout="wide"
)

ABA_EXCEL = "Camareiras"

COLUNAS = [
    "Data",
    "Cargo",
    "Colaborador",
    "Período",
    "Tipo de Entrega",
    "Andar",
    "Leito",
    "Enxoval",
    "Tamanho",
    "Quantidade",
    "Observação"
]

ITENS_COM_TAMANHO = [
    "Camisola",
    "Pijama",
    "Short"
]

LISTA_ENXOVAIS = [
    "Camisola",
    "Pijama",
    "Short",
    "Traçado",
    "Impermeável",
    "Toalha Ped",
    "Toalha Adulto",
    "Kit Higiene",
    "Manta",
    "Manta Ped",
    "Cueiro",
    "Fronha Janequini"
]

TAMANHOS = [
    "P",
    "M",
    "G",
    "GG",
    "3G",
    "4G"
]

# =====================================================
# CABEÇALHO
# =====================================================

st.title("🏥 Hotelaria Unimed")
st.subheader("Sistema de Gestão de Enxovais Hospitalares")

# =====================================================
# PLANILHA
# =====================================================

arquivo = st.file_uploader(
    "Selecione a planilha Excel",
    type=["xlsx"]
)

if arquivo:

    # =================================================
    # ABAS
    # =================================================

    cadastro_tab, historico_tab = st.tabs(
        ["📝 Cadastro", "📋 Histórico"]
    )

    # =================================================
    # CADASTRO
    # =================================================

    with cadastro_tab:

        st.header("Cadastro de Entrega")

        col1, col2 = st.columns(2)

        with col1:

            data = st.date_input(
                "Data",
                datetime.today()
            )

            cargo = st.selectbox(
                "Cargo",
                ["Camareira"]
            )

            colaborador = st.text_input(
                "Colaborador"
            )

            periodo = st.selectbox(
                "Período",
                ["Manhã", "Tarde", "Noite"]
            )

        with col2:

            tipo_entrega = st.selectbox(
                "Tipo de Entrega",
                ["Extra", "Enxoval Normal"]
            )

            andar = st.selectbox(
                "Andar",
                ["1", "2", "3", "4", "5", "6", "7"]
            )

            leito = st.text_input(
                "Leito"
            )

            observacao = st.text_input(
                "Observação"
            )

        st.divider()

        st.subheader("Enxovais")

        itens_registro = []

        for item in LISTA_ENXOVAIS:

            col_a, col_b, col_c = st.columns([3, 1, 1])

            with col_a:
                marcado = st.checkbox(
                    item,
                    key=f"chk_{item}"
                )

            with col_b:

                quantidade = st.number_input(
                    "Qtd",
                    min_value=0,
                    value=0,
                    key=f"qtd_{item}"
                )

            with col_c:

                if item in ITENS_COM_TAMANHO:

                    tamanho = st.selectbox(
                        "Tam",
                        [""] + TAMANHOS,
                        key=f"tam_{item}"
                    )

                else:
                    tamanho = ""

            if marcado and quantidade > 0:

                itens_registro.append({
                    "Enxoval": item,
                    "Tamanho": tamanho,
                    "Quantidade": quantidade
                })

        st.divider()

        if st.button(
            "✅ Salvar Entrega",
            use_container_width=True
        ):

            if not colaborador:

                st.error(
                    "Informe o colaborador."
                )

            elif not leito:

                st.error(
                    "Informe o leito."
                )

            elif len(itens_registro) == 0:

                st.error(
                    "Selecione pelo menos um enxoval."
                )

            else:

                workbook = openpyxl.load_workbook(
                    arquivo
                )

                ws = workbook[ABA_EXCEL]

                ultima_linha = ws.max_row + 1

                for item in itens_registro:

                    ws.cell(
                        ultima_linha,
                        2
                    ).value = data

                    ws.cell(
                        ultima_linha,
                        3
                    ).value = cargo

                    ws.cell(
                        ultima_linha,
                        4
                    ).value = colaborador

                    ws.cell(
                        ultima_linha,
                        5
                    ).value = periodo

                    ws.cell(
                        ultima_linha,
                        6
                    ).value = tipo_entrega

                    ws.cell(
                        ultima_linha,
                        7
                    ).value = andar

                    ws.cell(
                        ultima_linha,
                        8
                    ).value = leito

                    ws.cell(
                        ultima_linha,
                        9
                    ).value = item["Enxoval"]

                    ws.cell(
                        ultima_linha,
                        10
                    ).value = item["Tamanho"]

                    ws.cell(
                        ultima_linha,
                        11
                    ).value = item["Quantidade"]

                    ws.cell(
                        ultima_linha,
                        12
                    ).value = observacao

                    ultima_linha += 1

                nome_saida = "enxovais_atualizado.xlsx"

                workbook.save(nome_saida)

                workbook.close()

                st.success(
                    "Registro salvo com sucesso."
                )

                with open(
                    nome_saida,
                    "rb"
                ) as file:

                    st.download_button(
                        "⬇ Baixar planilha atualizada",
                        file,
                        file_name=nome_saida
                    )

    # =================================================
    # HISTÓRICO
    # =================================================

    with historico_tab:

        st.header(
            "Histórico de Entregas"
        )

        workbook = openpyxl.load_workbook(
            arquivo,
            data_only=True
        )

        ws = workbook[ABA_EXCEL]

        dados = []

        for row in ws.iter_rows(
            min_row=2,
            values_only=True
        ):

            linha = list(row)

            if len(linha) >= 12:

                dados.append(
                    linha[1:12]
                )

        workbook.close()

        if dados:

            df = pd.DataFrame(
                dados,
                columns=COLUNAS
            )

            pesquisa = st.text_input(
                "Pesquisar"
            )

            if pesquisa:

                filtro = df.astype(
                    str
                ).apply(
                    lambda x: x.str.contains(
                        pesquisa,
                        case=False
                    )
                )

                df = df[
                    filtro.any(axis=1)
                ]

            st.dataframe(
                df,
                use_container_width=True,
                height=600
            )

            st.metric(
                "Registros encontrados",
                len(df)
            )

        else:

            st.warning(
                "Nenhum registro encontrado."
            )

else:

    st.info(
        "Selecione uma planilha Excel para iniciar."
    )