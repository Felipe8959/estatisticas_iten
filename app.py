import streamlit as st
import pandas as pd
import calendar
import plotly.express as px
import locale
import numpy as np
locale.setlocale(locale.LC_TIME, "pt_BR")

st.set_page_config(page_title="Controle REP", page_icon="✅", layout="wide", initial_sidebar_state="expanded")

arquivo = "relordemservicogeral.xls"

# Lê o arquivo Excel com a linha 6 como cabeçalho
df = pd.read_excel(arquivo, header=6)

# Exclui linhas onde o conteúdo da terceira coluna está vazio
df = df.dropna(subset=[df.columns[2]])
# Exclui colunas onde o conteúdo da sexta linha esteja vazio
df = df.dropna(axis=1, subset=[5])

date_format = df.columns[8:12]
df[date_format] = df[date_format].apply(pd.to_datetime, errors='coerce')

# Leitura do arquivo 'dados.xlsx'
arquivo_dados = "dados.xlsx"


# Atrasos
df_dados = pd.read_excel(arquivo_dados, sheet_name="atrasos")
df_dados = df_dados.rename(index={0: 'Técnicos', 1: 'REP'})

df_dados_mes = pd.read_excel(arquivo_dados, sheet_name="atrasos2")
meses_do_ano = {i: calendar.month_name[i] for i in range(0, 12)}
df_dados_mes = df_dados_mes.rename(index=meses_do_ano)

# Revisões
df_dados2 = pd.read_excel(arquivo_dados, sheet_name="revisoes")
df_dados2 = df_dados2.rename(index={0: 'Técnicos', 1: 'REP'})

df_dados2_mes = pd.read_excel(arquivo_dados, sheet_name="revisoes2")
meses_do_ano = {i: calendar.month_name[i] for i in range(0, 12)}
df_dados2_mes = df_dados2_mes.rename(index=meses_do_ano)

# Envios
df_dados3 = pd.read_excel(arquivo_dados, sheet_name="envios")
df_dados3 = df_dados3.rename(index={0: 'Técnicos', 1: 'REP'})

df_dados3_mes = pd.read_excel(arquivo_dados, sheet_name="envios2")
meses_do_ano = {i: calendar.month_name[i] for i in range(0, 12)}
df_dados3_mes = df_dados3_mes.rename(index=meses_do_ano)

# Erros
df_dados4 = pd.read_excel(arquivo_dados, sheet_name="erros")
df_dados4 = df_dados4.rename(index={0: 'Técnicos', 1: 'REP'})

df_dados4_mes = pd.read_excel(arquivo_dados, sheet_name="erros2")
meses_do_ano = {i: calendar.month_name[i] for i in range(0, 12)}
df_dados4_mes = df_dados4_mes.rename(index=meses_do_ano)



# Menu de filtros
# st.sidebar.header("Filtros")
#selected_column = st.sidebar.selectbox("Escolha uma coluna para filtro", df.columns)
#selected_value = st.sidebar.text_input(f"Filtrar por valor em '{selected_column}'", "")

# Filtros
#if selected_value:
    #df = df[df[selected_column] == selected_value]

mes_atual = 'Junho'
ano_atual = '2023'

st.markdown(
    f"<h2 style='text-align: center;'>Relatório mensal ITEN ({mes_atual}/{ano_atual})</h2>",
    unsafe_allow_html=True
    )

# Dados
#st.header("Dados Filtrados")
#st.write(df)



# ------- Gráficos ------- #

st.header("🕐 Atrasos")

col1, col2 = st.columns(2)

# Calcular estatísticas
maior_valor = df_dados.values.max()
menor_valor = df_dados.values.min()
soma_total = df_dados.values.sum()
percentuais = (df_dados.values / soma_total) * 100
percentuais_formatados = [f'({percentual:.0f}%)' for percentual in percentuais.flatten()]

# Atrasos
with col1:
    fig_bar_chart = px.bar(df_dados.T,
                           labels={'value': 'Qtd de atrasos'}, 
                           title=f'Controle de atrasos em {mes_atual} - Total: {soma_total}'
                           )
    fig_bar_chart.update_xaxes(title_text='Setor/Laboratório')
    fig_bar_chart.update_traces(
        text=[f'{valor}\n{percentual}' for valor, percentual in zip(df_dados.values.flatten(), percentuais_formatados)],
        textposition='outside'
    )
    fig_bar_chart.update_layout(showlegend=False)
    st.plotly_chart(fig_bar_chart, use_container_width=True)


with col2:
    # Total de atrasos
    fig_total_atrasos = px.line(df_dados_mes,
                                labels={'value': 'Qtd de atrasos'},
                                title=f'Controle de atrasos em {ano_atual}')
    fig_total_atrasos.update_xaxes(title_text='Mês')
    fig_total_atrasos.update_layout(legend_title_text='Legenda')
    st.plotly_chart(fig_total_atrasos, use_container_width=True)



st.header("🔧 Revisões")

col1, col2 = st.columns(2)

# Calcular estatísticas
maior_valor2 = df_dados2.values.max()
menor_valor2 = df_dados2.values.min()
soma_total2 = df_dados2.values.sum()
percentuais2 = (df_dados2.values / soma_total2) * 100
percentuais_formatados2 = [f'({percentual:.0f}%)' for percentual in percentuais2.flatten()]

# Revisões
with col1:
    fig_bar_chart2 = px.bar(df_dados2.T, labels={'value': 'Qtd de revisões'}, 
                            title=f'Controle de revisões em {mes_atual} - Total: {soma_total2}'
                            )
    fig_bar_chart2.update_xaxes(title_text='Setor')
    fig_bar_chart2.update_traces(text=[f'{valor}\n{percentual}' for valor, 
                                       percentual in zip (df_dados2.values.flatten(), 
                                                          percentuais_formatados2)], textposition='outside')
    fig_bar_chart2.update_layout(showlegend=False)
    st.plotly_chart(fig_bar_chart2, use_container_width=True)

# Total de revisões
with col2:
    fig_total_revisoes = px.line(df_dados2_mes,
                                 labels={'value': 'Qtd de revisões'},
                                 title=f'Controle de revisões em {ano_atual}')
    fig_total_revisoes.update_xaxes(title_text='Mês')
    fig_total_revisoes.update_layout(legend_title_text='Legenda')
    st.plotly_chart(fig_total_revisoes, use_container_width=True)


st.header("✉️ Envios")

col1, col2 = st.columns(2)

# Calcular estatísticas
maior_valor3 = df_dados3.values.max()
menor_valor3 = df_dados3.values.min()
soma_total3 = df_dados3.values.sum()
percentuais3 = (df_dados3.values / soma_total3) * 100
percentuais_formatados3 = [f'({percentual:.0f}%)' for percentual in percentuais3.flatten()]

# Envios
with col1:
    fig_bar_chart3 = px.bar(df_dados3.T,
                            labels={'value': 'Qtd de envios'}, 
                            title=f'Controle de envios em {mes_atual} - Total: {soma_total3}')
    fig_bar_chart3.update_xaxes(title_text='Setor')
    fig_bar_chart3.update_traces(text=[f'{valor}\n{percentual}' for valor,
                                       percentual in zip(df_dados3.values.flatten(),
                                                         percentuais_formatados3)], textposition='outside')
    fig_bar_chart3.update_layout(showlegend=False)
    st.plotly_chart(fig_bar_chart3, use_container_width=True)

# Total de Envios
with col2:
    fig_total_envios = px.line(df_dados3_mes, labels={'value': 'Total de envios'}, title=f'Controle de envios 2023')
    fig_total_envios.update_xaxes(title_text='Mês')
    fig_total_envios.update_layout(legend_title_text='Legenda')
    st.plotly_chart(fig_total_envios, use_container_width=True)


st.header("❌ Erros")
col1, col2 = st.columns(2)

maior_valor4 = df_dados4.values.max()
menor_valor4 = df_dados4.values.min()
soma_total4 = df_dados4.values.sum()
soma_total4_mes = int(abs(df_dados4_mes.sum()))
percentuais4 = (df_dados4.values / soma_total4) * 100
percentuais_formatados4 = [f'({percentual:.0f}%)' for percentual in percentuais4.flatten()]

# Erros
with col1:
    fig_bar_chart4 = px.bar(df_dados4.T,
                            labels={'value': 'Qtd de erros'},
                            title=f'Controle de erros em {mes_atual} - Total: {soma_total4}')
    fig_bar_chart4.update_xaxes(title_text='Setor')
    fig_bar_chart4.update_traces(text=[f'{valor}\n{percentual}' for valor, percentual in zip(df_dados4.values.flatten(), percentuais_formatados4)], textposition='outside')
    fig_bar_chart4.update_layout(showlegend=False)
    st.plotly_chart(fig_bar_chart4, use_container_width=True)

# Total de erros
with col2:
    fig_total_erros = px.line(df_dados4_mes, labels={'value': 'Qtd de erros'}, title=f'Controle de erros 2023 Total: {soma_total4_mes}')
    fig_total_erros.update_xaxes(title_text='Mês')
    fig_total_erros.update_layout(showlegend=False)
    st.plotly_chart(fig_total_erros, use_container_width=True)



# ------- Comparação entre dois meses ------- #

st.header(f"📉 Comparação entre dois meses")

# Lista de todos os meses do ano
all_months = list(calendar.month_name)[1:]

# Adicione uma opção para selecionar o tipo de comparação
selected_comparison_type = st.selectbox(
    "Selecione o tipo de comparação:",
    ["", "Atrasos", "Revisões", "Envios", "Erros"]
)

try:
    # Verifica se um tipo de comparação foi selecionado
    if selected_comparison_type:
        # Adicione uma opção para selecionar dois meses
        selected_months = st.multiselect(
            "Selecione dois meses para comparação:",
            all_months,
            default=[], # Inicia vazio
            key='selected_months'  # Adiciona uma chave para identificar o componente
        )

        if len(selected_months) > 2:
            st.warning("Selecione apenas dois meses para comparação.")
            selected_months = selected_months[:2]  # Pega apenas os dois primeiros meses

        # Verifica se pelo menos dois meses foram selecionados
        elif len(selected_months) == 2:
            # Filtra os dados para os meses e o tipo de comparação selecionados
            if selected_comparison_type == "Atrasos":
                df_selected_data = df_dados_mes.loc[df_dados_mes.index.isin(selected_months)]
                selected_columns = df_dados_mes.columns
            elif selected_comparison_type == "Revisões":
                df_selected_data = df_dados2_mes.loc[df_dados_mes.index.isin(selected_months)]
                selected_columns = df_dados2_mes.columns
            elif selected_comparison_type == "Envios":
                df_selected_data = df_dados3_mes.loc[df_dados_mes.index.isin(selected_months)]
                selected_columns = df_dados3_mes.columns
            elif selected_comparison_type == "Erros":
                df_selected_data = df_dados4_mes.loc[df_dados_mes.index.isin(selected_months)]
                selected_columns = df_dados4_mes.columns
            # Verifica se o DataFrame resultante está vazio

            st.markdown(f"<h1 style='font-size: 24px'>Comparação de {selected_comparison_type} entre {selected_months[1]} e {selected_months[0]}</h1>", unsafe_allow_html=True)

            # Calcula a diferença entre os meses selecionados
            aumento_ou_queda = "aumento ⬆" if df_selected_data.iloc[1, 0] < df_selected_data.iloc[0, 0] else "queda ⬇"
                
            diferenca_dados = df_selected_data.diff().iloc[1]

            diferenca_total = int(abs(diferenca_dados.sum()))

            # Adiciona formatação de cor ao texto 'queda' e 'aumento'
            if selected_comparison_type == 'Envios':
                color = 'green' if aumento_ou_queda == 'aumento ⬆' else 'red'
            else:
                color = 'red' if aumento_ou_queda == 'aumento ⬆' else 'green'
            styled_text = f"<span style='color: {color};'>{aumento_ou_queda}</span>"

            # Exibe o texto de comparação total
            st.markdown(f"<strong>Total:</strong> Houve {styled_text} de {diferenca_total} {selected_comparison_type.lower()} entre {selected_months[0]} e {selected_months[1]}.", unsafe_allow_html=True)

            # Exibe o texto de comparação para cada coluna
            for coluna in selected_columns:
                aumento_ou_queda = "aumento ⬆" if df_selected_data.loc[selected_months[1], coluna] < df_selected_data.loc[selected_months[0], coluna] else "queda ⬇"
                valor_diferenca = int(abs(diferenca_dados[coluna]))

                # Adiciona formatação de cor ao texto 'queda' e 'aumento'
                if selected_comparison_type == 'Envios':
                    color = 'green' if aumento_ou_queda == 'aumento ⬆' else 'red'
                else:
                    color = 'red' if aumento_ou_queda == 'aumento ⬆' else 'green'
                styled_text = f"<span style='color: {color};'>{aumento_ou_queda}</span>"

                st.markdown(f"<strong>{coluna}</strong>: Houve {styled_text} de {valor_diferenca} {selected_comparison_type.lower()} entre {selected_months[0]} e {selected_months[1]}.", unsafe_allow_html=True)

            # Exibe um gráfico de linha para a comparação
            fig_comparacao = px.line(df_selected_data, labels={'value': f'Qtd de {selected_comparison_type.lower()}'}, title=f'Total de {selected_comparison_type.lower()} por mês')
            fig_comparacao.update_xaxes(title_text='Mês')
            fig_comparacao.update_layout(legend_title_text='Legenda')
            st.plotly_chart(fig_comparacao, use_container_width=True)
                
        elif len(selected_months) == 1:
            st.warning("Selecione duas colunas para comparação.")
        else:
            st.info("Selecione pelo menos uma coluna para comparação.")
    else:
        st.info("Selecione um tipo de comparação.")
except KeyError as e:
    st.warning(f"Erro ao acessar dados: {e}")