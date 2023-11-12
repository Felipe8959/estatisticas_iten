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
meses_do_ano = {i: calendar.month_name[i] for i in range(1, 13)}
df_dados_mes = df_dados_mes.rename(index=meses_do_ano)

# Revisões
df_dados2 = pd.read_excel(arquivo_dados, sheet_name="revisoes")
df_dados2 = df_dados2.rename(index={0: 'Técnicos', 1: 'REP'})

df_dados2_mes = pd.read_excel(arquivo_dados, sheet_name="revisoes2")
meses_do_ano = {i: calendar.month_name[i] for i in range(1, 13)}
df_dados2_mes = df_dados2_mes.rename(index=meses_do_ano)

# Envios
df_dados3 = pd.read_excel(arquivo_dados, sheet_name="envios")
df_dados3 = df_dados3.rename(index={0: 'Técnicos', 1: 'REP'})

df_dados3_mes = pd.read_excel(arquivo_dados, sheet_name="envios2")
meses_do_ano = {i: calendar.month_name[i] for i in range(1, 13)}
df_dados3_mes = df_dados3_mes.rename(index=meses_do_ano)

# Erros
df_dados4 = pd.read_excel(arquivo_dados, sheet_name="erros")
df_dados4 = df_dados4.rename(index={0: 'Técnicos', 1: 'REP'})

df_dados4_mes = pd.read_excel(arquivo_dados, sheet_name="erros2")
meses_do_ano = {i: calendar.month_name[i] for i in range(1, 13)}
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
st.header("⏰ Atrasos")

# Criando colunas
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
                           labels={'value': f'Total de atrasos: {soma_total}'}, 
                           title=f'Controle de atrasos em {mes_atual} - Total:  {soma_total}'
                           )
    fig_bar_chart.update_xaxes(title_text='Setor/Laboratório')
    fig_bar_chart.update_traces(
        text=[f'{valor}\n{percentual}' for valor, percentual in zip(df_dados.values.flatten(), percentuais_formatados)],
        textposition='outside'
    )
    fig_bar_chart.update_layout(showlegend=False)
    st.plotly_chart(fig_bar_chart, use_container_width=True)

# Total de atrasos
with col2:
    fig_total_atrasos = px.line(df_dados_mes,
                                labels={'value': 'Total de atrasos'},
                                title=f'Controle de atrasos em {ano_atual}')
    fig_total_atrasos.update_xaxes(title_text='Mês')
    fig_total_atrasos.update_layout(legend_title_text='Legenda')
    st.plotly_chart(fig_total_atrasos, use_container_width=True)


st.header("🔧 Revisões")
col1, col2 = st.columns(2)

# Revisões
with col1:
    fig_bar_chart2 = px.bar(df_dados2.T, labels={'value': 'Total de revisões'}, title='Controle de revisões')
    fig_bar_chart2.update_xaxes(title_text='Setor')
    fig_bar_chart2.update_traces(text=df_dados2.values.flatten(), textposition='outside')
    fig_bar_chart2.update_layout(showlegend=False)
    st.plotly_chart(fig_bar_chart2, use_container_width=True)

# Total de revisões
with col2:
    fig_total_revisoes = px.line(df_dados2_mes, labels={'value': 'Total de revisões'}, title='Total de revisões 2023')
    fig_total_revisoes.update_xaxes(title_text='Mês')
    fig_total_revisoes.update_layout(legend_title_text='Legenda')
    st.plotly_chart(fig_total_revisoes, use_container_width=True)


st.header("✉️ Envios")
col1, col2 = st.columns(2)

# Envios
with col1:
    fig_bar_chart3 = px.bar(df_dados3.T, labels={'value': 'Total de envios'}, title='Controle de envios')
    fig_bar_chart3.update_xaxes(title_text='Setor')
    fig_bar_chart3.update_traces(text=df_dados3.values.flatten(), textposition='outside')
    fig_bar_chart3.update_layout(showlegend=False)
    st.plotly_chart(fig_bar_chart3, use_container_width=True)

# Total de Envios
with col2:
    fig_total_envios = px.line(df_dados3_mes, labels={'value': 'Total de envios'}, title='Total de envios 2023')
    fig_total_envios.update_xaxes(title_text='Mês')
    fig_total_envios.update_layout(legend_title_text='Legenda')
    st.plotly_chart(fig_total_envios, use_container_width=True)


st.header("❌ Erros")
col1, col2 = st.columns(2)

# Erros
with col1:
    fig_bar_chart4 = px.bar(df_dados4.T, labels={'value': 'Total de erros'}, title='Controle de erros')
    fig_bar_chart4.update_xaxes(title_text='Setor')
    fig_bar_chart4.update_traces(text=df_dados4.values.flatten(), textposition='outside')
    fig_bar_chart4.update_layout(showlegend=False)
    st.plotly_chart(fig_bar_chart4, use_container_width=True)

# Total de erros
with col2:
    fig_total_revisoes = px.line(df_dados4_mes, labels={'value': 'Total de erros'}, title='Total de erros 2023')
    fig_total_revisoes.update_xaxes(title_text='Mês')
    fig_total_revisoes.update_layout(showlegend=False)
    st.plotly_chart(fig_total_revisoes, use_container_width=True)