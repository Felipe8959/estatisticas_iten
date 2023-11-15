import streamlit as st
import pandas as pd
import calendar
import plotly.express as px
import locale
import numpy as np

locale.setlocale(locale.LC_TIME, "pt_BR")

st.set_page_config(page_title="Controle REP", page_icon="✅", initial_sidebar_state="expanded")

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
meses_do_ano = {i: calendar.month_name[i] for i in range(0, 13)}
df_dados_mes = df_dados_mes.rename(index=meses_do_ano)

# Revisões
df_dados2 = pd.read_excel(arquivo_dados, sheet_name="revisoes")
df_dados2 = df_dados2.rename(index={0: 'Técnicos', 1: 'REP'})

df_dados2_mes = pd.read_excel(arquivo_dados, sheet_name="revisoes2")
meses_do_ano = {i: calendar.month_name[i] for i in range(0, 13)}
df_dados2_mes = df_dados2_mes.rename(index=meses_do_ano)

# Envios
df_dados3 = pd.read_excel(arquivo_dados, sheet_name="envios")
df_dados3 = df_dados3.rename(index={0: 'Técnicos', 1: 'REP'})

df_dados3_mes = pd.read_excel(arquivo_dados, sheet_name="envios2")
meses_do_ano = {i: calendar.month_name[i] for i in range(0, 13)}
df_dados3_mes = df_dados3_mes.rename(index=meses_do_ano)

# Erros
df_dados4 = pd.read_excel(arquivo_dados, sheet_name="erros")
df_dados4 = df_dados4.rename(index={0: 'Técnicos', 1: 'REP'})

df_dados4_mes = pd.read_excel(arquivo_dados, sheet_name="erros2")
meses_do_ano = {i: calendar.month_name[i] for i in range(0, 13)}
df_dados4_mes = df_dados4_mes.rename(index=meses_do_ano)


#--------------- Definir mês e ano ---------------#

df_info = pd.read_excel(arquivo_dados, sheet_name='info', header=None)

indice_mes = df_info[df_info.eq("#mes").any(axis=1)].index
indice_ano = df_info[df_info.eq("#ano").any(axis=1)].index

# Verificar se '#mes' foi encontrado
if not indice_mes.empty:
    # Obter o valor da célula abaixo de '#mes'
    mes_atual = df_info.iloc[indice_mes[0] + 1, 0]

    # Verificar se o valor é um número inteiro
    if isinstance(mes_atual, (int, float)):
        mes_atual = int(mes_atual)

        # Verificar se o número do mês está dentro do intervalo válido (1 a 12)
        if 1 <= mes_atual <= 12:
            mes_atual = calendar.month_name[mes_atual]
        else:
            print(f"Erro: Número do mês fora do intervalo válido (1 a 12): {mes_atual}")
    else:
        print(f"Erro: Valor abaixo de '#mes' não é um número: {mes_atual}")
else:
    print("#mes não encontrado na aba 'info'")


# Busca pelo valor correspondente a '#ano'
indice_ano = df_info[df_info.eq("#ano").any(axis=1)].index

# Verificar se '#ano' foi encontrado
if not indice_ano.empty:
    # Obter o valor da célula abaixo de '#ano'
    ano_atual = df_info.iloc[indice_ano[0] + 1, 1]

    # Converter para inteiro se possível
    try:
        ano_atual = int(ano_atual)
    except ValueError:
        print(f"Erro: Valor abaixo de '#ano' não é um número: {ano_atual}")
else:
    print("#ano não encontrado na aba 'info'")


st.markdown(
    f"<h2 style='text-align: center; font-size: 24px'>Relatório mensal ITEN ({mes_atual}/{ano_atual})</h2>",
    unsafe_allow_html=True
    )


#----------------------------------------#
#--------------- Gráficos ---------------#
#----------------------------------------#

#st.header("")
st.markdown(f"<h3 style='font-size: 24px'>🕐 Controle mensal</h3>", unsafe_allow_html=True)

# Calcular estatísticas
maior_valor = df_dados.values.max()
menor_valor = df_dados.values.min()
soma_total = df_dados.values.sum()
soma_total_mes = int(df_dados_mes.sum().sum())

percentuais = (df_dados.values / soma_total) * 100
percentuais_formatados = [f'({percentual:.0f}%)' for percentual in percentuais.flatten()]

# Atrasos
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


# Total de atrasos
fig_total_atrasos = px.line(df_dados_mes,
                            labels={'value': 'Qtd de atrasos'},
                            title=f'Controle de atrasos em {ano_atual} - Total: {soma_total_mes}')
fig_total_atrasos.update_xaxes(title_text='Mês')
fig_total_atrasos.update_layout(legend_title_text='Legenda')
st.plotly_chart(fig_total_atrasos, use_container_width=True)



#st.header("🔧 Revisões")
st.markdown(f"<h3 style='font-size: 24px'>🔧 Revisões</h3>", unsafe_allow_html=True)

# Calcular estatísticas
maior_valor2 = df_dados2.values.max()
menor_valor2 = df_dados2.values.min()
soma_total2 = df_dados2.values.sum()
soma_total2_mes = int(df_dados2_mes.sum().sum())
percentuais2 = (df_dados2.values / soma_total2) * 100
percentuais_formatados2 = [f'({percentual:.0f}%)' for percentual in percentuais2.flatten()]

# Revisões
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
fig_total_revisoes = px.line(df_dados2_mes,
                                labels={'value': 'Qtd de revisões'},
                                title=f'Controle de revisões em {ano_atual} - Total: {soma_total2_mes}')
fig_total_revisoes.update_xaxes(title_text='Mês')
fig_total_revisoes.update_layout(legend_title_text='Legenda')
st.plotly_chart(fig_total_revisoes, use_container_width=True)


#st.header("✉️ Envios")
st.markdown(f"<h3 style='font-size: 24px'>✉️ Envios</h3>", unsafe_allow_html=True)

# Calcular estatísticas
maior_valor3 = df_dados3.values.max()
menor_valor3 = df_dados3.values.min()
soma_total3 = df_dados3.values.sum()
soma_total3_mes = int(df_dados3_mes.sum().sum())
percentuais3 = (df_dados3.values / soma_total3) * 100
percentuais_formatados3 = [f'({percentual:.0f}%)' for percentual in percentuais3.flatten()]

# Envios
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
fig_total_envios = px.line(df_dados3_mes,
                            labels={'value': 'Qtd de envios'},
                            title=f'Controle de envios {ano_atual} - Total: {soma_total3_mes}')
fig_total_envios.update_xaxes(title_text='Mês')
fig_total_envios.update_layout(legend_title_text='Legenda')
st.plotly_chart(fig_total_envios, use_container_width=True)


#st.header("❌ Erros")
st.markdown(f"<h3 style='font-size: 24px'>❌ Erros</h3>", unsafe_allow_html=True)

col1, col2 = st.columns(2)

maior_valor4 = df_dados4.values.max()
menor_valor4 = df_dados4.values.min()
soma_total4 = df_dados4.values.sum()
soma_total4_mes = int(df_dados4_mes.sum().sum())
percentuais4 = (df_dados4.values / soma_total4) * 100
percentuais_formatados4 = [f'({percentual:.0f}%)' for percentual in percentuais4.flatten()]

# Erros
fig_bar_chart4 = px.bar(df_dados4.T,
                        labels={'value': 'Qtd de erros'},
                        title=f'Controle de erros em {mes_atual} - Total: {soma_total4}')
fig_bar_chart4.update_xaxes(title_text='Setor')
fig_bar_chart4.update_traces(text=[f'{valor}\n{percentual}' for valor, percentual in zip(df_dados4.values.flatten(), percentuais_formatados4)], textposition='outside')
fig_bar_chart4.update_layout(showlegend=False)
st.plotly_chart(fig_bar_chart4, use_container_width=True)

# Total de erros
fig_total_revisoes = px.line(df_dados4_mes,
                            labels={'value': 'Qtd de erros'},
                            title=f'Controle de erros {ano_atual} - Total: {soma_total4_mes}')
fig_total_revisoes.update_xaxes(title_text='Mês')
fig_total_revisoes.update_layout(showlegend=False)
st.plotly_chart(fig_total_revisoes, use_container_width=True)



# ------- Comparação entre dois meses ------- #

st.header(f"📉 Desempenho")
#st.markdown(f"<h3 style='font-size: 24px'>📉 Desempenho</h3>", unsafe_allow_html=True)

# Lista de todos os meses do ano
all_months = list(calendar.month_name)[1:] + ['média anual']

# Adicione uma opção para selecionar o tipo de comparação
selected_comparison_type = st.selectbox(
    "Selecione o tipo de comparação:",
    ["Atrasos", "Revisões", "Envios", "Erros"]
)

try:
    # Verifica se um tipo de comparação foi selecionado
    if selected_comparison_type:
        # Adicione uma opção para selecionar dois meses
        selected_months = st.multiselect(
            "Selecione dois meses para comparação:",
            all_months,
            default=['média anual', mes_atual],
            key='selected_months'  # Adiciona uma chave para identificar o componente
        )

        if len(selected_months) > 2:
            st.warning("Selecione apenas dois meses para comparação.")
            selected_months = selected_months.sorted()[:2]  # Pega apenas os dois primeiros meses

        # Verifica se pelo menos dois meses foram selecionados
        elif len(selected_months) == 2:

            # Ordena os meses automaticamente
            selected_months = sorted(selected_months, key=lambda x: (list(calendar.month_name).index(x) if x in list(calendar.month_name) else float('inf')))

            # Filtra os dados para os meses e o tipo de comparação selecionados
            if selected_comparison_type == "Atrasos":
                df_selected_data = df_dados_mes.loc[df_dados4_mes.index.isin(selected_months)]
                selected_columns = df_dados_mes.columns
            elif selected_comparison_type == "Revisões":
                df_selected_data = df_dados2_mes.loc[df_dados4_mes.index.isin(selected_months)]
                selected_columns = df_dados2_mes.columns
            elif selected_comparison_type == "Envios":
                df_selected_data = df_dados3_mes.loc[df_dados4_mes.index.isin(selected_months)]
                selected_columns = df_dados3_mes.columns
            elif selected_comparison_type == "Erros":
                df_selected_data = df_dados4_mes.loc[df_dados4_mes.index.isin(selected_months)]
                selected_columns = df_dados4_mes.columns

            # Adiciona uma condição para a média anual
            if 'média anual' in selected_months:
                # Filtra os dados para os meses e o tipo de comparação selecionados
                if selected_comparison_type == "Atrasos":
                    df_selected_data = df_dados_mes.loc[df_dados4_mes.index.isin(selected_months)]
                    df_media_anual = df_dados_mes.mean().round(0)
                elif selected_comparison_type == "Revisões":
                    df_selected_data = df_dados2_mes.loc[df_dados4_mes.index.isin(selected_months)]
                    df_media_anual = df_dados2_mes.mean().round(0)
                elif selected_comparison_type == "Envios":
                    df_selected_data = df_dados3_mes.loc[df_dados4_mes.index.isin(selected_months)]
                    df_media_anual = df_dados3_mes.mean().round(0)
                elif selected_comparison_type == "Erros":
                    df_selected_data = df_dados4_mes.loc[df_dados4_mes.index.isin(selected_months)]
                    df_media_anual = df_dados4_mes.mean().round(0)

                # Adiciona a média anual como uma nova linha no DataFrame
                df_selected_data.loc['média anual'] = df_media_anual


            st.markdown(f"<h1 style='font-size: 24px'>Comparação de {selected_comparison_type} entre {selected_months[1]} e {selected_months[0]}</h1>", unsafe_allow_html=True)

            # Soma todas as colunas e verifica se houve aumento ou queda
            soma_mais_recente = df_selected_data.loc[selected_months[1]].sum()
            soma_anterior = df_selected_data.loc[selected_months[0]].sum()
            if 'média anual' in selected_months:
                aumento_ou_queda = "aumento ⬆" if soma_mais_recente < soma_anterior else "queda ⬇"
            else:
                aumento_ou_queda = "aumento ⬆" if soma_mais_recente > soma_anterior else "queda ⬇"
            
            # Calcula a diferença entre os meses selecionados
            diferenca_dados = df_selected_data.diff().iloc[1]
            diferenca_total = int(abs(diferenca_dados.sum()))

            if diferenca_total == 0:
                aumento_ou_queda = 'estabilização'

            # Adiciona formatação de cor ao texto 'queda', 'aumento' e estabilização
            if selected_comparison_type == 'Envios':
                if aumento_ou_queda == 'aumento ⬆':
                    color = 'green'
                elif aumento_ou_queda == 'queda ⬇':
                    color = 'red'
                else:
                    color = 'blue'
            else:
                if aumento_ou_queda == 'aumento ⬆':
                    color = 'red'
                elif aumento_ou_queda == 'queda ⬇':
                    color = 'green'
                else:
                    color = 'blue'
        

            styled_text = f"<span style='color: {color};'>{aumento_ou_queda}</span>"

            # Exibe o texto de comparação total
            st.markdown(f"<strong>Total:</strong> Houve {styled_text} de {diferenca_total} {selected_comparison_type.lower()} entre {selected_months[1]} e {selected_months[0]}.", unsafe_allow_html=True)

            # Exibe o texto de comparação para cada coluna
            for coluna in selected_columns:
                if 'média anual' in selected_months:
                    aumento_ou_queda = "aumento ⬆" if df_selected_data.loc[selected_months[1], coluna] < df_selected_data.loc[selected_months[0], coluna] else "queda ⬇"
                else:
                    aumento_ou_queda = "aumento ⬆" if df_selected_data.loc[selected_months[1], coluna] > df_selected_data.loc[selected_months[0], coluna] else "queda ⬇"

                valor_diferenca = int(abs(diferenca_dados[coluna]))

                if valor_diferenca == 0:
                    aumento_ou_queda = "estabilização"
                    
                # Adiciona formatação de cor ao texto 'queda', 'aumento' e estabilização
                if selected_comparison_type == 'Envios':
                    if aumento_ou_queda == 'aumento ⬆':
                        color = 'green'
                    elif aumento_ou_queda == 'queda ⬇':
                        color = 'red'
                    else:
                        color = 'blue'
                else:
                    if aumento_ou_queda == 'aumento ⬆':
                        color = 'red'
                    elif aumento_ou_queda == 'queda ⬇':
                        color = 'green'
                    else:
                        color = 'blue'

                styled_text = f"<span style='color: {color};'>{aumento_ou_queda}</span>"

                st.markdown(f"<strong>{coluna}</strong>: Houve {styled_text} de {valor_diferenca} {selected_comparison_type.lower()} entre {selected_months[1]} e {selected_months[0]}.", unsafe_allow_html=True)

            # Exibe um gráfico de linha para a comparação
            fig_comparacao = px.line(df_selected_data, labels={'value': f'Qtd de {selected_comparison_type.lower()}'})
            categorias_ordenadas = ['média anual'] + df_selected_data.index.tolist()
            fig_comparacao.update_layout(legend_title_text='Legenda')
            if 'média anual' in selected_months:
                fig_comparacao.update_xaxes(title_text='Mês', categoryorder='array', categoryarray=categorias_ordenadas)
            else:
                fig_comparacao.update_xaxes(title_text='Mês', categoryorder='total ascending')
            st.plotly_chart(fig_comparacao, use_container_width=True)
            
                
        elif len(selected_months) == 1:
            st.warning("Selecione duas colunas para comparação.")
        else:
            st.info("Selecione pelo menos uma coluna para comparação.")
    else:
        st.info("Selecione um tipo de comparação.")
except KeyError as e:
    st.warning(f"Erro ao acessar dados: {e}")


# Menu de filtros
#st.sidebar.header("Filtros")
#selected_column = st.sidebar.selectbox("Escolha uma coluna para filtro", df.columns)
#selected_value = st.sidebar.text_input(f"Filtrar por valor em '{selected_column}'", "")

# Filtros
#if selected_value:
    #df = df[df[selected_column] == selected_value]

# Dados
#st.header("Dados Filtrados")
#st.write(df)