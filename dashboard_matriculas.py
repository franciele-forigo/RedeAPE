import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from io import BytesIO

# Configuração da página
st.set_page_config(layout="wide", page_title="Dashboard de Matrículas")
st.title("📊 Dashboard de Matrículas - Região Norte")

# Função principal
def main():
    # Upload do arquivo (substitui o caminho fixo)
    uploaded_file = st.file_uploader("Carregue o arquivo Excel", type=["xlsx"])
    
    if uploaded_file is not None:
        # Lista de anos/abas - ajustável via UI
        anos_selecionados = st.multiselect(
            "Selecione os anos para análise:",
            options=['2023', '2022', '2021', '2020', '2019', '2018', '2017'],
            default=['2023', '2022', '2021']
        )
        
        if not anos_selecionados:
            st.warning("Selecione pelo menos um ano!")
            return
            
        try:
            with st.spinner("Processando dados..."):
                # Dicionário para armazenar dados
                dados_anuais = {}
                
                # Ler cada aba selecionada
                for ano in anos_selecionados:
                    try:
                        df = pd.read_excel(uploaded_file, sheet_name=ano)
                        cols_to_keep = ['Instituição', 'Unidade', 'Tipo de Curso', 'Nome do Curso', 
                                       'Tipo de Oferta', 'Modalidade de Ensino', 'Matrículas']
                        df = df[cols_to_keep].copy()
                        df.rename(columns={'Matrículas': f"Matrículas_{ano}"}, inplace=True)
                        dados_anuais[ano] = df
                    except Exception as e:
                        st.warning(f"A aba {ano} não foi encontrada: {str(e)}")
                        continue
                
                if not dados_anuais:
                    st.error("Nenhuma aba válida encontrada!")
                    return
                
                # Juntar DataFrames
                chaves = ['Instituição', 'Unidade', 'Tipo de Curso', 'Nome do Curso', 
                         'Tipo de Oferta', 'Modalidade de Ensino']
                
                df_ranking = None
                for ano, df in dados_anuais.items():
                    coluna_matriculas = f"Matrículas_{ano}"
                    if df_ranking is None:
                        df_ranking = df[chaves + [coluna_matriculas]].copy()
                    else:
                        temp_df = df[chaves + [coluna_matriculas]]
                        df_ranking = pd.merge(df_ranking, temp_df, on=chaves, how='outer')
                
                # Processamento final
                colunas_matriculas = [c for c in df_ranking.columns if c.startswith('Matrículas_')]
                for col in colunas_matriculas:
                    df_ranking[col] = pd.to_numeric(df_ranking[col], errors='coerce').fillna(0)
                
                df_ranking['Total_Matrículas'] = df_ranking[colunas_matriculas].sum(axis=1)
                df_ranking = df_ranking.sort_values(by='Total_Matrículas', ascending=False)
                df_ranking.insert(0, 'Ranking', range(1, len(df_ranking) + 1))
                
                # Criar identificação resumida
                df_ranking['Identificação'] = (
                    df_ranking['Instituição'].str[:15] + ' - ' + 
                    df_ranking['Nome do Curso'].str[:20]
                )
                
                # ========== VISUALIZAÇÕES ==========
                st.success("Dados carregados com sucesso!")
                
                # 1. Filtros interativos
                st.sidebar.header("Filtros")
                min_matriculas = st.sidebar.slider(
                    "Mínimo de matrículas totais:",
                    min_value=0,
                    max_value=int(df_ranking['Total_Matrículas'].max()),
                    value=100
                )
                
                tipo_curso = st.sidebar.multiselect(
                    "Tipo de curso:",
                    options=df_ranking['Tipo de Curso'].unique(),
                    default=df_ranking['Tipo de Curso'].unique()
                )
                
                # Aplicar filtros
                df_filtrado = df_ranking[
                    (df_ranking['Total_Matrículas'] >= min_matriculas) &
                    (df_ranking['Tipo de Curso'].isin(tipo_curso))
                ]
                
                # 2. Métricas resumidas
                col1, col2, col3 = st.columns(3)
                col1.metric("Total de Cursos", len(df_filtrado))
                col2.metric("Top 1 Curso", df_filtrado.iloc[0]['Identificação'])
                col3.metric("Matrículas do Top 1", f"{df_filtrado.iloc[0]['Total_Matrículas']:,.0f}")
                
                # 3. Tabela interativa
                st.subheader("📋 Dados Completos")
                st.dataframe(df_filtrado, height=400, use_container_width=True)
                
                # 4. Visualizações em abas
                tab1, tab2, tab3 = st.tabs(["Top 20", "Evolução Anual", "Distribuição"])
                
                with tab1:
                    st.subheader("🏆 Top 20 Cursos")
                    top20 = df_filtrado.head(20)
                    
                    fig, ax = plt.subplots(figsize=(10, 8))
                    bars = ax.barh(
                        top20['Identificação'],
                        top20['Total_Matrículas'],
                        color='#1f77b4'
                    )
                    
                    ax.bar_label(bars, 
                               labels=[f"{x:,.0f}".replace(",", ".") for x in top20['Total_Matrículas']],
                               padding=5,
                               fontsize=10)
                    
                    ax.invert_yaxis()
                    ax.set_xlabel("Total de Matrículas")
                    plt.tight_layout()
                    st.pyplot(fig)
                    
                    # Botão de download
                    buffer = BytesIO()
                    plt.savefig(buffer, format='png')
                    st.download_button(
                        label="Baixar gráfico",
                        data=buffer,
                        file_name="top20_matriculas.png",
                        mime="image/png"
                    )
                
                with tab2:
                    st.subheader("📈 Evolução Anual")
                    num_cursos = st.slider("Número de cursos para mostrar:", 1, 20, 5)
                    top_evolucao = df_filtrado.head(num_cursos)
                    
                    fig, ax = plt.subplots(figsize=(12, 6))
                    anos = sorted([int(ano.split('_')[1]) for ano in colunas_matriculas])
                    
                    for idx, row in top_evolucao.iterrows():
                        valores = [row[f'Matrículas_{ano}'] for ano in anos]
                        ax.plot(anos, valores, marker='o', label=row['Identificação'])
                    
                    ax.set_title(f"Evolução das Matrículas - Top {num_cursos}")
                    ax.set_xlabel("Ano")
                    ax.set_ylabel("Matrículas")
                    ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
                    ax.grid(True)
                    st.pyplot(fig)
                
                with tab3:
                    st.subheader("🍕 Distribuição")
                    st.write("Participação por instituição (Top 10)")
                    
                    top_instituicoes = df_filtrado.groupby('Instituição')['Total_Matrículas']\
                                                .sum()\
                                                .sort_values(ascending=False)\
                                                .head(10)
                    
                    fig, ax = plt.subplots(figsize=(10, 10))
                    ax.pie(
                        top_instituicoes,
                        labels=top_instituicoes.index,
                        autopct='%1.1f%%',
                        startangle=90,
                        textprops={'fontsize': 10}
                    )
                    ax.set_title("Distribuição por Instituição")
                    st.pyplot(fig)
        
        except Exception as e:
            st.error(f"Erro durante o processamento: {str(e)}")

if __name__ == "__main__":
    main()