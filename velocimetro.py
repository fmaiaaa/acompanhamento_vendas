import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# Configuração da página
st.set_page_config(layout="wide", page_title="Dashboard de Metas de Vendas")

st.title("📊 Acompanhamento de Metas de Vendas")

# Sidebar para Upload
st.sidebar.header("Configurações")
uploaded_file = st.sidebar.file_uploader("Suba sua planilha de Vendas/Metas", type=["xlsx"])

if uploaded_file:
    # 1. Carregamento dos Dados
    df_vendas = pd.read_excel(uploaded_file, sheet_name="BD Vendas Completa")
    df_metas_raw = pd.read_excel(uploaded_file, sheet_name="Metas")

    # 2. Tratamento da Tabela de Metas (Transformar meses de colunas para linhas)
    meses_colunas = {
        'jan./1': 1, 'fev./1': 2, 'mar./1': 3, 'abr./1': 4, 
        'mai./1': 5, 'jun./1': 6, 'jul./1': 7, 'ago./1': 8, 
        'set./1': 9, 'out./1': 10, 'nov./1': 11, 'dez./1': 12
    }
    
    # Derreter a tabela de metas
    df_metas = df_metas_raw.melt(
        id_vars=['Empreendimento', 'Região'], 
        value_vars=list(meses_colunas.keys()),
        var_name='Mes_Texto', 
        value_name='Meta_Qtd'
    )
    df_metas['Mes_Num'] = df_metas['Mes_Texto'].map(meses_colunas)

    # 3. Filtros no Sidebar
    anos_disponiveis = sorted(df_vendas['Ano da Venda'].unique().tolist())
    ano_sel = st.sidebar.selectbox("Selecione o Ano", anos_disponiveis, index=len(anos_disponiveis)-1)
    
    meses_disponiveis = sorted(df_vendas[df_vendas['Ano da Venda'] == ano_sel]['Mês Venda'].unique().tolist())
    mes_sel = st.sidebar.selectbox("Selecione o Mês", meses_disponiveis)

    # 4. Filtragem dos Dados
    vendas_filtradas = df_vendas[(df_vendas['Ano da Venda'] == ano_sel) & (df_vendas['Mês Venda'] == mes_sel)]
    metas_filtradas = df_metas[df_metas['Mes_Num'] == mes_sel]

    # Função para criar o Medidor (Gauge Chart)
    def criar_medidor(titulo, realizado, meta, vgv, vendas_qtd):
        percentual = (realizado / meta * 100) if meta > 0 else 0
        
        fig = go.Figure(go.Indicator(
            mode = "gauge+number",
            value = percentual,
            number = {'suffix': "%", 'font': {'size': 20}},
            title = {'text': titulo, 'font': {'size': 18}},
            gauge = {
                'axis': {'range': [0, 100], 'tickwidth': 1},
                'bar': {'color': "#1f77b4"}, # Cor do ponteiro/barra
                'bgcolor': "white",
                'borderwidth': 2,
                'bordercolor': "gray",
                'steps': [
                    {'range': [0, 40], 'color': "#e74c3c"}, # Vermelho
                    {'range': [40, 80], 'color': "#f39c12"}, # Laranja
                    {'range': [80, 100], 'color': "#27ae60"} # Verde
                ],
                'threshold': {
                    'line': {'color': "black", 'width': 4},
                    'thickness': 0.75,
                    'value': 100
                }
            }
        ))
        
        fig.update_layout(height=250, margin=dict(l=20, r=20, t=50, b=20))
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Texto auxiliar ao lado/abaixo como na imagem
        c1, c2, c3 = st.columns(3)
        c1.metric("Vendas", f"{int(vendas_qtd)}")
        c2.metric("VGV", f"R$ {vgv/1e6:.1f} mi" if vgv >= 1e6 else f"R$ {vgv/1e3:.1f} mil")
        c3.metric("Meta", f"{int(meta)}")
        st.markdown("---")

    # 5. Dashboard - Visão Geral
    st.subheader(f"Resultado Geral - {mes_sel}/{ano_sel}")
    
    total_realizado = vendas_filtradas.shape[0] # Quantidade de vendas (IDs)
    total_meta = metas_filtradas['Meta_Qtd'].sum()
    total_vgv = vendas_filtradas['Valor Real de Venda'].sum()
    
    col_geral, col_vazia = st.columns([1, 2])
    with col_geral:
        criar_medidor("Geral", total_realizado, total_meta, total_vgv, total_realizado)

    # 6. Dashboard - Quebra por Região
    st.subheader("Indicadores por Região")
    regioes = sorted(df_vendas['Região'].dropna().unique())
    
    # Criar colunas para as regiões (3 por linha)
    cols = st.columns(3)
    for i, regiao in enumerate(regioes):
        with cols[i % 3]:
            vendas_reg = vendas_filtradas[vendas_filtradas['Região'] == regiao]
            realizado_reg = vendas_reg.shape[0]
            meta_reg = metas_filtradas[metas_filtradas['Região'] == regiao]['Meta_Qtd'].sum()
            vgv_reg = vendas_reg['Valor Real de Venda'].sum()
            
            criar_medidor(regiao, realizado_reg, meta_reg, vgv_reg, realizado_reg)

else:
    st.info("Por favor, suba o arquivo Excel na barra lateral para visualizar os indicadores.")