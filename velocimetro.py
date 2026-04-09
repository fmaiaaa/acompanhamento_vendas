import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import plotly.graph_objects as go

def criar_medidor(titulo, realizado, meta, vgv, vendas_qtd):
    """Gera o gráfico de medidor estilo sinaleiro."""
    percentual = (realizado / meta * 100) if (meta and meta > 0) else 0
    
    fig = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = percentual,
        number = {'suffix': "%", 'font': {'size': 24}, 'valueformat': '.1f'},
        title = {'text': titulo, 'font': {'size': 20}},
        gauge = {
            'axis': {'range': [0, 100], 'tickwidth': 1},
            'bar': {'color': "#2c3e50"}, 
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
    
    fig.update_layout(height=280, margin=dict(l=30, r=30, t=50, b=20))
    st.plotly_chart(fig, use_container_width=True)
    
    vgv_formatado = f"R$ {vgv/1e6:.2f} mi" if vgv >= 1e6 else f"R$ {vgv/1e3:.1f} mil"
    
    st.markdown(f"""
    <div style="text-align: center;">
        <small>Vendas:</small> <b>{int(vendas_qtd)}</b> | 
        <small>VGV:</small> <b>{vgv_formatado}</b> | 
        <small>Meta:</small> <b>{int(meta) if meta else 0}</b>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

def main():
    # Configuração da página
    st.set_page_config(layout="wide", page_title="Dashboard Direcional - Metas")
    st.title("📊 Acompanhamento de Metas de Vendas")

    # 1. Conexão com o Google Sheets
    conn = st.connection("gsheets", type=GSheetsConnection)
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1wpuNQvksot9CLhGgQRe7JlyDeRISEh_sc3-6VRDyQYk"

    try:
        # Lendo abas
        df_vendas = conn.read(spreadsheet=spreadsheet_url, worksheet="BD Vendas Completa", ttl="10m")
        df_metas_raw = conn.read(spreadsheet=spreadsheet_url, worksheet="Metas", ttl="10m")

        # 2. Tratamento da Tabela de Metas
        meses_colunas = {
            'jan./1': 1, 'fev./1': 2, 'mar./1': 3, 'abr./1': 4, 
            'mai./1': 5, 'jun./1': 6, 'jul./1': 7, 'ago./1': 8, 
            'set./1': 9, 'out./1': 10, 'nov./1': 11, 'dez./1': 12
        }
        
        df_metas = df_metas_raw.melt(
            id_vars=['Empreendimento', 'Região'], 
            value_vars=[col for col in df_metas_raw.columns if col in meses_colunas.keys()],
            var_name='Mes_Texto', 
            value_name='Meta_Qtd'
        )
        df_metas['Mes_Num'] = df_metas['Mes_Texto'].map(meses_colunas)

        # 3. Sidebar e Filtros
        st.sidebar.header("Filtros")
        df_vendas['Ano da Venda'] = pd.to_numeric(df_vendas['Ano da Venda'], errors='coerce')
        df_vendas['Mês Venda'] = pd.to_numeric(df_vendas['Mês Venda'], errors='coerce')
        
        anos_disponiveis = sorted(df_vendas['Ano da Venda'].dropna().unique().astype(int).tolist())
        ano_sel = st.sidebar.selectbox("Selecione o Ano", anos_disponiveis, index=len(anos_disponiveis)-1)
        
        meses_disponiveis = sorted(df_vendas[df_vendas['Ano da Venda'] == ano_sel]['Mês Venda'].dropna().unique().astype(int).tolist())
        mes_sel = st.sidebar.selectbox("Selecione o Mês", meses_disponiveis)

        # 4. Filtragem
        vendas_filtradas = df_vendas[(df_vendas['Ano da Venda'] == ano_sel) & (df_vendas['Mês Venda'] == mes_sel)]
        metas_filtradas = df_metas[df_metas['Mes_Num'] == mes_sel]

        # 5. Dashboard - Geral
        st.subheader(f"Resultado Geral - Mês {mes_sel} / {ano_sel}")
        total_realizado = vendas_filtradas.shape[0] 
        total_meta = metas_filtradas['Meta_Qtd'].sum()
        total_vgv = vendas_filtradas['Valor Real de Venda'].sum()
        
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            criar_medidor("Geral", total_realizado, total_meta, total_vgv, total_realizado)

        # 6. Dashboard - Regiões
        st.subheader("Indicadores por Região")
        regioes = sorted(df_vendas['Região'].dropna().unique())
        
        cols = st.columns(3)
        for i, regiao in enumerate(regioes):
            with cols[i % 3]:
                v_reg = vendas_filtradas[vendas_filtradas['Região'] == regiao]
                m_reg = metas_filtradas[metas_filtradas['Região'] == regiao]['Meta_Qtd'].sum()
                criar_medidor(regiao, v_reg.shape[0], m_reg, v_reg['Valor Real de Venda'].sum(), v_reg.shape[0])

    except Exception as e:
        st.error(f"Erro: {e}")

if __name__ == "__main__":
    main()
