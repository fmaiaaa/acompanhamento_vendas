# -*- coding: utf-8 -*-
"""
Ferramenta para cálculo da representatividade diária dos indicadores.
Extrai 36 meses de histórico do Salesforce, roda regressões OLS isoladas 
(dia da semana, dia do mês e mês) e converte o volume aditivo esperado
na participação percentual de cada dia dentro do mês alvo.
"""
import os
import io
import time
import calendar
import pandas as pd
import numpy as np
import streamlit as st
from datetime import date, timedelta

# -----------------------------------------------------------------------------
# Constantes Base
# -----------------------------------------------------------------------------
FUNIL_ETAPAS = ("agendamentos", "visitas", "pastas", "pastas_aprovadas", "vendas")
FUNIL_LABELS = {
    "agendamentos": "Agendamentos",
    "visitas": "Visitas",
    "pastas": "Pastas",
    "pastas_aprovadas": "Pastas Aprovadas",
    "vendas": "Vendas",
}

DIAS_SEMANA_PT = {
    0: "segunda", 1: "terça", 2: "quarta", 3: "quinta",
    4: "sexta", 5: "sábado", 6: "domingo"
}
MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}

# -----------------------------------------------------------------------------
# Conexão e Extração Salesforce
# -----------------------------------------------------------------------------
def conectar_salesforce_app():
    """Conecta ao Salesforce usando as secrets nativas do Streamlit."""
    try:
        sec = st.secrets["salesforce"]
        username = sec.get("USER", "")
        password = sec.get("PASSWORD", "")
        token = sec.get("TOKEN", "")
        domain = sec.get("DOMAIN", "login")
        
        from simple_salesforce import Salesforce
        kwargs = {"username": username, "password": password, "domain": domain}
        if token:
            kwargs["security_token"] = token
        return Salesforce(**kwargs)
    except Exception as e:
        st.error(f"Erro ao conectar no Salesforce: {e}")
        return None

def buscar_dados_salesforce_36m(sf):
    """Busca os últimos 36 meses inteiros de dados das 3 entidades comerciais."""
    hoje = date.today()
    inicio_treino = date(hoje.year - 3, hoje.month, 1)
    desde_str = f"{inicio_treino.isoformat()}T00:00:00Z"
    desde_date_str = inicio_treino.isoformat()

    with st.spinner("Extraindo Vendas (Opportunity)..."):
        soql_vendas = (
            "SELECT Id, ContratoGeradoEm__c FROM Opportunity "
            "WHERE DirecionalVendas__c = true "
            "AND Empreendimento__r.Regional__c = 'RJ' "
            "AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
            f"AND ContratoGeradoEm__c >= {desde_date_str}"
        )
        df_ven = pd.DataFrame(sf.query_all(soql_vendas).get("records", []))
        
    with st.spinner("Extraindo Pastas (Avaliacao_credito__c)..."):
        soql_pastas = (
            "SELECT Name, dataPrimeiroEnvioAnalise__c, dataAprovacaoSAFI__c "
            "FROM Avaliacao_credito__c "
            "WHERE Empreendimento__r.Regional__c = 'RJ' "
            "AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
            f"AND CreatedDate >= {desde_str}"
        )
        df_pas = pd.DataFrame(sf.query_all(soql_pastas).get("records", []))

    with st.spinner("Extraindo Agendamentos e Visitas (Event)..."):
        soql_ag = (
            "SELECT Codigo_do_agendamento__c, CreatedDate, Data_da_Visita__c "
            "FROM Event "
            "WHERE Unidade_de_negocio__c = 'Direcional' "
            "AND Regional__c = 'RJ' "
            "AND Empreendimento_de_interesse__c != null "
            f"AND CreatedDate >= {desde_str}"
        )
        df_ag = pd.DataFrame(sf.query_all(soql_ag).get("records", []))

    return df_ag, df_pas, df_ven, inicio_treino

# -----------------------------------------------------------------------------
# Processamento e Calendário
# -----------------------------------------------------------------------------
def formatar_data(serie):
    """Converte strings ISO para date."""
    return pd.to_datetime(serie, errors="coerce").dt.date

def montar_calendario(df_ag, df_pas, df_ven, inicio, fim):
    """Conta os eventos únicos por dia e monta o DataFrame de séries temporais."""
    # Deduplicação Básica (mantendo a data mais recente ou primeira não-nula)
    # Vendas
    if not df_ven.empty:
        df_ven["dt_contrato"] = formatar_data(df_ven.get("ContratoGeradoEm__c"))
        vendas_count = df_ven.dropna(subset=["dt_contrato"]).drop_duplicates("Id")["dt_contrato"].value_counts()
    else:
        vendas_count = pd.Series(dtype=float)

    # Pastas e Aprovadas
    if not df_pas.empty:
        df_pas["dt_envio"] = formatar_data(df_pas.get("dataPrimeiroEnvioAnalise__c"))
        df_pas["dt_safi"] = formatar_data(df_pas.get("dataAprovacaoSAFI__c"))
        pas_count = df_pas.dropna(subset=["dt_envio"]).drop_duplicates("Name")["dt_envio"].value_counts()
        aprov_count = df_pas.dropna(subset=["dt_safi"]).drop_duplicates("Name")["dt_safi"].value_counts()
    else:
        pas_count = pd.Series(dtype=float)
        aprov_count = pd.Series(dtype=float)

    # Agendamentos e Visitas
    if not df_ag.empty:
        df_ag["dt_criacao"] = formatar_data(df_ag.get("CreatedDate"))
        df_ag["dt_visita"] = formatar_data(df_ag.get("Data_da_Visita__c"))
        
        ag_count = df_ag.dropna(subset=["dt_criacao"]).drop_duplicates("Codigo_do_agendamento__c")["dt_criacao"].value_counts()
        vis_count = df_ag.dropna(subset=["dt_visita"]).drop_duplicates("Codigo_do_agendamento__c")["dt_visita"].value_counts()
    else:
        ag_count = pd.Series(dtype=float)
        vis_count = pd.Series(dtype=float)

    # Monta calendário base
    idx = pd.date_range(inicio, fim, freq="D")
    cal = pd.DataFrame({"data": [d.date() for d in idx]})
    
    # Preenche colunas
    cal["agendamentos"] = cal["data"].map(ag_count).fillna(0.0)
    cal["visitas"] = cal["data"].map(vis_count).fillna(0.0)
    cal["pastas"] = cal["data"].map(pas_count).fillna(0.0)
    cal["pastas_aprovadas"] = cal["data"].map(aprov_count).fillna(0.0)
    cal["vendas"] = cal["data"].map(vendas_count).fillna(0.0)
    
    # Variáveis dummy textuais
    cal["dia_mes"] = cal["data"].map(lambda d: d.day)
    cal["dia_semana"] = cal["data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()])
    cal["mes"] = cal["data"].map(lambda d: MESES_PT[d.month])
    
    return cal

# -----------------------------------------------------------------------------
# Regressão OLS de Efeitos Relativos
# -----------------------------------------------------------------------------
def matriz_explicativas_relativa(df: pd.DataFrame) -> np.ndarray:
    """One-hot encoding. Referências: dia 1, segunda-feira, janeiro."""
    n = len(df)
    X = np.zeros((n, 30 + 6 + 11 + 1), dtype=float)
    X[:, -1] = 1.0  # intercepto

    dias_semana_idx = {nome: i for i, nome in DIAS_SEMANA_PT.items()}
    meses_idx = {nome: i for i, nome in MESES_PT.items()}

    for i, row in enumerate(df.itertuples(index=False)):
        dia = int(row.dia_mes)
        if 2 <= dia <= 31:
            X[i, dia - 2] = 1.0
        
        ds = dias_semana_idx.get(str(row.dia_semana), None)
        if ds is not None and ds >= 1:
            X[i, 30 + (ds - 1)] = 1.0
            
        ms = meses_idx.get(str(row.mes), None)
        if ms is not None and ms >= 2:
            X[i, 30 + 6 + (ms - 2)] = 1.0
            
    return X

def estimar_efeitos_sazonais(treino: pd.DataFrame):
    """Roda OLS para isolar impactos de sazonalidade e retorna o dicionário de efeitos."""
    if treino.empty or float(treino["qtd"].sum()) <= 0:
        return None

    X = matriz_explicativas_relativa(treino)
    y = treino["qtd"].astype(float).values
    coef, *_ = np.linalg.lstsq(X, y, rcond=None)
    
    intercepto = float(coef[-1])

    # Construção de Dicionários Lookup
    efeito_dm = {"1": 0.0}
    for d in range(2, 32):
        efeito_dm[str(d)] = float(coef[d - 2])
        
    efeito_ds = {DIAS_SEMANA_PT[0]: 0.0}
    for i in range(1, 7):
        efeito_ds[DIAS_SEMANA_PT[i]] = float(coef[30 + (i - 1)])
        
    efeito_mes = {MESES_PT[1]: 0.0}
    for m in range(2, 13):
        efeito_mes[MESES_PT[m]] = float(coef[30 + 6 + (m - 2)])

    return {
        "intercepto": intercepto,
        "dia_mes": efeito_dm,
        "dia_semana": efeito_ds,
        "mes": efeito_mes,
    }

# -----------------------------------------------------------------------------
# Interface Streamlit
# -----------------------------------------------------------------------------
def to_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Representatividade')
        # Auto-ajuste de colunas
        worksheet = writer.sheets['Representatividade']
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)
    return output.getvalue()

def main():
    st.set_page_config(page_title="Representatividade Sazonal", layout="wide")
    
    st.title("Projeção de Representatividade Diária")
    st.markdown("""
    Esta ferramenta acessa o **Salesforce** em tempo real, capta os últimos **36 meses** 
    fechados, isola o efeito do tempo usando Regressão Linear (OLS) e calcula qual 
    o peso (percentual esperado) de cada dia para o mês alvo selecionado.
    """)
    
    hoje = date.today()
    col1, col2 = st.columns(2)
    with col1:
        ano_alvo = st.number_input("Ano da Projeção", min_value=2020, max_value=2040, value=hoje.year)
    with col2:
        mes_alvo = st.selectbox("Mês da Projeção", options=list(range(1, 13)), format_func=lambda x: MESES_PT[x].capitalize(), index=hoje.month - 1)

    if st.button("Buscar Dados e Calcular Representatividade", type="primary"):
        sf = conectar_salesforce_app()
        if not sf:
            return
            
        # 1. Buscar Dados
        df_ag, df_pas, df_ven, ini_treino = buscar_dados_salesforce_36m(sf)
        
        # O último mês fechado é o mês anterior ao atual
        fim_treino = date(hoje.year, hoje.month, 1) - timedelta(days=1)
        
        with st.spinner("Construindo calendário e processando modelo..."):
            # 2. Criar Calendário Treino
            cal_treino = montar_calendario(df_ag, df_pas, df_ven, ini_treino, fim_treino)
            
            # 3. Gerar Projeção para o Mês Alvo
            dias_no_mes = calendar.monthrange(ano_alvo, mes_alvo)[1]
            datas_alvo = [date(ano_alvo, mes_alvo, d) for d in range(1, dias_no_mes + 1)]
            
            # Inicializa Tabela Resultado
            df_resultado = pd.DataFrame()
            df_resultado["Data"] = datas_alvo
            df_resultado["Dia do Mês"] = df_resultado["Data"].map(lambda d: d.day)
            df_resultado["Dia da Semana"] = df_resultado["Data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()].capitalize())
            
            # Roda regressão para cada etapa
            for etapa in FUNIL_ETAPAS:
                df_temp = cal_treino[["data", "dia_mes", "dia_semana", "mes", etapa]].copy()
                df_temp.rename(columns={etapa: "qtd"}, inplace=True)
                
                efeitos = estimar_efeitos_sazonais(df_temp)
                
                if not efeitos:
                    df_resultado[f"{FUNIL_LABELS[etapa]} (%)"] = 0.0
                    continue
                
                intercepto = efeitos["intercepto"]
                esperados_diarios = []
                
                # Para cada dia no mês alvo, calcula o volume aditivo
                for d in datas_alvo:
                    e_mes = efeitos["mes"].get(MESES_PT[d.month], 0.0)
                    e_dm = efeitos["dia_mes"].get(str(d.day), 0.0)
                    e_ds = efeitos["dia_semana"].get(DIAS_SEMANA_PT[d.weekday()], 0.0)
                    
                    vol_esperado = intercepto + e_mes + e_dm + e_ds
                    esperados_diarios.append(max(vol_esperado, 0.0))
                
                # Normaliza para soma = 100%
                soma_mes = sum(esperados_diarios)
                if soma_mes > 0:
                    pcts = [(v / soma_mes) * 100.0 for v in esperados_diarios]
                else:
                    pcts = [0.0 for _ in esperados_diarios]
                    
                df_resultado[f"{FUNIL_LABELS[etapa]} (%)"] = pcts

        # Finalização da Tabela
        st.success(f"Cálculo concluído! Sazonalidade histórica baseada em dados de {ini_treino.strftime('%m/%Y')} a {fim_treino.strftime('%m/%Y')}.")
        
        # Mostra formatado no Streamlit, mas preserva os floats no Excel
        st.dataframe(
            df_resultado.style.format({c: "{:.2f}%" for c in df_resultado.columns if "(%)" in c}),
            use_container_width=True,
            hide_index=True
        )
        
        # 4. Botão de Download Automático
        excel_bytes = to_excel(df_resultado)
        nome_arquivo = f"Representatividade_{MESES_PT[mes_alvo]}_{ano_alvo}.xlsx"
        
        st.download_button(
            label=f"⬇️ Baixar Relatório em Excel ({nome_arquivo})",
            data=excel_bytes,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

if __name__ == "__main__":
    main()
