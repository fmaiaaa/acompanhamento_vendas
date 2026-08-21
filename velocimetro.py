# -*- coding: utf-8 -*-
"""
Acompanhamento de vendas — metas vs realizado (Direcional).
Planilha: BD Vendas Completa + Metas.
Design: Gaps Style (Transparência, Blur, Inter/Montserrat).
Funcionalidade: Engenharia Reversa, Comparativo MTD e Pesos de Coordenadores.

Arquivo único e autossuficiente — painel v2, dashboard comercial, poder de compra
e tempos de funil estão inline (sem imports de outros .py locais).
"""
import base64
import calendar
import copy
import html
import json
import math
import os
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

import velocimetro_cache as vc


def _st_dialog_decorator(title: str = "", **kwargs):
    """Compatível com Streamlit antigo e imports offline (mock sem st.dialog)."""
    dialog_fn = getattr(st, "dialog", None)
    if callable(dialog_fn):
        return dialog_fn(title, **kwargs)

    def _identity(fn):
        return fn

    return _identity

_mod_dir = Path(__file__).resolve().parent
if str(_mod_dir) not in sys.path:
    sys.path.insert(0, str(_mod_dir))

try:
    from velocimetro_feedbacks_previsao import (
        carregar_feedbacks_comerciais,
        carregar_previsao_vendas,
        render_aba_feedbacks_comerciais,
        render_aba_previsao_vendas,
    )
except ImportError:
    try:
        import importlib.util
        _fb_path = _mod_dir / "velocimetro_feedbacks_previsao.py"
        if _fb_path.is_file():
            _spec = importlib.util.spec_from_file_location(
                "velocimetro_feedbacks_previsao", _fb_path,
            )
            _fb_mod = importlib.util.module_from_spec(_spec)
            assert _spec.loader is not None
            _spec.loader.exec_module(_fb_mod)
            carregar_feedbacks_comerciais = _fb_mod.carregar_feedbacks_comerciais
            carregar_previsao_vendas = _fb_mod.carregar_previsao_vendas
            render_aba_feedbacks_comerciais = _fb_mod.render_aba_feedbacks_comerciais
            render_aba_previsao_vendas = _fb_mod.render_aba_previsao_vendas
        else:
            raise ImportError("velocimetro_feedbacks_previsao.py não encontrado")
    except Exception:
        carregar_feedbacks_comerciais = None  # type: ignore
        carregar_previsao_vendas = None  # type: ignore
        render_aba_feedbacks_comerciais = None  # type: ignore
        render_aba_previsao_vendas = None  # type: ignore

# -----------------------------------------------------------------------------
# Identificação da planilha e Arquivos Visuais
# -----------------------------------------------------------------------------
# Base consolidada (atualizada via GitHub Actions / botão «Atualizar dados»)
SPREADSHEET_CONSOLIDADA_ID = "1wpuNQvksot9CLhGgQRe7JlyDeRISEh_sc3-6VRDyQYk"
SPREADSHEET_ID = SPREADSHEET_CONSOLIDADA_ID

WS_VENDAS = "BD Vendas Completa"
WS_METAS = "Metas"
WS_ESTOQUE = "Base Estoque"

# Metas v2 (coordenadores + canal) — mesma planilha consolidada
SPREADSHEET_METAS_COORD_ID = SPREADSHEET_CONSOLIDADA_ID
WS_METAS_COORD = "Metas Coordenadores Comerciais"
SPREADSHEET_BASES_IVAN_ID = SPREADSHEET_CONSOLIDADA_ID
WS_CANAL = "Canal"
WS_CANAL_ALIASES = ("Canal", "Cópia de Canal", "Copia de Canal", "Cópia de canal")

# Funil comercial — mesma planilha consolidada
SPREADSHEET_FUNIL_ID = SPREADSHEET_CONSOLIDADA_ID
# Pastas / pastas aprovadas (aba BASE) — planilha consolidada
SPREADSHEET_PASTAS_ID: Optional[str] = SPREADSHEET_CONSOLIDADA_ID
# Agendamentos/visitas, pastas e vendas: relatórios Salesforce (evita limite do Sheets)
SF_REPORT_AGENDAMENTOS_ID = "00OU600000AcFGPMA3"
SF_REPORT_PASTAS_ID = "00OU600000FEOoDMAX"
SF_REPORT_VENDAS_ID = "00O3Z000005ZsPmUAK"
SF_REPORT_ESTOQUE_ID = "00OU600000FbbMXMAZ"
# Status de unidade (Produto__c) — alinhado ao relatório de estoque SF
ESTOQUE_STATUS_VENDAVEL = frozenset({"Disponível", "Mirror"})
ESTOQUE_STATUS_TODOS = (
    "Disponível",
    "Mirror",
    "Fora de venda",
    "Fora de Venda - Comercial",
)
ALIASES_STATUS_UNIDADE = [
    "StatusUnidade__c", "Status da Unidade", "Status_da_Unidade__c", "Status",
]
ABA_AGENDAMENTOS_VISITAS = "Dados Únicos"  # fallback Sheets se SF falhar
ABA_PASTAS_CANDIDATAS = (
    "BASE",
    "Base",
    "Pastas",
    "Pastas e Pastas Aprovadas",
    "Pastas Aprovadas",
    "BD Pastas",
)
FUNIL_ETAPAS = ("agendamentos", "visitas", "pastas", "pastas_aprovadas", "vendas")
PESOS_FUNIL_MTD: Dict[str, int] = {
    "agendamentos": 1,
    "visitas": 2,
    "pastas": 3,
    "pastas_aprovadas": 4,
    "vendas": 5,
}
FUNIL_DRIVERS = ("agendamentos", "visitas", "pastas", "pastas_aprovadas")
FUNIL_LABELS = {
    "agendamentos": "Agendamentos",
    "visitas": "Visitas",
    "pastas": "Pastas",
    "pastas_aprovadas": "Pastas aprovadas",
    "vendas": "Vendas",
}
FUNIL_DIGITAL_ETAPAS = (
    "leads",
    "agendamentos",
    "visitas",
    "pastas",
    "pastas_aprovadas",
    "vendas",
)
FUNIL_DIGITAL_LABELS = {
    "leads": "Leads",
    **FUNIL_LABELS,
}
FUNIL_DIGITAL_CORES = ["#011a3d", "#022654", "#04428f", "#1e60b3", "#cb0935", "#9e0828"]
# Conversões etapa → etapa seguinte
FUNIL_PARES_ETAPA = (
    ("agendamentos", "visitas"),
    ("visitas", "pastas"),
    ("pastas", "pastas_aprovadas"),
    ("pastas_aprovadas", "vendas"),
)
# Quatro blocos semanais não sobrepostos, recalculados em cada dia.
FUNIL_LAG_BLOCOS = ((1, 7), (8, 14), (15, 21), (22, 28))
FUNIL_LAGS = tuple(range(1, 29))  # compatibilidade/horizonte máximo
FUNIL_JANELAS_CONV = (7, 14, 21, 28)
# 7 conversões únicas: 4 etapa→etapa + 3 diretas adicionais→venda.
FUNIL_CONVERSOES = FUNIL_PARES_ETAPA + (
    ("agendamentos", "vendas"),
    ("visitas", "vendas"),
    ("pastas", "vendas"),
)
FUNIL_LAGS_PERFIL = tuple(range(1, 29))  # compatibilidade: horizonte dos blocos
FUNIL_JANELA_CONV = max(FUNIL_JANELAS_CONV)  # compatibilidade
FUNIL_JANELA_FORCA = 7  # janela da força de trabalho (atividade cruzada)
FUNIL_ITERS_CRUZADAS = 3  # iterações na projeção (efeitos contemporâneos cruzados)
FUNIL_RIDGE_ALPHA = 5.0
FUNIL_ANOS_TREINO = 3
# Ano mais recente, ano intermediário e ano mais antigo.
FUNIL_PESOS_ANUAIS = (1.00, 0.55, 0.30)
# Buffer de produção: 28 dias de lags antes do 1º dia do mês atual.
FUNIL_SOQL_BUFFER_LAGS = 28
# Painel web: vendas com janela curta (projeção OLS usa até 24m; evita SOQL de 36m).
PAINEL_MESES_VENDAS = 24
FUNIL_MODELO_SCHEMA = "funil_elasticnet_cal_lags_v1"
FUNIL_ELASTICNET_ALPHA = 0.5
FUNIL_ELASTICNET_L1_RATIO = 0.3

# BEGIN FUNIL_MODELO_PRODUCAO
# Bloco substituído atomicamente por treinar_modelo_funil.py (treino mensal).
FUNIL_MODELO_PRODUCAO: Optional[Dict[str, Any]] = {
    'schema_version': 'funil_elasticnet_cal_lags_v1',
    'conjunto': 'cal_lags',
    'alpha': 0.5,
    'l1_ratio': 0.3,
    'incluir_mes': True,
    'feature_names': [
        'cal_0',
        'cal_1',
        'cal_2',
        'cal_3',
        'cal_4',
        'cal_5',
        'cal_6',
        'cal_7',
        'cal_8',
        'cal_9',
        'cal_10',
        'cal_11',
        'cal_12',
        'cal_13',
        'cal_14',
        'cal_15',
        'cal_16',
        'cal_17',
        'cal_18',
        'cal_19',
        'cal_20',
        'cal_21',
        'cal_22',
        'cal_23',
        'cal_24',
        'cal_25',
        'cal_26',
        'cal_27',
        'cal_28',
        'cal_29',
        'cal_30',
        'cal_31',
        'cal_32',
        'cal_33',
        'cal_34',
        'cal_35',
        'cal_36',
        'cal_37',
        'cal_38',
        'cal_39',
        'cal_40',
        'cal_41',
        'cal_42',
        'cal_43',
        'cal_44',
        'cal_45',
        'cal_46',
        'cal_47',
        'cal_48',
        'cal_49',
        'agendamentos_lag1_7',
        'agendamentos_lag8_14',
        'agendamentos_lag15_21',
        'agendamentos_lag22_28',
        'visitas_lag1_7',
        'visitas_lag8_14',
        'visitas_lag15_21',
        'visitas_lag22_28',
        'pastas_lag1_7',
        'pastas_lag8_14',
        'pastas_lag15_21',
        'pastas_lag22_28',
        'pastas_aprovadas_lag1_7',
        'pastas_aprovadas_lag8_14',
        'pastas_aprovadas_lag15_21',
        'pastas_aprovadas_lag22_28',
        'vendas_lag1_7',
        'vendas_lag8_14',
        'vendas_lag15_21',
        'vendas_lag22_28',
    ],
    'coefs': {
        'agendamentos': [
            -11.050023757122998, 1.038487609988671, -0.418124043450884, 0.442395732483458, 0.0, 1.5043060233127707, -4.79664719604406, -5.604638520549051,
            0.3307953773367116, -0.6395487540339839, 0.0, 1.0386498125055847, 1.1726599921675265, -0.0, -4.698518480368115, -1.0308131825149127,
            3.6445785363732806, -1.447882060382428, -7.204983402900458, 0.0, 3.651255151779177, -3.5088313120886867, -0.0, 8.244037224651322,
            0.0, 0.0, 3.700517849269758, 5.993521345091159, -0.34789308136715963, 7.8649989063544545, -2.91418811651284, -11.801019009446438,
            -1.4794439627409222, -10.98289209693722, 50.66223419883655, 4.591966028051389, 0.27513531844632166, -29.79641318877947, -2.032225387855265, -0.6330064208964424,
            8.419611128654562, 6.163075652020995, -0.7073162326061971, -0.0, 3.133611441766049, 0.7368229250218166, 0.6725987573442797, 1.3207690091877553,
            -3.3988118810412016, -12.80896936444073, 0.028546673248611723, 0.017664390956704096, 0.0074582695645180296, 0.011935259954292352, 0.07354415744056826, 0.0392253539742714,
            0.022737692506707458, 0.0127343370903805, 0.02552922921923293, 0.0, -0.002926447724372564, -0.0, -0.023879040237688397, 0.0,
            0.0, -0.009814685318714121, -0.13605947923178324, -0.11051715958782876, -0.0480291768980494, -0.1111889736858423, 25.518674435130997,
        ],
        'visitas': [
            -0.1446158706482313, 0.0, -0.6576659613751795, 0.1389540086584349, -0.377873533205042, -0.146361483177348, -0.9987717886271178, -1.0849597046249826,
            0.0, -0.08800284212767584, 0.0, -0.0, -0.0, -0.0, -0.6446532418450247, -0.0,
            -0.0, 1.2697311523488617, -0.0, 1.058348499769358, 0.0, -0.0, 1.373782691160814, 0.2591807150881517,
            -0.0, 0.0, 0.8108179657369299, 1.3478352049918847, -0.0, 0.5569768904217159, -0.0, -1.1072165391982218,
            -2.308244319772371, -3.960928349127075, -3.113707031167763, -1.0535131169204035, 18.6605747509046, -1.0550990313684698, 1.596962357264277, -1.298001927847425,
            0.47226600860013807, -0.0, 0.0, -2.0851054556848343, -0.584850135478679, -0.0, 0.5413827948190845, 2.0906848416485118,
            1.1774488379230574, -3.3758146603146897, 0.0074536506421286125, 0.003978534663514848, 0.001880905278990792, 0.00236741202701496, 0.030892313051985438, 0.02112699072691853,
            0.009491838610875316, 0.007083399341173249, 0.01385231404697147, -0.0024286416883260793, -0.0, 0.0009979376020420963, 0.0008007206085815955, -0.0,
            0.0, 0.0, -0.03218068527766599, -0.017397647848058598, -0.021067391256373167, -0.029106308549339335, 6.760799775315082,
        ],
        'pastas': [
            -2.257660160372139, 0.37804668806549585, -0.8188908909767566, -0.7911168539879757, -0.0, -0.6276025072557446, -0.5693125888781922, -0.0,
            0.36730776869997234, -0.6090992493485636, 0.0, -0.0, 0.0, -0.6407029879254541, -0.22363716339833273, -0.9959623769656082,
            -0.0995014340792171, 0.0, -0.4446914006054063, 0.053702775562173344, 0.0, -0.0, 0.25681307624841726, 0.4879161608206745,
            0.25974634112636374, 2.829001296348825, 3.3903746625625044, 1.6216671625960872, 0.0, 1.5054829998167254, 0.0, 0.0,
            2.169910744495752, 0.8866311180081654, 2.5190340402856277, 2.568863992477152, 1.8252468917425655, -15.952118304331036, 4.089806676160004, -0.0,
            3.0244750667088462, 0.0, 0.0, -0.21947233759338208, -1.4048245833108872, 0.0, -1.145801155726143, -0.9141429197182945,
            0.5667724453793453, -3.735056172948177, 0.0011883536587739955, -0.0, -0.0, -0.0, 0.0038239110262645822, 0.0,
            -0.0, -0.0, 0.053076383684951864, 0.01628370008719105, 0.010181029816810011, 0.01134665219209089, 0.01919491590431224, 0.0,
            0.0034907449374374897, -0.003983682455728812, -0.0, -0.0, 0.0, 0.0, 8.541321401889125,
        ],
        'pastas_aprovadas': [
            -0.0, 1.286418801392201, -0.0, 0.1096047048215759, -0.0, -0.0, -0.0, -0.0,
            -0.0, -0.6873312685412176, -0.0, 0.0, -0.0, -0.45349523807398806, 0.0, -0.0,
            0.0, -0.22197244123571663, -0.6668303696314954, 0.0, 0.0, -0.023492855597622053, 0.41712895519882753, -0.173797479705193,
            -0.0, 0.9175906237994016, 0.42912541252567155, 0.5037806961811774, 0.0, 0.04251140271987944, 0.0, 1.2070697707842875,
            -0.0, 0.004452492087191327, 0.0, 0.9641332371870946, 0.28237430750078957, -6.032825556173662, 0.6036149400664597, -0.2409258248307588,
            0.0, 0.4605254209469333, 0.49979582870123335, 0.0, -0.0, 0.0, 0.0, -0.18869037029264502,
            -0.2726160873448889, -0.6656027990114094, 0.00013888108779133894, 0.0, 0.0, 0.0, 0.0012466375312928557, 0.00042487296661787815,
            -0.0, -0.0, 0.0211420449696927, 0.00404804179782973, 0.0, 0.00029221606057710744, 0.019431117704225863, 0.012134450327900354,
            0.0, 0.0, -0.0, 0.0, 0.0, 0.005232421713483153, 2.8626292290311195,
        ],
        'vendas': [
            0.20528593276015503, 0.698417783067199, 0.3675451925045679, 0.321792704983368, -0.0, -0.0, -0.0, 0.0,
            -0.0, -0.8532504518005523, -0.053521301656589367, -0.0, 0.0, -0.20691941235240396, -0.0, -0.0,
            -0.0, -0.0, -0.13175897883196047, 0.0, -0.0, -0.0, -0.0, 0.0,
            0.18071359527530556, 0.0, 0.0, 0.1894548720542877, 0.0, 0.389496984057287, -0.0, 0.35908831598201363,
            -0.0, -0.061743672343900026, -0.09141691518218586, 0.0, 2.0783056039277445, -1.8354991984619378, 0.0, -0.12725787236565514,
            -0.0, 0.467109201218153, 0.33104757574626836, 0.0, -0.0, -0.003614909853667342, -0.0, -0.0,
            0.0009870482915211455, -0.0, -0.0, -0.00021316560066822872, -0.0005062140142649687, -0.00027505732719980804, -0.0, -0.0,
            -0.0015675284840701315, -0.0009382118647754134, 0.012169596966085192, 0.0, -0.0, 0.0, 0.008202816037259452, -0.0,
            -0.0, 0.0, 0.014770382036083303, 0.011190452780320471, 0.003443453390416612, 0.01630494307484225, 1.5042820591847148,
        ],
    },
    'r2s': {
        'agendamentos': 0.5674702970929021,
        'visitas': 0.6609137060358623,
        'pastas': 0.5620657058767049,
        'pastas_aprovadas': 0.38674799652549374,
        'vendas': 0.2636979650360082,
    },
    'r2s_medias': {
        'agendamentos': 0.38820001794154535,
        'visitas': 0.4442918693794071,
        'pastas': 0.24223162547035249,
        'pastas_aprovadas': 0.205858153864531,
        'vendas': 0.01357569731627517,
    },
    'holdout': {
        'agendamentos': {
            'r2': 0.4356756297641634,
            'mae': 23.213446897234107,
        },
        'visitas': {
            'r2': 0.5446585691158319,
            'mae': 7.062861678422211,
        },
        'pastas': {
            'r2': 0.44995359427173354,
            'mae': 9.128324211916174,
        },
        'pastas_aprovadas': {
            'r2': 0.2420440524849954,
            'mae': 5.911561624976788,
        },
        'vendas': {
            'r2': 0.19815761714186142,
            'mae': 3.1835725181959447,
        },
    },
    'holdout_dias': 90,
    'medias': {
        'incluir_mes': True,
        'lags': [
            1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0,
            9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0,
            17.0, 18.0, 19.0, 20.0, 21.0, 22.0, 23.0, 24.0,
            25.0, 26.0, 27.0, 28.0,
        ],
        'forca_mu': 0.16360326522795815,
        'etapas': {
            'agendamentos': {
                'mu': 76.60923770184003,
                'media_dia_semana': {
                    'domingo': 36.93963254593176,
                    'quarta': 61.98416050686378,
                    'quinta': 142.90126715945092,
                    'segunda': 60.36886102403345,
                    'sexta': 82.67951425554384,
                    'sábado': 77.37736842105264,
                    'terça': 74.43625914315571,
                },
                'media_dia_mes': {
                    1: 60.13425925925926,
                    2: 81.64814814814814,
                    3: 73.9560185185185,
                    4: 82.05324074074073,
                    5: 80.1027397260274,
                    6: 78.9634703196347,
                    7: 69.06392694063926,
                    8: 64.31050228310502,
                    9: 77.71232876712328,
                    10: 71.92922374429224,
                    11: 79.10730593607305,
                    12: 80.93378995433791,
                    13: 77.85159817351597,
                    14: 76.04566210045662,
                    15: 66.17808219178082,
                    16: 74.93835616438356,
                    17: 79.58904109589041,
                    18: 76.39497716894977,
                    19: 69.05936073059361,
                    20: 76.05479452054793,
                    21: 82.04794520547945,
                    22: 67.80821917808218,
                    23: 76.84018264840182,
                    24: 86.40182648401827,
                    25: 80.07762557077625,
                    26: 79.8173515981735,
                    27: 81.94748858447488,
                    28: 85.9337899543379,
                    29: 74.007371007371,
                    30: 90.8698795180723,
                    31: 71.3201581027668,
                },
                'media_mes': {
                    'abril': 95.0036036036036,
                    'agosto': 70.18432769367766,
                    'dezembro': 58.16477768090672,
                    'fevereiro': 67.42706333973129,
                    'janeiro': 64.84829991281605,
                    'julho': 70.1207075962539,
                    'junho': 63.23398576512456,
                    'maio': 76.75239755884918,
                    'março': 87.11072362685267,
                    'novembro': 93.4063063063063,
                    'outubro': 94.77506538796862,
                    'setembro': 77.14054054054054,
                },
                'lag_mu': {
                    'agendamentos_lag1_7': 536.7742395794216,
                    'agendamentos_lag8_14': 536.8396545249718,
                    'agendamentos_lag15_21': 536.5288021028914,
                    'agendamentos_lag22_28': 535.1260983852798,
                    'visitas_lag1_7': 198.6619601952685,
                    'visitas_lag8_14': 198.523695080736,
                    'visitas_lag15_21': 198.2122418325197,
                    'visitas_lag22_28': 197.6323695080736,
                    'pastas_lag1_7': 204.10874953060457,
                    'pastas_lag8_14': 203.90769808486667,
                    'pastas_lag15_21': 203.74787833270747,
                    'pastas_lag22_28': 203.64115659031165,
                    'pastas_aprovadas_lag1_7': 74.10544498685692,
                    'pastas_aprovadas_lag8_14': 74.19038678182501,
                    'pastas_aprovadas_lag15_21': 74.08396545249718,
                    'pastas_aprovadas_lag22_28': 74.17506571535863,
                    'vendas_lag1_7': 38.555013143071726,
                    'vendas_lag8_14': 38.6723995493804,
                    'vendas_lag15_21': 38.739992489673305,
                    'vendas_lag22_28': 38.83274502440857,
                },
            },
            'visitas': {
                'mu': 28.36395043184378,
                'media_dia_semana': {
                    'domingo': 25.637270341207348,
                    'quarta': 22.024815205913413,
                    'quinta': 23.198521647307288,
                    'segunda': 25.600313479623825,
                    'sexta': 25.83474128827878,
                    'sábado': 52.20263157894737,
                    'terça': 24.064263322884013,
                },
                'media_dia_mes': {
                    1: 27.65740740740741,
                    2: 29.92824074074074,
                    3: 26.627314814814817,
                    4: 29.77777777777778,
                    5: 26.47945205479452,
                    6: 28.38812785388128,
                    7: 26.897260273972602,
                    8: 25.24657534246575,
                    9: 28.76712328767123,
                    10: 26.319634703196343,
                    11: 28.545662100456617,
                    12: 27.35616438356164,
                    13: 28.915525114155248,
                    14: 28.22146118721461,
                    15: 25.85616438356164,
                    16: 27.59817351598173,
                    17: 26.787671232876708,
                    18: 30.287671232876708,
                    19: 26.94063926940639,
                    20: 31.068493150684933,
                    21: 28.95890410958904,
                    22: 26.933789954337897,
                    23: 31.00228310502283,
                    24: 28.557077625570773,
                    25: 28.057077625570777,
                    26: 28.14840182648402,
                    27: 31.2283105022831,
                    28: 31.906392694063925,
                    29: 27.660933660933665,
                    30: 30.94216867469879,
                    31: 28.276679841897234,
                },
                'media_mes': {
                    'abril': 31.824324324324323,
                    'agosto': 25.198575244879788,
                    'dezembro': 23.156931124673065,
                    'fevereiro': 23.823416506717855,
                    'janeiro': 27.864864864864867,
                    'julho': 22.68782518210198,
                    'junho': 21.243772241992886,
                    'maio': 30.32606800348736,
                    'março': 29.564080209241503,
                    'novembro': 38.168468468468475,
                    'outubro': 36.45335658238884,
                    'setembro': 28.95765765765766,
                },
                'lag_mu': {
                    'agendamentos_lag1_7': 536.7742395794216,
                    'agendamentos_lag8_14': 536.8396545249718,
                    'agendamentos_lag15_21': 536.5288021028914,
                    'agendamentos_lag22_28': 535.1260983852798,
                    'visitas_lag1_7': 198.6619601952685,
                    'visitas_lag8_14': 198.523695080736,
                    'visitas_lag15_21': 198.2122418325197,
                    'visitas_lag22_28': 197.6323695080736,
                    'pastas_lag1_7': 204.10874953060457,
                    'pastas_lag8_14': 203.90769808486667,
                    'pastas_lag15_21': 203.74787833270747,
                    'pastas_lag22_28': 203.64115659031165,
                    'pastas_aprovadas_lag1_7': 74.10544498685692,
                    'pastas_aprovadas_lag8_14': 74.19038678182501,
                    'pastas_aprovadas_lag15_21': 74.08396545249718,
                    'pastas_aprovadas_lag22_28': 74.17506571535863,
                    'vendas_lag1_7': 38.555013143071726,
                    'vendas_lag8_14': 38.6723995493804,
                    'vendas_lag15_21': 38.739992489673305,
                    'vendas_lag22_28': 38.83274502440857,
                },
            },
            'pastas': {
                'mu': 29.147277506571534,
                'media_dia_semana': {
                    'domingo': 8.850918635170602,
                    'quarta': 31.4493136219641,
                    'quinta': 33.713833157338975,
                    'segunda': 30.278474399164054,
                    'sexta': 33.79936642027455,
                    'sábado': 32.84789473684211,
                    'terça': 33.14315569487984,
                },
                'media_dia_mes': {
                    1: 25.488425925925924,
                    2: 31.828703703703702,
                    3: 27.81481481481481,
                    4: 28.458333333333332,
                    5: 28.917808219178077,
                    6: 28.271689497716892,
                    7: 26.751141552511413,
                    8: 27.691780821917806,
                    9: 30.828767123287665,
                    10: 27.020547945205475,
                    11: 29.541095890410958,
                    12: 28.897260273972595,
                    13: 29.499999999999996,
                    14: 26.37442922374429,
                    15: 26.874429223744286,
                    16: 27.15296803652968,
                    17: 27.360730593607304,
                    18: 29.707762557077622,
                    19: 27.34246575342465,
                    20: 29.93607305936073,
                    21: 28.098173515981728,
                    22: 27.141552511415526,
                    23: 30.280821917808215,
                    24: 29.91552511415525,
                    25: 30.458904109589035,
                    26: 33.51598173515981,
                    27: 35.05936073059361,
                    28: 31.634703196347026,
                    29: 29.48648648648648,
                    30: 33.26024096385542,
                    31: 29.15810276679842,
                },
                'media_mes': {
                    'abril': 35.3981981981982,
                    'agosto': 26.626001780943906,
                    'dezembro': 21.553618134263296,
                    'fevereiro': 31.99808061420346,
                    'janeiro': 33.53443766346992,
                    'julho': 25.298647242455775,
                    'junho': 28.266014234875446,
                    'maio': 31.7794245858762,
                    'março': 36.14908456843941,
                    'novembro': 29.044144144144145,
                    'outubro': 24.69311246730602,
                    'setembro': 25.059459459459458,
                },
                'lag_mu': {
                    'agendamentos_lag1_7': 536.7742395794216,
                    'agendamentos_lag8_14': 536.8396545249718,
                    'agendamentos_lag15_21': 536.5288021028914,
                    'agendamentos_lag22_28': 535.1260983852798,
                    'visitas_lag1_7': 198.6619601952685,
                    'visitas_lag8_14': 198.523695080736,
                    'visitas_lag15_21': 198.2122418325197,
                    'visitas_lag22_28': 197.6323695080736,
                    'pastas_lag1_7': 204.10874953060457,
                    'pastas_lag8_14': 203.90769808486667,
                    'pastas_lag15_21': 203.74787833270747,
                    'pastas_lag22_28': 203.64115659031165,
                    'pastas_aprovadas_lag1_7': 74.10544498685692,
                    'pastas_aprovadas_lag8_14': 74.19038678182501,
                    'pastas_aprovadas_lag15_21': 74.08396545249718,
                    'pastas_aprovadas_lag22_28': 74.17506571535863,
                    'vendas_lag1_7': 38.555013143071726,
                    'vendas_lag8_14': 38.6723995493804,
                    'vendas_lag15_21': 38.739992489673305,
                    'vendas_lag22_28': 38.83274502440857,
                },
            },
            'pastas_aprovadas': {
                'mu': 10.567480285392415,
                'media_dia_semana': {
                    'domingo': 2.840419947506562,
                    'quarta': 11.427138331573392,
                    'quinta': 11.132523759239705,
                    'segunda': 13.027690700104495,
                    'sexta': 12.704857444561776,
                    'sábado': 11.832105263157894,
                    'terça': 11.017763845350055,
                },
                'media_dia_mes': {
                    1: 10.833333333333332,
                    2: 13.678240740740739,
                    3: 10.625,
                    4: 12.067129629629628,
                    5: 10.009132420091323,
                    6: 10.538812785388128,
                    7: 9.863013698630136,
                    8: 10.084474885844749,
                    9: 10.577625570776254,
                    10: 8.666666666666666,
                    11: 10.173515981735159,
                    12: 10.598173515981735,
                    13: 10.410958904109588,
                    14: 8.611872146118722,
                    15: 10.605022831050228,
                    16: 10.068493150684931,
                    17: 10.239726027397259,
                    18: 9.28538812785388,
                    19: 8.465753424657533,
                    20: 11.118721461187214,
                    21: 10.168949771689496,
                    22: 9.214611872146119,
                    23: 11.801369863013699,
                    24: 9.116438356164384,
                    25: 10.43607305936073,
                    26: 12.426940639269404,
                    27: 12.205479452054794,
                    28: 11.844748858447486,
                    29: 11.35872235872236,
                    30: 12.055421686746987,
                    31: 10.699604743083004,
                },
                'media_mes': {
                    'abril': 13.979279279279279,
                    'agosto': 9.644701691896705,
                    'dezembro': 8.275501307759374,
                    'fevereiro': 10.61612284069098,
                    'janeiro': 11.496076721883174,
                    'julho': 9.930280957336109,
                    'junho': 10.953736654804272,
                    'maio': 12.797733217088059,
                    'março': 11.559721011333917,
                    'novembro': 8.864864864864865,
                    'outubro': 8.811682650392328,
                    'setembro': 9.799999999999999,
                },
                'lag_mu': {
                    'agendamentos_lag1_7': 536.7742395794216,
                    'agendamentos_lag8_14': 536.8396545249718,
                    'agendamentos_lag15_21': 536.5288021028914,
                    'agendamentos_lag22_28': 535.1260983852798,
                    'visitas_lag1_7': 198.6619601952685,
                    'visitas_lag8_14': 198.523695080736,
                    'visitas_lag15_21': 198.2122418325197,
                    'visitas_lag22_28': 197.6323695080736,
                    'pastas_lag1_7': 204.10874953060457,
                    'pastas_lag8_14': 203.90769808486667,
                    'pastas_lag15_21': 203.74787833270747,
                    'pastas_lag22_28': 203.64115659031165,
                    'pastas_aprovadas_lag1_7': 74.10544498685692,
                    'pastas_aprovadas_lag8_14': 74.19038678182501,
                    'pastas_aprovadas_lag15_21': 74.08396545249718,
                    'pastas_aprovadas_lag22_28': 74.17506571535863,
                    'vendas_lag1_7': 38.555013143071726,
                    'vendas_lag8_14': 38.6723995493804,
                    'vendas_lag15_21': 38.739992489673305,
                    'vendas_lag22_28': 38.83274502440857,
                },
            },
            'vendas': {
                'mu': 5.491625985730379,
                'media_dia_semana': {
                    'domingo': 2.662992125984252,
                    'quarta': 4.972016895459346,
                    'quinta': 4.947201689545935,
                    'segunda': 6.280564263322885,
                    'sexta': 5.666314677930307,
                    'sábado': 8.506315789473684,
                    'terça': 5.405433646812957,
                },
                'media_dia_mes': {
                    1: 6.768518518518517,
                    2: 7.564814814814815,
                    3: 6.939814814814813,
                    4: 7.002314814814814,
                    5: 5.369863013698629,
                    6: 5.643835616438356,
                    7: 5.034246575342466,
                    8: 5.477168949771689,
                    9: 4.917808219178082,
                    10: 3.34931506849315,
                    11: 4.529680365296803,
                    12: 4.771689497716895,
                    13: 5.511415525114155,
                    14: 4.1894977168949765,
                    15: 5.027397260273972,
                    16: 5.294520547945204,
                    17: 4.58904109589041,
                    18: 5.043378995433789,
                    19: 4.15296803652968,
                    20: 5.4908675799086755,
                    21: 5.068493150684931,
                    22: 4.470319634703197,
                    23: 4.963470319634703,
                    24: 5.691780821917808,
                    25: 6.547945205479452,
                    26: 5.502283105022831,
                    27: 5.968036529680364,
                    28: 6.705479452054794,
                    29: 6.2653562653562656,
                    30: 7.250602409638555,
                    31: 5.284584980237153,
                },
                'media_mes': {
                    'abril': 7.726126126126126,
                    'agosto': 4.54853072128228,
                    'dezembro': 4.543156059285092,
                    'fevereiro': 5.229366602687141,
                    'janeiro': 5.495204882301657,
                    'julho': 5.059313215400625,
                    'junho': 6.141459074733097,
                    'maio': 7.190061028770708,
                    'março': 5.924149956408021,
                    'novembro': 5.112612612612613,
                    'outubro': 4.284219703574543,
                    'setembro': 4.574774774774775,
                },
                'lag_mu': {
                    'agendamentos_lag1_7': 536.7742395794216,
                    'agendamentos_lag8_14': 536.8396545249718,
                    'agendamentos_lag15_21': 536.5288021028914,
                    'agendamentos_lag22_28': 535.1260983852798,
                    'visitas_lag1_7': 198.6619601952685,
                    'visitas_lag8_14': 198.523695080736,
                    'visitas_lag15_21': 198.2122418325197,
                    'visitas_lag22_28': 197.6323695080736,
                    'pastas_lag1_7': 204.10874953060457,
                    'pastas_lag8_14': 203.90769808486667,
                    'pastas_lag15_21': 203.74787833270747,
                    'pastas_lag22_28': 203.64115659031165,
                    'pastas_aprovadas_lag1_7': 74.10544498685692,
                    'pastas_aprovadas_lag8_14': 74.19038678182501,
                    'pastas_aprovadas_lag15_21': 74.08396545249718,
                    'pastas_aprovadas_lag22_28': 74.17506571535863,
                    'vendas_lag1_7': 38.555013143071726,
                    'vendas_lag8_14': 38.6723995493804,
                    'vendas_lag15_21': 38.739992489673305,
                    'vendas_lag22_28': 38.83274502440857,
                },
            },
        },
    },
    'totais_hist': {
        'agendamentos': 81282.40120165226,
        'visitas': 30094.151408186248,
        'pastas': 30925.2614344724,
        'pastas_aprovadas': 11212.096582801352,
        'vendas': 5826.615170859932,
    },
    'efeitos_lags_vendas': {
        'r2': 0.24870224559717446,
        'r2s_etapa': {
            'agendamentos': 0.24775124466944143,
            'visitas': 0.24272503308511717,
            'pastas': 0.2599467507748041,
            'pastas_aprovadas': 0.24438595385933515,
        },
        'lags': [
            7.0, 14.0, 21.0, 28.0,
        ],
        'resumo': [
            {
                'etapa': 'agendamentos',
                'label': 'Agendamentos',
                'lag_pico': 7,
                'efeito_pico': 0.0019669587277141467,
                'lag_meia_vida': 28,
                'efeito_lag1': 0.0019669587277141467,
                'efeito_acum': -0.0034514094976863764,
            },
            {
                'etapa': 'visitas',
                'label': 'Visitas',
                'lag_pico': 14,
                'efeito_pico': 0.007285848389944166,
                'lag_meia_vida': 28,
                'efeito_lag1': 0.004130254665291372,
                'efeito_acum': -0.008679391390556006,
            },
            {
                'etapa': 'pastas',
                'label': 'Pastas',
                'lag_pico': 7,
                'efeito_pico': 0.021527553626313714,
                'lag_meia_vida': 7,
                'efeito_lag1': 0.021527553626313714,
                'efeito_acum': 0.017291153061893477,
            },
            {
                'etapa': 'pastas_aprovadas',
                'label': 'Pastas aprovadas',
                'lag_pico': 28,
                'efeito_pico': 0.04155465400520727,
                'lag_meia_vida': 7,
                'efeito_lag1': 0.03718516864179573,
                'efeito_acum': 0.04155465400520727,
            },
        ],
        'perfis': {
            'agendamentos': [
                {
                    'lag': 7,
                    'inicio_lag': 1,
                    'efeito': 0.0019669587277141467,
                    'acumulado': 0.0019669587277141467,
                },
                {
                    'lag': 14,
                    'inicio_lag': 8,
                    'efeito': -0.0004681388327013951,
                    'acumulado': 0.0014988198950127516,
                },
                {
                    'lag': 21,
                    'inicio_lag': 15,
                    'efeito': -0.0028585121666814344,
                    'acumulado': -0.0013596922716686828,
                },
                {
                    'lag': 28,
                    'inicio_lag': 22,
                    'efeito': -0.0020917172260176936,
                    'acumulado': -0.0034514094976863764,
                },
            ],
            'visitas': [
                {
                    'lag': 7,
                    'inicio_lag': 1,
                    'efeito': 0.004130254665291372,
                    'acumulado': 0.004130254665291372,
                },
                {
                    'lag': 14,
                    'inicio_lag': 8,
                    'efeito': 0.0031555937246527934,
                    'acumulado': 0.007285848389944166,
                },
                {
                    'lag': 21,
                    'inicio_lag': 15,
                    'efeito': -0.0086744866459523,
                    'acumulado': -0.001388638256008135,
                },
                {
                    'lag': 28,
                    'inicio_lag': 22,
                    'efeito': -0.00729075313454787,
                    'acumulado': -0.008679391390556006,
                },
            ],
            'pastas': [
                {
                    'lag': 7,
                    'inicio_lag': 1,
                    'efeito': 0.021527553626313714,
                    'acumulado': 0.021527553626313714,
                },
                {
                    'lag': 14,
                    'inicio_lag': 8,
                    'efeito': -0.00027824337621382557,
                    'acumulado': 0.02124931025009989,
                },
                {
                    'lag': 21,
                    'inicio_lag': 15,
                    'efeito': -0.004973819052033632,
                    'acumulado': 0.016275491198066257,
                },
                {
                    'lag': 28,
                    'inicio_lag': 22,
                    'efeito': 0.001015661863827219,
                    'acumulado': 0.017291153061893477,
                },
            ],
            'pastas_aprovadas': [
                {
                    'lag': 7,
                    'inicio_lag': 1,
                    'efeito': 0.03718516864179573,
                    'acumulado': 0.03718516864179573,
                },
                {
                    'lag': 14,
                    'inicio_lag': 8,
                    'efeito': 0.0010596285003743409,
                    'acumulado': 0.03824479714217007,
                },
                {
                    'lag': 21,
                    'inicio_lag': 15,
                    'efeito': -0.006741923024754099,
                    'acumulado': 0.03150287411741597,
                },
                {
                    'lag': 28,
                    'inicio_lag': 22,
                    'efeito': 0.010051779887791294,
                    'acumulado': 0.04155465400520727,
                },
            ],
        },
    },
    'treino_inicio': '2023-07-01',
    'treino_fim': '2026-06-30',
    'n_obs': 1061,
    'treinado_em': '2026-07-19T04:10:21',
}
# END FUNIL_MODELO_PRODUCAO
FUNIL_CORES_DRIVER = {
    "agendamentos": "#04428f",
    "visitas": "#cb0935",
    "pastas": "#0f766e",
    "pastas_aprovadas": "#b45309",
}
FUNIL_CORES_NIVEIS = ["#022654", "#04428f", "#1e60b3", "#cb0935", "#9e0828"]
# Ordem das fases de oportunidade (Track Funnel — pipeline comercial RJ)
TRACK_FUNIL_FASES = (
    "Aguardando aprovação comercial",
    "Aguardando atendimento Corretor",
    "Em atendimento",
    "Visita Agendada",
    "Visita Realizada",
    "Em elaboração",
    "Em Análise de Crédito",
    "Análise de Crédito Realizada",
    "Proposta",
    "Proposta Reprovada",
    "Proposta Aprovada",
    "Enviado Aprovação Comissões",
    "Reprovado Comissões",
    "Aprovado Comissões",
    "Análise SAFI",
    "Aprovado SAFI",
    "Rejeitado SAFI",
    "Enviado Aprovação Pré Soluto",
    "Aprovado Pré soluto",
    "Rejeitado Pro Soluto",
    "Contrato gerado",
    "Contrato comunicado",
    "Contrato com pendência comercial",
    "Enviado para Aprovação",
    "Reprovado",
    "Rejeitado",
    "Em cessão",
    "Fechado e ganho",
    "Fechado e perdido",
    "Cancelado",
    "Distratado",
    "Repassado",
)
TRACK_FUNIL_CORES = [
    "#b8d4f0", "#9ec5eb", "#7ab3e8", "#5a9fe3",
    "#3d8bdc", "#2a7ad4", "#1e60b3", "#04428f",
    "#033a7a", "#022654", "#4a5568",
]
ORIGENS_NUCLEO_DIGITAL = frozenset({
    "Lead ADS", "WhatsApp", "Portal Vertical", "Cadastre-se", "Chat", "Chatbot",
    "Simulador financeiro", "Simulador Virtual", "Landing page oferta",
    "Landing Page Breve Lancamento", "LP Compra Online", "LP Fica na Boa Direcional",
    "LP MCMV", "LP Squad Dire", "Messenger FB", "Blog Direcional", "Blog Riva",
    "Agendar visita", "Fale conosco", "Fale com o consultor", "Google Meu Negócio",
    "DV ON", "RVON", "RV ON", "Squad Dire", "Origem SDR",
})
ORIGENS_NUCLEO_DIGITAL_NORM = frozenset(x.lower() for x in ORIGENS_NUCLEO_DIGITAL)
ALIASES_FASE_OPORTUNIDADE = ["Fase", "StageName"]
ALIASES_OPP_FECHADA = ["Fechada", "IsClosed"]
ALIASES_OPP_MUDANCA_FASE = [
    "Data mudança fase", "Data mudanca fase", "LastStageChangeDate",
    "Data da última mudança de fase",
]
ALIASES_EMPREENDIMENTO = [
    "Empreendimento", "Obra", "Nome do Empreendimento", "Nome do empreendimento",
]
COLUNAS_PASTAS_ALIASES = [
    "Data Primeiro Envio Análise", "Data Primeiro Envio Analise",
    "Data do Primeiro Envio Análise", "Data do Primeiro Envio Analise",
    "Primeiro Envio Análise", "Primeiro Envio Analise",
    "Data 1º Envio Análise", "Data 1o Envio Analise",
    "Data da Análise", "Data da Analise", "Data Análise", "Data Analise",
]
COLUNAS_PASTAS_APROV_ALIASES = [
    "Data Aprovação SAFI", "Data Aprovacao SAFI",
    "Data da Aprovação SAFI", "Data da Aprovacao SAFI",
    "Data de Aprovação SAFI", "Data de Aprovacao SAFI",
    "Aprovação SAFI", "Aprovacao SAFI",
    "Data Aprov. SAFI", "Data Aprov SAFI",
    "SAFI Approval Date", "Approval Date SAFI",
]

_DIR_APP = Path(__file__).resolve().parent
LOGO_TOPO_ARQUIVO = "502.57_LOGO DIRECIONAL_V2F-01.png"
FAVICON_ARQUIVO = "502.57_LOGO D_COR_V3F.png"
FUNDO_CADASTRO_ARQUIVO = "fundo_cadastrorh.jpg"
URL_LOGO_DIRECIONAL_EMAIL = "https://logodownload.org/wp-content/uploads/2021/04/direcional-engenharia-logo.png"

# Paleta alinhada à Ficha Credenciamento / Vendas RJ
COR_AZUL_ESC = "#04428f"
COR_VERMELHO = "#cb0935"
COR_VERMELHO_ESCURO = "#9e0828"
COR_FUNDO_CARD = "rgba(255, 255, 255, 0.78)"
COR_BORDA = "#eef2f6"
COR_TEXTO_PRETO = "#000000"
COR_TEXTO_MUTED = "#000000"
COR_TEXTO_LABEL = "#000000"
COR_INPUT_BG = "#f0f2f6"

MESES_TEXTO_MAP = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
    "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4, "maio": 5, "junho": 6,
    "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
}


def _hex_rgb_triplet(hex_color: str) -> str:
    """Converte #RRGGBB em 'r, g, b' para uso em rgba(...)."""
    x = (hex_color or "").strip().lstrip("#")
    if len(x) != 6:
        return "0, 0, 0"
    return f"{int(x[0:2], 16)}, {int(x[2:4], 16)}, {int(x[4:6], 16)}"


RGB_AZUL_CSS = _hex_rgb_triplet(COR_AZUL_ESC)
RGB_VERMELHO_CSS = _hex_rgb_triplet(COR_VERMELHO)


# -----------------------------------------------------------------------------
# Funções de Design (Ficha Direcional)
# -----------------------------------------------------------------------------
def _resolver_png_raiz(nome: str) -> Path | None:
    """Procura o PNG na pasta do app e na pasta pai."""
    for base in (_DIR_APP, _DIR_APP.parent):
        p = base / nome
        if p.is_file():
            return p
    return None

def _resolver_imagem_fundo_local(nome: str) -> Path | None:
    """Imagem JPG/PNG na pasta do app ou na pasta pai."""
    for base in (_DIR_APP, _DIR_APP.parent):
        for ext in (".jpg", ".jpeg", ".JPG", ".JPEG", ".png", ".PNG"):
            stem = Path(nome).stem
            p = base / f"{stem}{ext}"
            if p.is_file():
                return p
        p = base / nome
        if p.is_file():
            return p
    return None

def _css_url_fundo_cadastro() -> str:
    """String para `url(...)` no CSS: data-URL or URL https (fallback)."""
    p = _resolver_imagem_fundo_local(FUNDO_CADASTRO_ARQUIVO)
    if p and p.is_file():
        try:
            raw = p.read_bytes()
            suf = p.suffix.lower()
            mime = "image/jpeg" if suf in (".jpg", ".jpeg") else "image/png"
            b64 = base64.b64encode(raw).decode("ascii")
            return f"data:{mime};base64,{b64}"
        except OSError:
            pass
    return (
        "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab"
        "?auto=format&fit=crop&w=1920&q=80"
    )

def _logo_arquivo_local() -> str | None:
    p_topo = _resolver_png_raiz(LOGO_TOPO_ARQUIVO)
    if p_topo:
        return str(p_topo)
    for name in ("logo_direcional.png", "logo_direcional.jpg", "logo_direcional.jpeg", "logo.png"):
        p = _DIR_APP / "assets" / name
        if p.is_file():
            return str(p)
    return None

def _logo_url_secrets() -> str | None:
    try:
        if hasattr(st, "secrets"):
            b = st.secrets.get("branding")
            if isinstance(b, dict):
                u = (b.get("LOGO_URL") or "").strip()
                if u:
                    return u
    except Exception:
        pass
    return None

def _logo_url_drive_por_id_arquivo() -> str | None:
    fid = (os.environ.get("DIrecIONAL_LOGO_FILE_ID") or "").strip()
    if len(fid) < 10:
        return None
    return f"https://drive.google.com/uc?export=view&id={fid}"

def _exibir_logo_topo() -> None:
    """Logo centralizada no topo."""
    path = _logo_arquivo_local()
    url = _logo_url_secrets() or _logo_url_drive_por_id_arquivo()
    try:
        if path:
            ext = Path(path).suffix.lower().lstrip(".")
            mime = "image/png" if ext == "png" else "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            st.markdown(
                f'<div class="ficha-logo-wrap"><img src="data:{mime};base64,{b64}" alt="Direcional" /></div>',
                unsafe_allow_html=True,
            )
            return
        if url:
            st.markdown(
                f'<div class="ficha-logo-wrap"><img src="{html.escape(url)}" alt="Direcional" /></div>',
                unsafe_allow_html=True,
            )
    except Exception:
        pass

def _cabecalho_pagina() -> None:
    _exibir_logo_topo()
    st.markdown(
        f'<div class="ficha-hero-stack">'
        f'<div class="ficha-hero">'
        f'<p class="ficha-title">Acompanhamento de metas de vendas</p>'
        f"</div>"
        f'<div class="ficha-hero-bar-wrap" aria-hidden="true">'
        f'<div class="ficha-hero-bar"></div>'
        f"</div>"
        f"</div>",
        unsafe_allow_html=True,
    )


def aplicar_estilo() -> None:
    bg_url = _css_url_fundo_cadastro()
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800;900&family=Inter:wght@400;500;600;700&display=swap');
        @keyframes fichaFadeIn {{
            from {{ opacity: 0; transform: translateY(18px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        @keyframes fichaShimmer {{
            0% {{ background-position: 0% 50%; }}
            100% {{ background-position: 200% 50%; }}
        }}
        html, body, :root, [data-testid="stApp"] {{
            color-scheme: light !important;
        }}
        /* Tabelas sempre claras, mesmo com Windows em modo escuro */
        .stDataFrame,
        [data-testid="stDataFrame"],
        [data-testid="stDataFrame"] > div,
        [data-testid="stDataFrameResizable"],
        [data-testid="stTable"],
        [data-testid="stTable"] table {{
            color-scheme: light !important;
            background-color: #ffffff !important;
        }}
        [data-testid="stTable"] th,
        [data-testid="stTable"] td {{
            background-color: #ffffff !important;
            color: #000000 !important;
        }}
        html, body {{
            font-family: 'Inter', sans-serif;
            color: {COR_TEXTO_LABEL};
            background: transparent !important;
            background-color: transparent !important;
        }}
        .stApp,
        [data-testid="stApp"] {{
            background:
                linear-gradient(135deg, rgba({RGB_AZUL_CSS}, 0.82) 0%, rgba(30, 58, 95, 0.55) 38%, rgba({RGB_VERMELHO_CSS}, 0.22) 72%, rgba(15, 23, 42, 0.45) 100%),
                url("{bg_url}") center / cover no-repeat !important;
            background-attachment: scroll !important;
            background-color: transparent !important;
        }}
        [data-testid="stAppViewContainer"] {{
            background: transparent !important;
            background-color: transparent !important;
        }}
        header[data-testid="stHeader"],
        [data-testid="stHeader"] {{
            background: transparent !important;
            background-color: transparent !important;
            background-image: none !important;
            border: none !important;
            box-shadow: none !important;
            backdrop-filter: none !important;
            -webkit-backdrop-filter: none !important;
        }}
        [data-testid="stHeader"] > div,
        [data-testid="stHeader"] header {{
            background: transparent !important;
            background-color: transparent !important;
            background-image: none !important;
            box-shadow: none !important;
        }}
        [data-testid="stDecoration"] {{
            background: transparent !important;
            background-color: transparent !important;
        }}
        [data-testid="stSidebar"] {{ display: none !important; }}
        [data-testid="stSidebarCollapsedControl"] {{ display: none !important; }}
        [data-testid="stToolbar"] {{
            background: transparent !important;
            background-color: transparent !important;
            background-image: none !important;
            border: none !important;
            border-radius: 0 !important;
            box-shadow: none !important;
            color: rgba(255, 255, 255, 0.92) !important;
        }}
        [data-testid="stToolbar"] button,
        [data-testid="stToolbar"] a {{
            color: rgba(255, 255, 255, 0.92) !important;
            background: transparent !important;
            background-color: transparent !important;
        }}
        [data-testid="stHeader"] button {{
            background: transparent !important;
            background-color: transparent !important;
        }}
        [data-testid="stToolbar"] svg {{
            fill: currentColor !important;
            color: inherit !important;
        }}
        [data-testid="stToolbar"] svg path[stroke] {{
            stroke: currentColor !important;
        }}
        [data-testid="stToolbar"] button:hover,
        [data-testid="stToolbar"] a:hover,
        [data-testid="stHeader"] button:hover {{
            background: rgba(255, 255, 255, 0.12) !important;
        }}
        [data-testid="stMain"] {{
            padding-left: clamp(14px, 5vw, 56px) !important;
            padding-right: clamp(14px, 5vw, 56px) !important;
            padding-top: clamp(12px, 3.5vh, 40px) !important;
            padding-bottom: clamp(14px, 4vh, 44px) !important;
            box-sizing: border-box !important;
        }}
        section.main > div {{
            padding-top: 0.25rem !important;
            padding-bottom: 0.35rem !important;
        }}
        .block-container {{
            max-width: 1700px !important;
            margin-left: auto !important;
            margin-right: auto !important;
            margin-top: clamp(4px, 1vh, 14px) !important;
            margin-bottom: clamp(4px, 1vh, 14px) !important;
            padding: 1.45rem 2.25rem 1.55rem 2.25rem !important;
            background: rgba(255, 255, 255, 0.78) !important;
            backdrop-filter: blur(18px) saturate(1.15);
            -webkit-backdrop-filter: blur(18px) saturate(1.15);
            border-radius: 24px !important;
            border: 1px solid rgba(255, 255, 255, 0.45) !important;
            box-shadow:
                0 4px 6px -1px rgba({RGB_AZUL_CSS}, 0.06),
                0 24px 48px -12px rgba({RGB_AZUL_CSS}, 0.18),
                inset 0 1px 0 rgba(255, 255, 255, 0.55) !important;
            animation: fichaFadeIn 0.7s cubic-bezier(0.22, 1, 0.36, 1) both;
        }}
        /* Títulos de seção do dashboard: azuis (inclui spans internos do Streamlit) */
        h1, h2, h3, h4,
        h1 *, h2 *, h3 *, h4 *,
        [data-testid="stMarkdownContainer"] h1,
        [data-testid="stMarkdownContainer"] h2,
        [data-testid="stMarkdownContainer"] h3,
        [data-testid="stMarkdownContainer"] h4,
        [data-testid="stMarkdownContainer"] h1 *,
        [data-testid="stMarkdownContainer"] h2 *,
        [data-testid="stMarkdownContainer"] h3 *,
        [data-testid="stMarkdownContainer"] h4 *,
        [data-testid="stHeadingWithAction"],
        [data-testid="stHeadingWithAction"] *,
        .stHeading, .stHeading * {{
            font-family: 'Montserrat', sans-serif !important;
            color: {COR_AZUL_ESC} !important;
            font-weight: 800 !important;
        }}
        h1, h2, h3, h4,
        [data-testid="stHeadingWithAction"],
        .stHeading {{
            text-align: center !important;
        }}
        /* Títulos de gráficos (#####): pretos */
        h5, h6,
        h5 *, h6 *,
        [data-testid="stMarkdownContainer"] h5,
        [data-testid="stMarkdownContainer"] h6,
        [data-testid="stMarkdownContainer"] h5 *,
        [data-testid="stMarkdownContainer"] h6 * {{
            font-family: 'Montserrat', sans-serif !important;
            color: {COR_TEXTO_PRETO} !important;
            font-weight: 700 !important;
            text-align: center !important;
        }}
        /* Texto geral do dashboard: preto (não sobrescreve títulos) */
        .block-container,
        .block-container > div p,
        .block-container label,
        .block-container li,
        [data-testid="stMarkdownContainer"] > p,
        [data-testid="stCaption"],
        [data-testid="stCaptionContainer"],
        [data-testid="stCaptionContainer"] p,
        [data-testid="stWidgetLabel"] p,
        [data-testid="stWidgetLabel"] label,
        div[data-baseweb="select"] span,
        .stSelectbox label,
        .stMultiSelect label {{
            color: {COR_TEXTO_PRETO} !important;
        }}
        .block-container span:not(h1 span):not(h2 span):not(h3 span):not(h4 span) {{
            color: {COR_TEXTO_PRETO};
        }}
        h1 span, h2 span, h3 span, h4 span,
        [data-testid="stHeadingWithAction"] span {{
            color: {COR_AZUL_ESC} !important;
        }}
        .ficha-logo-wrap {{
            text-align: center;
            padding: 0.1rem 0 0.45rem 0;
        }}
        .ficha-logo-wrap img {{
            max-height: 72px;
            width: auto;
            max-width: min(280px, 85vw);
            height: auto;
            object-fit: contain;
            display: inline-block;
            vertical-align: middle;
        }}
        .ficha-hero-stack {{
            width: 100%;
            max-width: 100%;
            margin-bottom: 0.35rem;
            box-sizing: border-box;
        }}
        .ficha-hero {{
            text-align: center;
            padding: 0.5rem 0 0 0;
            margin: 0 auto 0 auto;
            max-width: 640px;
            animation: fichaFadeIn 0.85s cubic-bezier(0.22, 1, 0.36, 1) 0.1s both;
        }}
        .ficha-hero .ficha-title {{
            font-family: 'Montserrat', sans-serif;
            font-size: clamp(1.35rem, 3.5vw, 1.75rem);
            font-weight: 900;
            color: {COR_AZUL_ESC};
            margin: 0;
            line-height: 1.25;
            letter-spacing: -0.02em;
        }}
        .ficha-hero-bar-wrap {{
            width: 100%;
            max-width: 100%;
            margin: clamp(0.85rem, 2.4vw, 1.2rem) 0;
            padding: 0;
            box-sizing: border-box;
        }}
        .ficha-hero-bar {{
            height: 4px;
            width: 100%;
            margin: 0;
            border-radius: 999px;
            background: linear-gradient(90deg, {COR_AZUL_ESC}, {COR_VERMELHO}, {COR_AZUL_ESC});
            background-size: 200% 100%;
            animation: fichaShimmer 4s ease-in-out infinite alternate;
        }}
        .vel-kpi-row {{
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            margin-bottom: 1.25rem;
        }}
        .vel-kpi {{
            flex: 1 1 20%;
            background: linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(250,251,252,0.9) 100%);
            border: 1px solid rgba(226, 232, 240, 0.9);
            border-radius: 14px;
            padding: 14px 16px;
            text-align: center;
            box-shadow: 0 2px 8px rgba({RGB_AZUL_CSS}, 0.06);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        .vel-kpi:hover {{
            transform: translateY(-4px);
            box-shadow: 0 10px 20px -5px rgba({RGB_AZUL_CSS}, 0.15);
        }}
        .vel-kpi .lbl {{
            font-size: 0.72rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: {COR_TEXTO_PRETO};
            opacity: 0.85;
        }}
        /* Valores dos boxes (rótulos de dados): mantêm azul / vermelho */
        .vel-kpi .val {{
            font-family: 'Montserrat', sans-serif;
            font-size: 1.35rem;
            font-weight: 800;
            color: {COR_AZUL_ESC} !important;
            margin-top: 6px;
        }}
        .vel-kpi .val--red {{ color: {COR_VERMELHO} !important; }}
        div[data-testid="stMetric"] {{
            background: rgba(255,255,256,0.6);
            padding: 12px;
            border-radius: 12px;
            border: 1px solid {COR_BORDA};
        }}
        div[data-baseweb="input"] {{
            border-radius: 10px !important;
            border: 1px solid #e2e8f0 !important;
            background-color: {COR_INPUT_BG} !important;
            transition: border-color 0.2s ease, box-shadow 0.2s ease;
        }}
        div[data-baseweb="input"]:focus-within {{
            border-color: rgba({RGB_AZUL_CSS}, 0.35) !important;
            box-shadow: 0 0 0 3px rgba({RGB_AZUL_CSS}, 0.08) !important;
        }}
        /* Títulos em azul Direcional */
        h1, h2, h3, h4, h5, h6,
        [data-testid="stMarkdownContainer"] h1,
        [data-testid="stMarkdownContainer"] h2,
        [data-testid="stMarkdownContainer"] h3,
        [data-testid="stMarkdownContainer"] h4,
        [data-testid="stMarkdownContainer"] h5,
        [data-testid="stMarkdownContainer"] h6,
        [data-testid="stSubheader"],
        .stSubheader {{
            color: {COR_AZUL_ESC} !important;
            font-family: 'Montserrat', sans-serif !important;
        }}
        /* Barra de abas centralizada */
        .stTabs [data-baseweb="tab-list"] {{
            justify-content: center !important;
            gap: 0.35rem;
        }}
        .stTabs [data-baseweb="tab"] {{
            font-family: 'Montserrat', sans-serif !important;
            font-weight: 700 !important;
            color: {COR_AZUL_ESC} !important;
        }}
        .stTabs [aria-selected="true"] {{
            border-bottom: 3px solid {COR_VERMELHO} !important;
            color: {COR_AZUL_ESC} !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# Lógicas de Integração e Gsheets
# -----------------------------------------------------------------------------

def _secrets_connections_gsheets() -> Dict[str, Any]:
    try:
        sec = st.secrets
        if hasattr(sec, "get") and sec.get("connections"):
            g = sec["connections"].get("gsheets")
            if g is not None:
                return dict(g)
    except Exception:
        pass
    return {}


def _normalizar_private_key_toml(pk: str) -> str:
    s = (pk or "").strip()
    if not s:
        return s
    if "\\n" in s and "\n" not in s:
        s = s.replace("\\n", "\n")
    return s


def montar_service_account_info(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if not raw:
        return None
    chaves = (
        "type", "project_id", "private_key_id", "private_key",
        "client_email", "client_id", "auth_uri", "token_uri",
        "auth_provider_x509_cert_url", "client_x509_cert_url",
    )
    out: Dict[str, Any] = {}
    for k in chaves:
        v = raw.get(k)
        if v is None: continue
        if isinstance(v, str): v = v.strip()
        if v == "": continue
        out[k] = v
    if "private_key" in out:
        out["private_key"] = _normalizar_private_key_toml(str(out["private_key"]))
    if "private_key" not in out or "client_email" not in out:
        return None
    typ = str(out.get("type") or "").strip()
    if not typ:
        out["type"] = "service_account"
    if "token_uri" not in out:
        out["token_uri"] = "https://oauth2.googleapis.com/token"
    if "auth_uri" not in out:
        out["auth_uri"] = "https://accounts.google.com/o/oauth2/auth"
    return out


def spreadsheet_id_de_secrets(cfg: Dict[str, Any]) -> str:
    for k in ("spreadsheet_id", "SPREADSHEET_ID", "spreadsheet", "planilha_id"):
        v = str(cfg.get(k) or "").strip()
        if v: return v
    return SPREADSHEET_ID


def valores_para_dataframe(rows: List[List[str]]) -> pd.DataFrame:
    if not rows: return pd.DataFrame()
    header = [str(c).strip() for c in rows[0]]
    w = len(header)
    if w == 0: return pd.DataFrame()
    body = rows[1:]
    if not body: return pd.DataFrame(columns=header)
    norm: List[List[str]] = []
    for r in body:
        cells = [str(c) for c in r]
        if len(cells) < w: cells = cells + [""] * (w - len(cells))
        else: cells = cells[:w]
        norm.append(cells)
    return pd.DataFrame(norm, columns=header)


def ler_aba_gsheets(
    service_account_info: Dict[str, Any],
    spreadsheet_id: str,
    worksheet: str,
    aliases: Optional[Tuple[str, ...]] = None,
) -> pd.DataFrame:
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(spreadsheet_id.strip())
    nome = worksheet.strip()

    def _resolver_aba() -> Any:
        candidatos: List[str] = [nome]
        if aliases:
            candidatos.extend(list(aliases))
        vistos: set = set()
        ordem: List[str] = []
        for c in candidatos:
            c = c.strip()
            if not c or c.lower() in vistos:
                continue
            vistos.add(c.lower())
            ordem.append(c)
        titulos_map = {w.title.strip().lower(): w for w in sh.worksheets()}
        for cand in ordem:
            cl = cand.lower()
            if cl in titulos_map:
                return titulos_map[cl]
        for cand in ordem:
            cl = cand.lower()
            for tl, w in titulos_map.items():
                if tl == cl:
                    return w
            for tl, w in titulos_map.items():
                if tl.endswith(cl) or cl in tl:
                    return w
        titulos = [w.title for w in sh.worksheets()]
        raise gspread.WorksheetNotFound(
            f"Aba {nome!r} não encontrada. Abas: {titulos}"
        )

    ws = _resolver_aba()
    return valores_para_dataframe(ws.get_all_values())


def _fingerprint_credenciais(info: Dict[str, Any]) -> str:
    pk = str(info.get("private_key") or "")
    return str(hash(pk))[-12:] if pk else "0"


@st.cache_data(ttl=300, show_spinner=False)
def ler_planilha_aba_df(
    spreadsheet_id: str,
    worksheet: str,
    _cred_fp: str,
    aliases: Optional[Tuple[str, ...]] = None,
) -> pd.DataFrame:
    raw = _secrets_connections_gsheets()
    info = montar_service_account_info(raw)
    if not info: raise ValueError("Credenciais [connections.gsheets] ausentes ou incompleta.")
    return ler_aba_gsheets(info, spreadsheet_id, worksheet, aliases=aliases)


def _cabecalho_tem_coluna(header: List[str], aliases: List[str]) -> bool:
    lows = [str(h).strip().lower() for h in header]
    for a in aliases:
        al = a.strip().lower()
        if any(h == al or al in h for h in lows):
            return True
    return False


def _df_parece_pastas(df: pd.DataFrame) -> bool:
    if df is None or df.empty:
        return False
    return bool(
        achar_coluna(df, COLUNAS_PASTAS_ALIASES)
        or achar_coluna(df, COLUNAS_PASTAS_APROV_ALIASES)
    )


@st.cache_data(ttl=300, show_spinner=False)
def carregar_df_pastas_funil(
    spreadsheet_funil_id: str,
    spreadsheet_principal_id: str,
    spreadsheet_pastas_id: str,
    _cred_fp: str,
) -> Tuple[pd.DataFrame, str]:
    """
    Localiza a base de pastas:
      1) abas candidatas (BASE, Pastas, ...)
      2) varredura de abas com colunas Data Criação Pasta / Data Aprovação SAFI
    Procura na planilha de pastas (se informada), na de agendamentos e na principal.
    Retorna (df, origem descritiva).
    """
    import gspread
    from google.oauth2.service_account import Credentials

    raw = _secrets_connections_gsheets()
    info = montar_service_account_info(raw)
    if not info:
        raise ValueError("Credenciais [connections.gsheets] ausentes ou incompleta.")

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)

    ids: List[str] = []
    for sid in (spreadsheet_pastas_id, spreadsheet_funil_id, spreadsheet_principal_id):
        s = (sid or "").strip()
        if s and s not in ids:
            ids.append(s)

    # 1) candidatos por nome
    for sid in ids:
        try:
            sh = gc.open_by_key(sid)
        except Exception:
            continue
        titulos = {w.title.strip().lower(): w.title for w in sh.worksheets()}
        for cand in ABA_PASTAS_CANDIDATAS:
            real = titulos.get(cand.strip().lower())
            if not real:
                continue
            try:
                df = valores_para_dataframe(sh.worksheet(real).get_all_values())
                df = normalizar_colunas(df)
                if _df_parece_pastas(df):
                    return df, f"{real} ({sid[:8]}…)"
            except Exception:
                continue

    # 2) varredura por colunas no cabeçalho
    for sid in ids:
        try:
            sh = gc.open_by_key(sid)
        except Exception:
            continue
        for w in sh.worksheets():
            try:
                header = [str(c) for c in (w.row_values(1) or [])]
            except Exception:
                continue
            if not (
                _cabecalho_tem_coluna(header, COLUNAS_PASTAS_ALIASES)
                or _cabecalho_tem_coluna(header, COLUNAS_PASTAS_APROV_ALIASES)
            ):
                continue
            try:
                df = valores_para_dataframe(w.get_all_values())
                df = normalizar_colunas(df)
                if _df_parece_pastas(df):
                    return df, f"{w.title} ({sid[:8]}…)"
            except Exception:
                continue

    return pd.DataFrame(), ""


def parse_valor_br(val: Any) -> float:
    if val is None or (isinstance(val, float) and pd.isna(val)): return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip().replace("R$", "").replace(" ", "").replace("\xa0", "")
    if not s or s.lower() == "nan" or s.lower() == "null": return 0.0
    s = re.sub(r"[^\d.,\-]", "", s)
    if not s: return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return 0.0


def canal_de_imobiliaria(val: Any) -> str:
    """Canal = LEFT(Imobiliária, 3) — mesma regra da planilha BD Vendas."""
    s = str(val or "").strip().upper()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    return s[:3]


def extrair_mes_da_data_venda(val: Any) -> Optional[int]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    dt = parse_data_serie(pd.Series([val])).iloc[0]
    if pd.notna(dt):
        m = int(dt.month)
        if 1 <= m <= 12:
            return m
    s = str(val).strip()
    if not s or s in ("nan", "null", ""):
        return None

    if "/" in s:
        parts = s.split("/")
        if len(parts) >= 2 and parts[1].isdigit():
            m = int(parts[1])
            if 1 <= m <= 12:
                return m

    for k, v in MESES_TEXTO_MAP.items():
        if k in s.lower():
            return v
    return None


def extrair_ano_da_data_venda(val: Any) -> Optional[int]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    dt = parse_data_serie(pd.Series([val])).iloc[0]
    if pd.notna(dt):
        y = int(dt.year)
        if y > 2000:
            return y
    s = str(val).strip()
    if not s or s in ("nan", "null", ""):
        return None

    if "/" in s:
        parts = s.split("/")
        if len(parts) >= 3:
            ano_str = re.sub(r"[^\d]", "", parts[2].split()[0])
            if len(ano_str) == 4 and ano_str.isdigit():
                return int(ano_str)

    cleaned = re.sub(r"[^\d]", "", s)
    if len(cleaned) >= 4:
        if len(cleaned) == 8 and cleaned[:4].isdigit():
            ano = int(cleaned[:4])
            if ano > 2000:
                return ano
        ano = int(cleaned[-4:])
        if ano > 2000:
            return ano
    return None


def extrair_mes_looker(val: Any) -> Optional[int]:
    """Mês a partir de colunas Looker / Mês Venda (número, ISO ou texto)."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    n = pd.to_numeric(val, errors="coerce")
    if pd.notna(n):
        m = int(n)
        if 1 <= m <= 12:
            return m
    return extrair_mes_da_data_venda(val)


def extrair_ano_looker(val: Any) -> Optional[int]:
    """Ano a partir de colunas Looker / Ano da Venda (número, ISO ou texto)."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    n = pd.to_numeric(val, errors="coerce")
    if pd.notna(n):
        y = int(n)
        if y > 2000:
            return y
    return extrair_ano_da_data_venda(val)


def aplicar_mes_ano_vendas(
    df: pd.DataFrame,
    cols_data: List[str],
    col_mes_venda: Optional[str] = None,
    col_ano_venda: Optional[str] = None,
    col_mes_looker: Optional[str] = None,
) -> Tuple[pd.Series, pd.Series]:
    """
    Preenche _mes/_ano coalescendo datas ISO do SF (Data da venda, Contrato gerado em)
    e campos Mês Venda / Ano da Venda quando a data principal vem vazia.
    """
    mes = pd.Series(pd.NA, index=df.index, dtype="Int64")
    ano = pd.Series(pd.NA, index=df.index, dtype="Int64")

    for col in cols_data:
        if not col or col not in df.columns:
            continue
        m = df[col].map(extrair_mes_da_data_venda)
        a = df[col].map(extrair_ano_da_data_venda)
        mes = mes.fillna(m.astype("Int64"))
        ano = ano.fillna(a.astype("Int64"))

    if col_mes_venda and col_mes_venda in df.columns:
        mes = mes.fillna(df[col_mes_venda].map(extrair_mes_looker).astype("Int64"))
    if col_mes_looker and col_mes_looker in df.columns:
        mes = mes.fillna(df[col_mes_looker].map(extrair_mes_looker).astype("Int64"))

    if col_ano_venda and col_ano_venda in df.columns:
        ano = ano.fillna(df[col_ano_venda].map(extrair_ano_looker).astype("Int64"))

    return mes, ano


def _serie_data_contrato(df: pd.DataFrame, col_contrato: Optional[str]) -> pd.Series:
    """Datetime de contrato para filtros de competência (fallback quando _mes/_ano vazios)."""
    if not col_contrato or col_contrato not in df.columns:
        return pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    return parse_data_serie(df[col_contrato])


def filtrar_vendas_competencia(
    df: pd.DataFrame,
    anos_sel: List[int],
    meses_sel: List[int],
    col_contrato: Optional[str] = None,
) -> pd.DataFrame:
    """
    Filtra vendas por ano/mês de competência.
    Usa _ano/_mes e, quando vazios, a data de Contrato gerado em.
    """
    out = df.copy()

    if anos_sel:
        dt_contrato = _serie_data_contrato(out, col_contrato)
        mask_ano = out["_ano"].isin(anos_sel)
        if dt_contrato.notna().any():
            mask_ano = mask_ano | (out["_ano"].isna() & dt_contrato.dt.year.isin(anos_sel))
        out = out.loc[mask_ano].copy()

    if meses_sel:
        # Recalcula datas após filtro de ano (índice de out muda).
        dt_contrato = _serie_data_contrato(out, col_contrato)
        mask_mes = out["_mes"].isin(meses_sel)
        if dt_contrato.notna().any():
            mask_mes = mask_mes | (out["_mes"].isna() & dt_contrato.dt.month.isin(meses_sel))
        out = out.loc[mask_mes].copy()

    return out


def _norm_txt_col(s: Any) -> str:
    """Normaliza texto de coluna para match sem acento/caixa."""
    import unicodedata
    t = str(s or "").strip().lower()
    t = unicodedata.normalize("NFKD", t)
    return "".join(c for c in t if not unicodedata.combining(c))


def achar_coluna(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    cols_norm = {_norm_txt_col(c): c for c in cols}
    # 1) match exato (com e sem acento)
    for a in aliases:
        al = str(a).strip().lower()
        for c in cols:
            if al == str(c).strip().lower():
                return c
        an = _norm_txt_col(a)
        if an in cols_norm:
            return cols_norm[an]
    # 2) substring
    for a in aliases:
        al = str(a).strip().lower()
        an = _norm_txt_col(a)
        for c in cols:
            cl = str(c).strip().lower()
            cn = _norm_txt_col(c)
            if al and (al in cl or an in cn):
                return c
    return None


def achar_coluna_aprovacao_safi(df: pd.DataFrame) -> Optional[str]:
    """Localiza a coluna Data Aprovação SAFI (prioridade: contém 'safi' + 'aprov')."""
    col = achar_coluna(df, COLUNAS_PASTAS_APROV_ALIASES)
    if col:
        return col
    if df is None or df.empty:
        return None
    candidatas: List[str] = []
    for c in df.columns:
        cn = _norm_txt_col(c)
        if "safi" in cn and "aprov" in cn:
            candidatas.append(c)
        elif "safi" in cn and "data" in cn:
            candidatas.append(c)
    return candidatas[0] if candidatas else None


def achar_coluna_primeiro_envio_analise(df: pd.DataFrame) -> Optional[str]:
    """Localiza Data Primeiro Envio Análise (prioridade: 'primeiro' + 'envio')."""
    col = achar_coluna(df, COLUNAS_PASTAS_ALIASES)
    if col and "primeiro" in _norm_txt_col(col) and "envio" in _norm_txt_col(col):
        return col
    if df is None or df.empty:
        return None
    for c in df.columns:
        cn = _norm_txt_col(c)
        if "primeiro" in cn and "envio" in cn:
            return c
        if "1o" in cn and "envio" in cn:
            return c
        if "1" in cn and "envio" in cn and "analise" in cn:
            return c
    # fallback: aliases genéricos (ex.: Data da Análise) se não houver coluna específica
    return col


ALIASES_VENDA_COMERCIAL = [
    "Venda Comercial?", "Venda Comercial", "Venda comercial?",
    "Venda comercial", "Comercial?",
]


def filtrar_vendas_comerciais(df: pd.DataFrame) -> pd.DataFrame:
    """Mantém apenas vendas comerciais (Venda Comercial? = 1 / SIM / TRUE)."""
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    col = achar_coluna(df, ALIASES_VENDA_COMERCIAL)
    if not col:
        return df
    mask = (
        (pd.to_numeric(df[col], errors="coerce") == 1)
        | (df[col].astype(str).str.strip().str.upper().isin(["SIM", "TRUE", "1", "1.0"]))
    )
    return df.loc[mask].copy()


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def coluna_existe(df: pd.DataFrame, nome: str) -> bool:
    return nome in df.columns


def melt_metas(df_metas_raw: pd.DataFrame) -> pd.DataFrame:
    df = normalizar_colunas(df_metas_raw)
    df["_row_id"] = range(len(df))
    
    id_vars = [c for c in df.columns if str(c).lower() in ["empreendimento", "região", "regiao", "obra", "coordenador"]]
    id_vars_merge = id_vars + ["_row_id"]
    
    for c in id_vars:
        df[c] = df[c].fillna("Não Informado").astype(str).str.strip()
        df.loc[df[c] == "", c] = "Não Informado"
        df.loc[df[c].str.lower() == "nan", c] = "Não Informado"

    if "Coordenador" not in df.columns:
        df["Coordenador"] = "Não Informado"
        id_vars_merge.append("Coordenador")

    cols_qtd = [c for c in df.columns if re.match(r'^qtd\s*(1[0-2]|[1-9])$', str(c).lower().strip())]
    cols_vgv = [c for c in df.columns if re.match(r'^vgv\s*(1[0-2]|[1-9])$', str(c).lower().strip())]

    if not [c for c in df.columns if str(c).lower() in ["empreendimento", "obra"]] or not cols_qtd:
        return pd.DataFrame(columns=["Empreendimento", "Região", "Coordenador", "Mes_Num", "Meta_Qtd", "Meta_VGV"])

    df_qtd = df.melt(id_vars=id_vars_merge, value_vars=cols_qtd, var_name="Mes_Str", value_name="Meta_Qtd")
    df_qtd["Mes_Num"] = df_qtd["Mes_Str"].str.extract(r'(\d+)')[0].astype(int)
    df_qtd.drop(columns=["Mes_Str"], inplace=True)
    df_qtd["Meta_Qtd"] = pd.to_numeric(df_qtd["Meta_Qtd"], errors="coerce").fillna(0)

    if cols_vgv:
        df_vgv = df.melt(id_vars=id_vars_merge, value_vars=cols_vgv, var_name="Mes_Str", value_name="Meta_VGV")
        df_vgv["Mes_Num"] = df_vgv["Mes_Str"].str.extract(r'(\d+)')[0].astype(int)
        df_vgv.drop(columns=["Mes_Str"], inplace=True)
        df_vgv["Meta_VGV"] = df_vgv["Meta_VGV"].apply(parse_valor_br)
        out = pd.merge(df_qtd, df_vgv, on=id_vars_merge + ["Mes_Num"], how="outer").fillna(0)
    else:
        out = df_qtd.copy()
        out["Meta_VGV"] = 0.0

    out.drop(columns=["_row_id"], inplace=True, errors="ignore")

    for c in out.columns:
        if str(c).lower() in ["empreendimento", "obra"]:
            out.rename(columns={c: "Empreendimento"}, inplace=True)
        elif str(c).lower() in ["região", "regiao"]:
            out.rename(columns={c: "Região"}, inplace=True)
            
    return out


def fmt_num(v: Any) -> str:
    """Um decimal se houver parte fracionária; inteiro caso contrário."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    if abs(f % 1) < 0.05:
        return str(int(round(f)))
    return f"{f:.1f}".replace(".", ",")


def fmt_br_milhoes(v: float) -> str:
    if v >= 1e6:
        return f"R$ {fmt_num(v / 1e6)} mi"
    if v >= 1e3:
        return f"R$ {fmt_num(v / 1e3)} mil"
    if abs(float(v) % 1) < 0.05:
        return f"R$ {int(round(v)):,}".replace(",", ".")
    s = fmt_num(v)
    return f"R$ {s}"


def fmt_qtd(v: float) -> str:
    """Retorna int se .0, ou 1 decimal se fracionado."""
    return fmt_num(v)


def fmt_pct_valor(v: Any) -> str:
    """Percentual com mesma regra de decimais."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{fmt_num(float(v))}%"


def _coluna_tabela_texto(nome: str) -> bool:
    n = str(nome or "").strip()
    if n in (
        "Empreendimento", "Coordenador", "HabiteSe", "Indicador",
        "Ultimo_Periodo", "Preço médio venda MTD", "Preço médio tabela",
    ):
        return True
    nl = n.lower()
    return nl.startswith("habite") or "periodo" in nl


def _rotulo_coluna_tabela(nome: str) -> str:
    chave = str(nome or "").strip()
    if chave in ROTULOS_COLUNAS_TABELA:
        return ROTULOS_COLUNAS_TABELA[chave]
    if chave.startswith("VSO_") and chave.endswith("d"):
        dias = chave.replace("VSO_", "").replace("d", "")
        return f"Velocidade de sell-out ({dias} dias)"
    if chave.startswith("PS_VGV_"):
        sufixo = chave.replace("PS_VGV_", "").replace("_", " ").strip()
        return f"Pro soluto sobre VGV — {sufixo} (%)"
    if chave.startswith("Sinais_VGV_"):
        sufixo = chave.replace("Sinais_VGV_", "").replace("_", " ").strip()
        return f"Sinais sobre VGV — {sufixo} (%)"
    return chave.replace("_", " ")


def _nomes_colunas_unicos(nomes: List[str]) -> List[str]:
    """Evita colunas duplicadas após renomear (quebra Styler do pandas/Streamlit)."""
    vistos: Dict[str, int] = {}
    out: List[str] = []
    for nome in nomes:
        base = str(nome or "").strip() or "Coluna"
        n = vistos.get(base, 0)
        if n == 0:
            out.append(base)
        else:
            out.append(f"{base} ({n + 1})")
        vistos[base] = n + 1
    return out


def preparar_df_tabela_exibicao(df: pd.DataFrame) -> pd.DataFrame:
    """Renomeia estatísticas por extenso e formata números (0 ou 1 decimal)."""
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    out = df.copy()
    novos_nomes = _nomes_colunas_unicos([_rotulo_coluna_tabela(c) for c in out.columns])
    out.columns = novos_nomes
    for col_orig, col_novo in zip(df.columns, novos_nomes):
        if _coluna_tabela_texto(col_orig):
            out[col_novo] = df[col_orig].astype(str).replace({"nan": "—", "None": "—"})
            continue
        serie = df[col_orig]
        if pd.api.types.is_numeric_dtype(serie) or serie.dtype == object:
            out[col_novo] = serie.map(fmt_num)
    return out


def _ordenar_empreendimento_primeiro(df: pd.DataFrame) -> pd.DataFrame:
    """Coloca Empreendimento na primeira coluna (facilita pin ao rolar)."""
    if df is None or df.empty or "Empreendimento" not in df.columns:
        return df
    cols = ["Empreendimento"] + [c for c in df.columns if c != "Empreendimento"]
    return df[cols]


def _config_colunas_tabela(colunas) -> Dict[str, Any]:
    """Configura colunas largas; Empreendimento fica fixa à esquerda."""
    cfg: Dict[str, Any] = {}
    for c in colunas:
        pin = c == "Empreendimento"
        cfg[c] = st.column_config.TextColumn(
            c,
            width="large",
            pinned=True if pin else None,
        )
    return cfg


def _exibir_dataframe_preparada(
    disp: pd.DataFrame,
    *,
    styler: Optional[Any] = None,
    ocultar_indice: bool = True,
) -> None:
    cfg = _config_colunas_tabela(disp.columns)
    try:
        st.dataframe(
            styler if styler is not None else disp,
            use_container_width=True,
            hide_index=ocultar_indice,
            column_config=cfg,
        )
    except (KeyError, ValueError, TypeError):
        st.dataframe(
            disp,
            use_container_width=True,
            hide_index=ocultar_indice,
            column_config=cfg,
        )


def _styler_desenquadramento(disp: pd.DataFrame, col_display: str) -> Optional[Any]:
    """Destaca desenquadramento > 50%; tolerante a colunas duplicadas / Streamlit."""
    if disp is None or disp.empty or col_display not in disp.columns:
        return None

    def _style_desenq(val: Any) -> str:
        try:
            s = str(val).replace(",", ".")
            if float(s) > 50:
                return "color: #cb0935; font-weight: bold"
        except (TypeError, ValueError):
            pass
        return ""

    try:
        base = disp.loc[:, ~disp.columns.duplicated()].copy()
        if col_display not in base.columns:
            return None
        return base.style.map(_style_desenq, subset=[col_display])
    except (KeyError, ValueError, TypeError):
        return None


def exibir_tabela(
    df: pd.DataFrame,
    *,
    styler: Optional[Any] = None,
    ocultar_indice: bool = True,
) -> None:
    """Exibe tabela com rótulos por extenso, números formatados e colunas largas."""
    if df is None or df.empty:
        return
    disp = _ordenar_empreendimento_primeiro(preparar_df_tabela_exibicao(df))
    _exibir_dataframe_preparada(disp, styler=styler, ocultar_indice=ocultar_indice)


ROTULOS_COLUNAS_TABELA: Dict[str, str] = {
    "Empreendimento": "Empreendimento",
    "Coordenador": "Coordenador",
    "Unidades_Estoque": "Unidades em estoque",
    "Diff_Avaliacao": "Diferença de avaliação",
    "VGV_Estoque": "VGV em estoque",
    "Pct_Garden": "Percentual garden (%)",
    "HabiteSe": "Habite-se",
    "Vendas_Mes_Nec": "Vendas mês necessárias",
    "Desenquadramento_Pct": "Desenquadramento (%)",
    "Ato_Necessario": "Ato necessário",
    "Pct_Tabela_Direta": "Percentual tabela direta (%)",
    "Pct_Tabela_Investidor": "Percentual tabela investidor (%)",
    "Vendas_Realizadas": "Vendas realizadas",
    "Vendas_Futuras": "Vendas futuras",
    "Vendas_Comunicadas": "Vendas comunicadas",
    "VCX": "Volta ao caixa (quantidade)",
    "Pct_Pro_Soluto": "Percentual pro soluto (%)",
    "Pct_Fluxo_Escalonado": "Percentual fluxo escalonado (%)",
    "Renda_Media_90d": "Renda média (90 dias)",
    "Pastas_Aprovadas": "Pastas aprovadas",
    "Pastas_PC_Suficiente": "Pastas com poder de compra suficiente",
    "Ineficiencia_Qtd": "Ineficiência (quantidade)",
    "Ineficiencia_Pct": "Ineficiência (%)",
    "Meta_Dia": "Meta do dia",
    "Meta_Semana": "Meta da semana",
    "Meta_Mes": "Meta do mês",
    "Pct_Meta_Dia": "Percentual meta do dia (%)",
    "Pct_Meta_Semana": "Percentual meta da semana (%)",
    "Pct_Meta_Mes": "Percentual meta do mês (%)",
    "Agendamentos": "Agendamentos",
    "Visitas": "Visitas",
    "Pastas": "Pastas",
    "Pastas_Aprov": "Pastas aprovadas (funil)",
    "Conv_Ag_Vis": "Conversão agendamento → visita (%)",
    "Conv_Vis_Pas": "Conversão visita → pasta (%)",
    "Conv_Pas_Ap": "Conversão pasta → aprovada (%)",
    "Conv_Ap_Ven": "Conversão aprovada → venda (%)",
    "Preco_Medio_Venda": "Preço médio de venda",
    "Preco_Medio_Tabela": "Preço médio de tabela",
    "Meta_Ano": "Meta do ano",
    "Vendido_Ano": "Vendido no ano",
    "Unidades_Disponiveis": "Unidades disponíveis",
    "Unidades_Liberadas": "Unidades liberadas",
    "Unidades_Total_SF": "Unidades totais (Salesforce)",
    "Unidades_Total": "Unidades totais",
    "Pct_Disp_Liberadas": "Percentual disponível sobre liberadas (%)",
    "Pct_Liberadas_Total": "Percentual liberadas sobre total (%)",
    "Sinal_Sobre_Renda_Pct": "Sinal sobre renda (%)",
    "Dias_Agend_Visita": "Dias médios (agendamento → visita)",
    "N_Agend_Visita": "Quantidade (agendamento → visita)",
    "Dias_Visita_Pasta": "Dias médios (visita → pasta)",
    "N_Visita_Pasta": "Quantidade (visita → pasta)",
    "Dias_Pasta_Aprov": "Dias médios (pasta → aprovada)",
    "N_Pasta_Aprov": "Quantidade (pasta → aprovada)",
    "Dias_Aprov_Venda": "Dias médios (aprovada → venda)",
    "N_Aprov_Venda": "Quantidade (aprovada → venda)",
    "Vendas": "Vendas",
    "Hipereficiencia_Qtd": "Hipereficiência (quantidade)",
    "Hipereficiencia_Pct": "Hipereficiência (%)",
    "Indicador": "Indicador",
    "Ultimo_Periodo": "Último período",
    "QTD_MTD_Parcial": "Quantidade MTD parcial",
    "Preço médio venda MTD": "Preço médio venda MTD",
    "Preço médio tabela": "Preço médio tabela",
    "Gap %": "Gap percentual (%)",
    "Estoque_Un": "Unidades em estoque",
    "m2_Total": "Metros quadrados totais",
    "Realizado_Vendas": "Vendas realizadas",
    "Meta_Vendas_Dia": "Meta de vendas do dia",
    "Meta_Vendas_Mes": "Meta de vendas do mês",
    "Pct_Meta_Dia": "Percentual meta do dia (%)",
    "Pct_Meta_Mes": "Percentual meta do mês (%)",
    "Ineficiencia": "Ineficiência (quantidade)",
    "Pct_Venda_Futura": "Percentual venda futura (%)",
    "Meta_Vendas": "Meta de vendas",
    "Pct_Meta": "Percentual da meta (%)",
    "Unidades": "Unidades",
    "Share_Pct": "Participação no estoque (%)",
    "PS_VGV_Media_12m": "Pro soluto sobre VGV — média 12 meses (%)",
    "Sinais_VGV_Media_12m": "Sinais sobre VGV — média 12 meses (%)",
    "PS_VGV_MTD_vs_12m": "Pro soluto MTD versus média 12 meses (p.p.)",
    "Sinais_VGV_MTD_vs_12m": "Sinais MTD versus média 12 meses (p.p.)",
    "MTD_Vendas": "Quantidade mês — Vendas (MTD parcial)",
    "Media_Hist_Vendas": "Quantidade média histórica — Vendas (MTD parcial)",
    "MTD_Agendamentos": "Quantidade mês — Agendamentos (MTD parcial)",
    "Media_Hist_Agendamentos": "Quantidade média histórica — Agendamentos (MTD parcial)",
    "MTD_Visitas": "Quantidade mês — Visitas (MTD parcial)",
    "Media_Hist_Visitas": "Quantidade média histórica — Visitas (MTD parcial)",
    "MTD_Pastas": "Quantidade mês — Pastas (MTD parcial)",
    "Media_Hist_Pastas": "Quantidade média histórica — Pastas (MTD parcial)",
    "MTD_Pastas_Aprov": "Quantidade mês — Pastas aprovadas (MTD parcial)",
    "Media_Hist_Pastas_Aprov": "Quantidade média histórica — Pastas aprovadas (MTD parcial)",
    "Preco_Medio_Venda_MTD": "Preço médio de venda MTD",
    "Gap_Preco_Pct": "Gap preço venda vs tabela (%)",
    "Share_Estoque_Pct": "Participação no estoque (%)",
    "Share_Vendas_Pct": "Participação nas vendas (%)",
    "m2_Total": "Metros quadrados totais",
}


def fmt_funil_valor(v: float) -> str:
    """Valor exibido no funil: sempre inteiro (já arredondado para cima)."""
    return str(int(math.ceil(max(0.0, float(v)))))


def ceil_funil_totais(totais: Dict[str, float]) -> Dict[str, float]:
    """Arredonda para cima todos os volumes do funil."""
    return {e: float(math.ceil(max(0.0, float((totais or {}).get(e, 0.0))))) for e in FUNIL_ETAPAS}


FUNIL_FONTE_TAMANHO = 16
FUNIL_TEXTO_BRANCO = "#ffffff"


def _layout_plotly_preto(fig: go.Figure, titulo: str = "", altura: int = 380, margin_r: int = 120) -> go.Figure:
    """Layout padrão: todo texto preto (Inter)."""
    layout: Dict[str, Any] = dict(
        margin=dict(l=10, r=margin_r, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=altura,
        font=dict(family="Inter", color=COR_TEXTO_PRETO, size=FUNIL_FONTE_TAMANHO),
    )
    if titulo:
        layout["title"] = dict(
            text=titulo,
            font=dict(family="Inter", color=COR_TEXTO_PRETO, size=FUNIL_FONTE_TAMANHO),
        )
    fig.update_layout(**layout)
    fig.update_xaxes(tickfont=dict(color=COR_TEXTO_PRETO, family="Inter", size=FUNIL_FONTE_TAMANHO))
    fig.update_yaxes(tickfont=dict(color=COR_TEXTO_PRETO, family="Inter", size=FUNIL_FONTE_TAMANHO))
    return fig


def _fonte_funil_plotly() -> Dict[str, Any]:
    """Fonte uniforme preta para rótulos do funil (dentro e fora do bloco)."""
    return dict(size=FUNIL_FONTE_TAMANHO, color=COR_TEXTO_PRETO, family="Inter")


def _criar_fig_funil(
    labels: List[str],
    valores: List[float],
    titulo: str = "",
    cores: Optional[List[str]] = None,
    altura: int = 380,
) -> go.Figure:
    """
    Funil Plotly — rótulos sempre fora do bloco, fonte uniforme preta.
    Volumes sempre arredondados para cima (ceil).
    """
    vals = [float(math.ceil(max(0.0, float(v)))) for v in valores]
    textos = [fmt_funil_valor(v) for v in vals]
    fonte = _fonte_funil_plotly()
    fig = go.Figure(go.Funnel(
        y=labels,
        x=vals,
        text=textos,
        textinfo="text",
        textposition="outside",
        insidetextfont=fonte,
        outsidetextfont=fonte,
        marker={"color": cores or FUNIL_CORES_NIVEIS},
        connector={"fillcolor": "rgba(4, 66, 143, 0.15)"},
    ))
    return _layout_plotly_preto(fig, titulo=titulo, altura=altura, margin_r=140)


DIAS_SEMANA_PT = {
    0: "segunda",
    1: "terça",
    2: "quarta",
    3: "quinta",
    4: "sexta",
    5: "sábado",
    6: "domingo",
}
MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}


def janela_treino_meses_exatos(hoje: Optional[date] = None) -> Tuple[date, date]:
    """
    36 meses-calendário fechados, excluindo o mês atual.
    Ex.: julho/2026 → 01/07/2023 a 30/06/2026.
    """
    hoje = hoje or date.today()
    inicio = date(hoje.year - FUNIL_ANOS_TREINO, hoje.month, 1)
    if hoje.month == 1:
        fim = date(hoje.year - 1, 12, 31)
    else:
        ano_fim, mes_fim = hoje.year, hoje.month - 1
        fim = date(ano_fim, mes_fim, calendar.monthrange(ano_fim, mes_fim)[1])
    return inicio, fim


def janela_treino_52_semanas(hoje: Optional[date] = None) -> Tuple[date, date]:
    """Compatibilidade: encaminha para janela de meses exatos."""
    return janela_treino_meses_exatos(hoje)


_TZ_BR = "America/Sao_Paulo"


def parse_data_serie(serie: pd.Series) -> pd.Series:
    """
    Converte datas Salesforce/relatórios para datetime64[ns] naive.

    Datas ISO do SF (`2025-07-01T12:14:25.000+0000` ou `2025-07-01`) NÃO podem
    usar dayfirst=True — o pandas troca 2025-07-01 por 2025-01-07 e invalida
    dias > 12, o que derruba agendamentos/visitas do mês atual.
    """
    if serie is None:
        return pd.Series(dtype="datetime64[ns]")
    raw = serie
    as_str = raw.map(
        lambda x: ""
        if x is None or (isinstance(x, float) and pd.isna(x))
        else str(x).strip()
    )
    out = pd.Series(pd.NaT, index=raw.index, dtype="datetime64[ns]")

    mask_iso = as_str.str.match(r"^\d{4}-\d{2}-\d{2}", na=False)
    if mask_iso.any():
        has_time = mask_iso & as_str.str.contains("T", na=False)
        date_only = mask_iso & ~has_time
        if has_time.any():
            ts = pd.to_datetime(raw.loc[has_time], errors="coerce", utc=True)
            out.loc[has_time] = ts.dt.tz_convert(_TZ_BR).dt.tz_localize(None)
        if date_only.any():
            out.loc[date_only] = pd.to_datetime(
                as_str.loc[date_only], format="%Y-%m-%d", errors="coerce"
            )

    vazios = {"", "nan", "none", "nat", "null", "na", "n/a", "-"}
    mask_rest = out.isna() & ~as_str.str.lower().isin(vazios)
    if mask_rest.any():
        out.loc[mask_rest] = pd.to_datetime(
            raw.loc[mask_rest], dayfirst=True, errors="coerce"
        )

    if out.isna().all():
        nums = pd.to_numeric(raw, errors="coerce")
        if nums.notna().any():
            med = float(nums.dropna().median())
            if med > 1e12:
                out = pd.to_datetime(nums, unit="ms", errors="coerce")
            elif med > 1e9:
                out = pd.to_datetime(nums, unit="s", errors="coerce")
            else:
                out = pd.to_datetime(
                    nums, unit="D", origin="1899-12-30", errors="coerce"
                )
    return out


def _as_date_funil(val: Any) -> Optional[date]:
    """Normaliza valor para date (evita falha de lookup Timestamp vs date)."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, pd.Timestamp):
        return val.date()
    try:
        dt = parse_data_serie(pd.Series([val])).iloc[0]
        if pd.isna(dt):
            return None
        return pd.Timestamp(dt).date()
    except Exception:
        return None


def _indice_por_data_cal(cal: pd.DataFrame) -> Dict[date, int]:
    """Mapa data → posição no calendário (índice resetado)."""
    out: Dict[date, int] = {}
    for i, r in cal.reset_index(drop=True).iterrows():
        d = _as_date_funil(r["data"])
        if d is not None:
            out[d] = int(i)
    return out


def _media_etapa_dia_semana(treino: pd.DataFrame, etapa: str, d: date) -> float:
    if treino.empty or etapa not in treino.columns:
        return 0.0
    ds = DIAS_SEMANA_PT[d.weekday()]
    sub = treino[treino["dia_semana"] == ds]
    if not sub.empty:
        return max(float(sub[etapa].mean()), 0.0)
    return max(float(treino[etapa].mean()), 0.0)


def _garantir_previsoes_futuras_funil(
    cal_reg: pd.DataFrame,
    cal_med: pd.DataFrame,
    idx_por_data: Dict[date, int],
    dias_futuros: List[date],
    treino: pd.DataFrame,
    medias: Dict[str, Any],
    coefs: Dict[str, np.ndarray],
    incluir_mes: bool,
    lags: Tuple[int, ...],
) -> None:
    """Evita projeção zerada no restante do mês (fallback: médias / histórico)."""
    for d in dias_futuros:
        i = idx_por_data.get(d)
        if i is None:
            continue
        for etapa in FUNIL_ETAPAS:
            v_reg = float(cal_reg.at[i, etapa])
            if v_reg < 1e-6:
                row_reg = cal_reg.loc[[i]]
                v_reg = _prever_linha_reg_funil(
                    coefs[etapa], row_reg, incluir_mes, lags, etapa, usar_cal_lags=True
                )
                if v_reg < 1e-6:
                    v_med = _prever_linha_medias_funil(
                        cal_med.iloc[i], d, etapa, medias, incluir_mes
                    )
                    if v_med > 1e-6:
                        v_reg = v_med
                    elif treino is not None and not treino.empty:
                        v_reg = _media_etapa_dia_semana(treino, etapa, d)
                cal_reg.at[i, etapa] = v_reg
            v_med = float(cal_med.at[i, etapa])
            if v_med < 1e-6:
                v_med = _prever_linha_medias_funil(
                    cal_med.iloc[i], d, etapa, medias, incluir_mes
                )
                if v_med < 1e-6 and treino is not None and not treino.empty:
                    v_med = _media_etapa_dia_semana(treino, etapa, d)
                cal_med.at[i, etapa] = v_med


def serie_diaria_contratos(
    df_vendas: pd.DataFrame,
    col_contrato: str,
    col_qtd: str = "_qtd_venda",
    col_vgv: str = "_vgv_venda",
) -> pd.DataFrame:
    """Agrega vendas comerciais por data de 'Contrato gerado em'."""
    base = df_vendas.copy()
    base["_dt_contrato"] = parse_data_serie(base[col_contrato])
    base = base.dropna(subset=["_dt_contrato"])
    if base.empty:
        return pd.DataFrame(columns=["data", "qtd", "vgv"])

    base["_data"] = base["_dt_contrato"].dt.normalize()
    qtd_col = col_qtd if col_qtd in base.columns else None
    vgv_col = col_vgv if col_vgv in base.columns else None

    if qtd_col:
        base["_q"] = pd.to_numeric(base[qtd_col], errors="coerce").fillna(0.0)
    else:
        base["_q"] = 1.0
    if vgv_col:
        base["_v"] = pd.to_numeric(base[vgv_col], errors="coerce").fillna(0.0)
    else:
        base["_v"] = 0.0

    agg = (
        base.groupby("_data", as_index=False)
        .agg(qtd=("_q", "sum"), vgv=("_v", "sum"))
        .rename(columns={"_data": "data"})
    )
    agg["data"] = pd.to_datetime(agg["data"]).dt.date
    return agg


def calendario_diario(inicio: date, fim: date, serie: pd.DataFrame) -> pd.DataFrame:
    """Calendário completo com zeros nos dias sem venda (necessário para a regressão)."""
    idx = pd.date_range(inicio, fim, freq="D")
    cal = pd.DataFrame({"data": [d.date() for d in idx]})
    mapa = {r["data"]: (float(r["qtd"]), float(r["vgv"])) for _, r in serie.iterrows()}
    cal["qtd"] = cal["data"].map(lambda d: mapa.get(d, (0.0, 0.0))[0])
    cal["vgv"] = cal["data"].map(lambda d: mapa.get(d, (0.0, 0.0))[1])
    cal["dia_mes"] = cal["data"].map(lambda d: d.day)
    cal["dia_semana"] = cal["data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()])
    cal["mes"] = cal["data"].map(lambda d: MESES_PT[d.month])
    return cal


def _matriz_explicativas(df: pd.DataFrame, incluir_mes: bool = True) -> np.ndarray:
    """One-hot (numpy) de dia do mês + dia da semana (+ mês opcional) + intercepto."""
    n = len(df)
    n_cols = 31 + 7 + (12 if incluir_mes else 0) + 1
    X = np.zeros((n, n_cols), dtype=float)
    X[:, -1] = 1.0  # intercepto

    dias_semana_idx = {nome: i for i, nome in DIAS_SEMANA_PT.items()}
    meses_idx = {nome: i for i, nome in MESES_PT.items()}

    for i, row in enumerate(df.itertuples(index=False)):
        dia = int(row.dia_mes)
        if 1 <= dia <= 31:
            X[i, dia - 1] = 1.0
        ds = dias_semana_idx.get(str(row.dia_semana), None)
        if ds is not None:
            X[i, 31 + ds] = 1.0
        if incluir_mes:
            ms = meses_idx.get(str(row.mes), None)
            if ms is not None:
                X[i, 31 + 7 + (ms - 1)] = 1.0
    return X


def treinar_regressao_vendas_diarias(
    treino: pd.DataFrame,
    incluir_mes: bool = True,
) -> np.ndarray:
    """OLS via numpy.linalg.lstsq. Retorna vetor de coeficientes."""
    X = _matriz_explicativas(treino, incluir_mes=incluir_mes)
    y = treino["qtd"].astype(float).values
    coef, *_ = np.linalg.lstsq(X, y, rcond=None)
    return coef


def prever_qtd_dias(
    coef: np.ndarray,
    datas: List[date],
    incluir_mes: bool = True,
) -> np.ndarray:
    if not datas:
        return np.array([])
    df = pd.DataFrame({"data": datas})
    df["dia_mes"] = df["data"].map(lambda d: d.day)
    df["dia_semana"] = df["data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()])
    df["mes"] = df["data"].map(lambda d: MESES_PT[d.month])
    X = _matriz_explicativas(df, incluir_mes=incluir_mes)
    pred = X @ coef
    return np.maximum(pred, 0.0)


def _r2_treino(
    treino: pd.DataFrame,
    coef: np.ndarray,
    incluir_mes: bool = True,
) -> float:
    X = _matriz_explicativas(treino, incluir_mes=incluir_mes)
    y = treino["qtd"].astype(float).values
    y_hat = X @ coef
    ss_res = float(np.sum((y - y_hat) ** 2))
    ss_tot = float(np.sum((y - y.mean()) ** 2))
    if ss_tot <= 0:
        return 0.0
    return 1.0 - (ss_res / ss_tot)


def calcular_medias_sazonais(
    treino: pd.DataFrame,
    incluir_mes: bool = True,
) -> Dict[str, Any]:
    """Médias históricas por dia da semana, dia do mês e (opcionalmente) mês."""
    mu = float(treino["qtd"].mean()) if len(treino) else 0.0
    if mu <= 0:
        mu = 1e-9
    out: Dict[str, Any] = {
        "mu": mu,
        "incluir_mes": incluir_mes,
        "media_dia_semana": {
            k: float(v) for k, v in treino.groupby("dia_semana")["qtd"].mean().items()
        },
        "media_dia_mes": {
            int(k): float(v) for k, v in treino.groupby("dia_mes")["qtd"].mean().items()
        },
    }
    if incluir_mes:
        out["media_mes"] = {
            k: float(v) for k, v in treino.groupby("mes")["qtd"].mean().items()
        }
    return out


def prever_qtd_medias(
    datas: List[date],
    medias: Dict[str, Any],
    incluir_mes: bool = True,
) -> np.ndarray:
    """
    Combinação multiplicativa das médias.
    Com mês:   pred = m_ds × m_dm × m_mes / μ²
    Sem mês:   pred = m_ds × m_dm / μ
    """
    if not datas:
        return np.array([])
    mu = float(medias["mu"])
    m_ds = medias["media_dia_semana"]
    m_dm = medias["media_dia_mes"]
    m_mes = medias.get("media_mes") or {}
    usar_mes = incluir_mes and bool(m_mes)
    out: List[float] = []
    for d in datas:
        a = float(m_ds.get(DIAS_SEMANA_PT[d.weekday()], mu))
        b = float(m_dm.get(d.day, mu))
        if usar_mes:
            c = float(m_mes.get(MESES_PT[d.month], mu))
            out.append(max((a * b * c) / (mu * mu), 0.0))
        else:
            out.append(max((a * b) / mu, 0.0))
    return np.asarray(out, dtype=float)


def _matriz_explicativas_relativa(df: pd.DataFrame) -> np.ndarray:
    """
    One-hot com categorias de referência omitidas:
      dia do mês → dia 1
      dia da semana → segunda
      mês → janeiro
    + intercepto (= nível esperado em segunda + dia 1 + janeiro).
    """
    n = len(df)
    # 30 dias (2–31) + 6 dias semana (terça–domingo) + 11 meses (fev–dez) + intercepto
    X = np.zeros((n, 30 + 6 + 11 + 1), dtype=float)
    X[:, -1] = 1.0

    dias_semana_idx = {nome: i for i, nome in DIAS_SEMANA_PT.items()}  # 0=segunda
    meses_idx = {nome: i for i, nome in MESES_PT.items()}  # 1=janeiro

    for i, row in enumerate(df.itertuples(index=False)):
        dia = int(row.dia_mes)
        if 2 <= dia <= 31:
            X[i, dia - 2] = 1.0  # colunas 0..29
        ds = dias_semana_idx.get(str(row.dia_semana), None)
        if ds is not None and ds >= 1:  # terça=1 .. domingo=6
            X[i, 30 + (ds - 1)] = 1.0
        ms = meses_idx.get(str(row.mes), None)
        if ms is not None and ms >= 2:  # fevereiro=2 .. dezembro=12
            X[i, 30 + 6 + (ms - 2)] = 1.0
    return X


def estimar_efeitos_sazonais(treino: pd.DataFrame) -> Optional[Dict[str, Any]]:
    """
    OLS com referência em segunda-feira, dia 1 e janeiro.
    Retorna efeitos aditivos (qtd) e índices relativos (baseline = 1.0).
    """
    if treino.empty or float(treino["qtd"].sum()) <= 0 or len(treino) < 30:
        return None

    X = _matriz_explicativas_relativa(treino)
    y = treino["qtd"].astype(float).values
    coef, *_ = np.linalg.lstsq(X, y, rcond=None)
    intercepto = float(coef[-1])
    base = intercepto if abs(intercepto) > 1e-9 else 1.0

    # dia do mês
    e_dm = [0.0]  # dia 1
    for d in range(2, 32):
        e_dm.append(float(coef[d - 2]))
    df_dm = pd.DataFrame({
        "categoria": [str(d) for d in range(1, 32)],
        "efeito": e_dm,
        "indice": [(base + e) / base for e in e_dm],
    })

    # dia da semana (ordem: segunda → domingo)
    nomes_ds = [DIAS_SEMANA_PT[i] for i in range(7)]
    e_ds = [0.0]  # segunda
    for i in range(1, 7):
        e_ds.append(float(coef[30 + (i - 1)]))
    df_ds = pd.DataFrame({
        "categoria": nomes_ds,
        "efeito": e_ds,
        "indice": [(base + e) / base for e in e_ds],
    })

    # mês (ordem: janeiro → dezembro)
    nomes_mes = [MESES_PT[i] for i in range(1, 13)]
    e_mes = [0.0]  # janeiro
    for m in range(2, 13):
        e_mes.append(float(coef[30 + 6 + (m - 2)]))
    df_mes = pd.DataFrame({
        "categoria": nomes_mes,
        "efeito": e_mes,
        "indice": [(base + e) / base for e in e_mes],
    })

    y_hat = X @ coef
    ss_res = float(np.sum((y - y_hat) ** 2))
    ss_tot = float(np.sum((y - y.mean()) ** 2))
    r2 = (1.0 - ss_res / ss_tot) if ss_tot > 0 else 0.0

    return {
        "intercepto": intercepto,
        "r2": r2,
        "dia_mes": df_dm,
        "dia_semana": df_ds,
        "mes": df_mes,
    }


def _distribuir_gap_por_pesos(
    gap: float,
    pesos: np.ndarray,
    dias: List[date],
    qtd_mtd: float,
    dia_hoje: int,
    arredondar_cima: bool = False,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    ritmo_diario: List[Dict[str, Any]] = []
    ritmo_acum: List[Dict[str, Any]] = [{"dia": dia_hoje, "acum": qtd_mtd}]
    if not dias:
        return pd.DataFrame(ritmo_diario), pd.DataFrame(ritmo_acum)
    w = np.maximum(np.asarray(pesos, dtype=float), 0.0)
    soma = float(w.sum())
    if soma <= 1e-9:
        w = np.ones(len(dias), dtype=float)
        soma = float(len(dias))
    nec = gap * (w / soma)
    if arredondar_cima and gap > 0:
        nec = np.ceil(nec)
    running = qtd_mtd
    for i, d in enumerate(dias):
        v = float(nec[i])
        ritmo_diario.append({"dia": d.day, "qtd": v})
        running += v
        ritmo_acum.append({"dia": d.day, "acum": running})
    return pd.DataFrame(ritmo_diario), pd.DataFrame(ritmo_acum)


JANELAS_FIM_MES = (15, 10, 5)


def calcular_intensidade_fim_mes(treino: pd.DataFrame, janela: int = 15) -> float:
    """
    Razão empírica: média dos últimos `janela` dias do mês / média do restante.
    """
    if treino.empty:
        return 1.4
    limiar = max(1, 31 - janela + 1)
    early = treino[treino["dia_mes"] < limiar]["qtd"]
    late = treino[treino["dia_mes"] >= limiar]["qtd"]
    m_early = float(early.mean()) if len(early) else 0.0
    m_late = float(late.mean()) if len(late) else 0.0
    if m_early <= 1e-9:
        return 1.4
    ratio = m_late / m_early
    return float(np.clip(ratio, 1.10, 2.8))


def calcular_intensidades_fim_mes(
    treino: pd.DataFrame,
    janelas: Tuple[int, ...] = JANELAS_FIM_MES,
) -> Dict[int, float]:
    return {int(j): calcular_intensidade_fim_mes(treino, janela=int(j)) for j in janelas}


def fatores_sazonalidade_fim_mes(
    dias: List[date],
    ultimo_dia: int,
    intensidades: Dict[int, float],
    janelas: Tuple[int, ...] = JANELAS_FIM_MES,
) -> np.ndarray:
    """
    Combina sazonalidades de 15, 10 e 5 dias (produto de rampas).
    Quanto mais perto do fim do mês, maior o fator.
    """
    if not dias:
        return np.array([])
    out: List[float] = []
    for d in dias:
        f = 1.0
        for janela in janelas:
            intens = float(intensidades.get(int(janela), 1.0))
            inicio_boost = max(1, ultimo_dia - int(janela) + 1)
            if d.day >= inicio_boost:
                span = max(1, ultimo_dia - inicio_boost)
                t = (d.day - inicio_boost) / span  # 0 → 1
                f *= 1.0 + (intens - 1.0) * t
        out.append(f)
    return np.asarray(out, dtype=float)


def aplicar_sazonalidade_fim_mes(
    pesos: np.ndarray,
    dias: List[date],
    ultimo_dia: int,
    intensidades: Dict[int, float],
    janelas: Tuple[int, ...] = JANELAS_FIM_MES,
) -> np.ndarray:
    """Multiplica pesos pela sazonalidade combinada (15/10/5 dias)."""
    if len(pesos) == 0:
        return pesos
    fat = fatores_sazonalidade_fim_mes(dias, ultimo_dia, intensidades, janelas=janelas)
    return np.maximum(np.asarray(pesos, dtype=float), 0.0) * fat


def _series_atingido_projetado(
    dias_passados: List[date],
    dias_futuros: List[date],
    mapa_real: Dict[Any, float],
    pred_futuro: np.ndarray,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    diaria: List[Dict[str, Any]] = []
    acum: List[Dict[str, Any]] = []
    running = 0.0
    for d in dias_passados:
        q = float(mapa_real.get(d, 0.0))
        diaria.append({"dia": d.day, "qtd": q, "tipo": "Atingida"})
        running += q
        acum.append({"dia": d.day, "acum": running, "tipo": "Atingida"})
    for i, d in enumerate(dias_futuros):
        q = float(pred_futuro[i]) if i < len(pred_futuro) else 0.0
        diaria.append({"dia": d.day, "qtd": q, "tipo": "Projetada"})
        running += q
        acum.append({"dia": d.day, "acum": running, "tipo": "Projetada"})
    return pd.DataFrame(diaria), pd.DataFrame(acum)


def projetar_vendas_mes_atual(
    df_vendas: pd.DataFrame,
    col_contrato: str,
    meta_vgv_mes: float,
    meta_qtd_mes: float = 0.0,
    hoje: Optional[date] = None,
    incluir_mes: bool = True,
) -> Optional[Dict[str, Any]]:
    """
    Projeção do mês corrente com duas lógicas (mesma janela de meses exatos):
      1) Regressão OLS (dia_mês + dia_semana [+ mês])
      2) Médias sazonais combinadas (multiplicativo)
    Com incluir_mes=False, remove o efeito/beta de mês (só dia do mês e dia da semana).
    Treino: do mês atual − 1 ano até o fim do mês anterior
    (ex.: jul/2026 → 01/07/2025 a 30/06/2026).
    """
    hoje = hoje or date.today()
    inicio, fim_treino = janela_treino_meses_exatos(hoje)
    serie = serie_diaria_contratos(df_vendas, col_contrato)
    if serie.empty:
        return None

    treino = calendario_diario(inicio, fim_treino, serie)
    if treino["qtd"].sum() <= 0 or len(treino) < 30:
        return None

    modelo = treinar_regressao_vendas_diarias(treino, incluir_mes=incluir_mes)
    medias = calcular_medias_sazonais(treino, incluir_mes=incluir_mes)

    ano, mes = hoje.year, hoje.month
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    dias_mes = [date(ano, mes, d) for d in range(1, ultimo_dia + 1)]
    dias_passados = [d for d in dias_mes if d <= hoje]
    dias_futuros = [d for d in dias_mes if d > hoje]

    mapa_real = {r["data"]: float(r["qtd"]) for _, r in serie.iterrows()}
    mapa_vgv = {r["data"]: float(r["vgv"]) for _, r in serie.iterrows()}

    qtd_mtd = float(sum(mapa_real.get(d, 0.0) for d in dias_passados))
    vgv_mtd = float(sum(mapa_vgv.get(d, 0.0) for d in dias_passados))

    pred_reg_mes = prever_qtd_dias(modelo, dias_mes, incluir_mes=incluir_mes)
    pred_med_mes = prever_qtd_medias(dias_mes, medias, incluir_mes=incluir_mes)

    # Reforço explícito de fim de mês: sazonalidade em 15, 10 e 5 dias
    intensidades_fim = calcular_intensidades_fim_mes(treino, janelas=JANELAS_FIM_MES)
    pred_reg_mes = aplicar_sazonalidade_fim_mes(pred_reg_mes, dias_mes, ultimo_dia, intensidades_fim)
    pred_med_mes = aplicar_sazonalidade_fim_mes(pred_med_mes, dias_mes, ultimo_dia, intensidades_fim)

    n_passados = len(dias_passados)
    pred_reg = pred_reg_mes[n_passados:] if len(pred_reg_mes) else np.array([])
    pred_med = pred_med_mes[n_passados:] if len(pred_med_mes) else np.array([])

    mu_dia = max(float(treino["qtd"].mean()), 0.0)
    if len(pred_reg) > 0 and float(np.sum(pred_reg)) < 1e-6 and mu_dia > 0:
        pred_reg = np.full(len(pred_reg), mu_dia)
        pred_reg_mes = np.concatenate([pred_reg_mes[:n_passados], pred_reg])
    if len(pred_med) > 0 and float(np.sum(pred_med)) < 1e-6 and mu_dia > 0:
        pred_med = np.full(len(pred_med), mu_dia)
        pred_med_mes = np.concatenate([pred_med_mes[:n_passados], pred_med])

    # Comparativo dia a dia: projetado (modelo) × realizado (mês atual)
    comp_reg: List[Dict[str, Any]] = []
    comp_med: List[Dict[str, Any]] = []
    for i, d in enumerate(dias_mes):
        real = float(mapa_real.get(d, 0.0)) if d <= hoje else None
        p_reg = float(pred_reg_mes[i]) if i < len(pred_reg_mes) else 0.0
        p_med = float(pred_med_mes[i]) if i < len(pred_med_mes) else 0.0
        comp_reg.append({"dia": d.day, "realizado": real, "projetado": p_reg})
        comp_med.append({"dia": d.day, "realizado": real, "projetado": p_med})

    qtd_proj_reg = qtd_mtd + float(pred_reg.sum() if len(pred_reg) else 0.0)
    qtd_proj_med = qtd_mtd + float(pred_med.sum() if len(pred_med) else 0.0)

    # Ticket médio dos 30 dias anteriores à projeção (estoque muda constantemente)
    fim_ticket = hoje - timedelta(days=1)
    inicio_ticket = hoje - timedelta(days=30)
    qtd_30 = float(sum(mapa_real.get(inicio_ticket + timedelta(days=i), 0.0) for i in range(30)))
    vgv_30 = float(sum(mapa_vgv.get(inicio_ticket + timedelta(days=i), 0.0) for i in range(30)))
    ticket_medio = (vgv_30 / qtd_30) if qtd_30 > 0 else 0.0
    if ticket_medio <= 0:
        q_tr = float(treino["qtd"].sum())
        v_tr = float(treino["vgv"].sum())
        ticket_medio = (v_tr / q_tr) if q_tr > 0 else 0.0

    rest_reg = qtd_proj_reg - qtd_mtd
    rest_med = qtd_proj_med - qtd_mtd
    vgv_proj_reg = vgv_mtd + rest_reg * ticket_medio
    vgv_proj_med = vgv_mtd + rest_med * ticket_medio
    pct_vgv_reg = (vgv_proj_reg / meta_vgv_mes * 100.0) if meta_vgv_mes and meta_vgv_mes > 0 else 0.0
    pct_vgv_med = (vgv_proj_med / meta_vgv_mes * 100.0) if meta_vgv_mes and meta_vgv_mes > 0 else 0.0

    meta_qtd = float(meta_qtd_mes or 0.0)
    gap_qtd = max(0.0, meta_qtd - qtd_mtd)

    ritmo_reg_d, ritmo_reg_a = _distribuir_gap_por_pesos(
        gap_qtd, pred_reg, dias_futuros, qtd_mtd, hoje.day, arredondar_cima=True
    )
    ritmo_med_d, ritmo_med_a = _distribuir_gap_por_pesos(
        gap_qtd, pred_med, dias_futuros, qtd_mtd, hoje.day, arredondar_cima=True
    )

    diaria_reg, acum_reg = _series_atingido_projetado(
        dias_passados, dias_futuros, mapa_real, pred_reg
    )
    diaria_med, acum_med = _series_atingido_projetado(
        dias_passados, dias_futuros, mapa_real, pred_med
    )

    intensidade_resumo = float(np.mean(list(intensidades_fim.values()))) if intensidades_fim else 1.0

    return {
        "hoje": hoje,
        "inicio_treino": inicio,
        "fim_treino": fim_treino,
        "incluir_mes": incluir_mes,
        "qtd_mtd": qtd_mtd,
        "vgv_mtd": vgv_mtd,
        "ticket_medio": ticket_medio,
        "inicio_ticket_30d": inicio_ticket,
        "fim_ticket_30d": fim_ticket,
        "ultimo_dia": ultimo_dia,
        "meta_vgv_mes": meta_vgv_mes,
        "meta_qtd_mes": meta_qtd,
        "gap_qtd_meta": gap_qtd,
        "r2_treino": _r2_treino(treino, modelo, incluir_mes=incluir_mes),
        "medias": medias,
        "intensidades_fim_mes": intensidades_fim,
        "intensidade_fim_mes": intensidade_resumo,
        "qtd_projetada_mes": qtd_proj_reg,
        "vgv_projetado": vgv_proj_reg,
        "pct_vgv_meta": pct_vgv_reg,
        "diaria": diaria_reg,
        "acumulado": acum_reg,
        "ritmo_meta_diario": ritmo_reg_d,
        "ritmo_meta_acum": ritmo_reg_a,
        "qtd_projetada_medias": qtd_proj_med,
        "vgv_projetado_medias": vgv_proj_med,
        "pct_vgv_medias": pct_vgv_med,
        "diaria_medias": diaria_med,
        "acumulado_medias": acum_med,
        "ritmo_meta_diario_medias": ritmo_med_d,
        "ritmo_meta_acum_medias": ritmo_med_a,
        "comparativo_diario_reg": pd.DataFrame(comp_reg),
        "comparativo_diario_medias": pd.DataFrame(comp_med),
        "pred_reg_mes": pred_reg_mes,
        "modelo": modelo,
    }


def _plotly_key(*parts: str) -> str:
    """Chave única para st.plotly_chart (evita colisão com/sem efeito de mês)."""
    return "_".join(re.sub(r"[^a-zA-Z0-9_]+", "_", str(p).strip().lower()) for p in parts if p)


def _proj_plot_key(proj: Dict[str, Any], *parts: str) -> str:
    modo = "sem_mes" if not proj.get("incluir_mes", True) else "com_mes"
    return _plotly_key("proj_vendas", modo, *parts)


def _plot_projecao_mes(
    titulo: str,
    caption: str,
    proj: Dict[str, Any],
    diaria: pd.DataFrame,
    acumulado: pd.DataFrame,
    chart_id: str = "reg",
) -> None:
    """Gráfico de projeção: realizado até hoje + projetado a partir de amanhã (sem meta)."""
    st.markdown(f"##### {titulo}")
    if caption:
        st.caption(caption)

    dia_hoje = proj["hoje"].day
    ating_d = diaria[diaria["tipo"] == "Atingida"] if not diaria.empty else diaria
    proj_d = diaria[diaria["tipo"] == "Projetada"] if not diaria.empty else diaria
    ating_a = acumulado[acumulado["tipo"] == "Atingida"] if not acumulado.empty else acumulado
    proj_a = acumulado[acumulado["tipo"] == "Projetada"] if not acumulado.empty else acumulado

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    if not ating_d.empty:
        fig.add_trace(
            go.Bar(
                x=ating_d["dia"], y=ating_d["qtd"], name="Realizado (dia)",
                marker_color=COR_AZUL_ESC, opacity=0.85,
            ),
            secondary_y=False,
        )
    if not proj_d.empty:
        fig.add_trace(
            go.Bar(
                x=proj_d["dia"], y=proj_d["qtd"], name="Projetado (dia)",
                marker_color="rgba(203, 9, 53, 0.45)",
            ),
            secondary_y=False,
        )

    if not ating_a.empty:
        fig.add_trace(
            go.Scatter(
                x=ating_a["dia"], y=ating_a["acum"], mode="lines+markers",
                name="Acumulado realizado",
                line=dict(color=COR_AZUL_ESC, width=3), marker=dict(size=7),
            ),
            secondary_y=True,
        )
    if not proj_a.empty:
        x_proj = [dia_hoje] + proj_a["dia"].tolist()
        y_proj = [float(ating_a["acum"].iloc[-1]) if not ating_a.empty else 0.0] + proj_a["acum"].tolist()
        fig.add_trace(
            go.Scatter(
                x=x_proj, y=y_proj, mode="lines+markers",
                name="Acumulado projetado",
                line=dict(color=COR_VERMELHO, width=3, dash="dash"),
                marker=dict(size=7, color=COR_VERMELHO),
            ),
            secondary_y=True,
        )

    fig.add_vline(
        x=dia_hoje, line_width=1, line_dash="dot", line_color="#64748b",
        annotation_text="Hoje", annotation_position="top",
        annotation_font=dict(color=COR_TEXTO_PRETO, size=11, family="Inter"),
    )
    fig.update_layout(
        barmode="group",
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        hovermode="x unified",
        height=420,
    )
    fig.update_xaxes(
        title_text="Dia do mês",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
        range=[0.5, proj["ultimo_dia"] + 0.5],
    )
    fig.update_yaxes(
        title_text="Qtd. no dia",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        secondary_y=False,
        showgrid=False,
    )
    fig.update_yaxes(
        title_text="Qtd. acumulada",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        secondary_y=True,
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
        key=_proj_plot_key(proj, chart_id, "mes"),
    )


def _plot_meta_diaria_integrada(proj: Dict[str, Any]) -> None:
    """
    Um único gráfico com 3 linhas (somente diário, sem acumulado):
      1) Realizado até hoje
      2) Meta diária (regressão) a partir de amanhã
      3) Meta diária (médias) a partir de amanhã
    """
    ints = proj.get("intensidades_fim_mes") or {}
    st.caption(
        " · ".join(f"Sazonalidade {j}d: {fmt_num(float(ints.get(j, 1)))}x" for j in JANELAS_FIM_MES)
    )

    dia_hoje = proj["hoje"].day
    diaria = proj.get("diaria", pd.DataFrame())
    ating_d = (
        diaria[diaria["tipo"] == "Atingida"]
        if (not diaria.empty and "tipo" in diaria.columns)
        else pd.DataFrame()
    )
    ritmo_reg = proj.get("ritmo_meta_diario", pd.DataFrame())
    ritmo_med = proj.get("ritmo_meta_diario_medias", pd.DataFrame())

    fig = go.Figure()

    if not ating_d.empty:
        textos_real = [fmt_qtd(float(v)) for v in ating_d["qtd"]]
        fig.add_trace(
            go.Scatter(
                x=ating_d["dia"],
                y=ating_d["qtd"],
                mode="lines+markers+text",
                name="Realizado até hoje",
                text=textos_real,
                textposition="top center",
                textfont=dict(size=10, color=COR_AZUL_ESC, family="Inter"),
                line=dict(color=COR_AZUL_ESC, width=3),
                marker=dict(size=8, color=COR_AZUL_ESC),
            )
        )

    if ritmo_reg is not None and not ritmo_reg.empty:
        textos_reg = [fmt_qtd(float(v)) for v in ritmo_reg["qtd"]]
        fig.add_trace(
            go.Scatter(
                x=ritmo_reg["dia"],
                y=ritmo_reg["qtd"],
                mode="lines+markers+text",
                name="Meta diária (regressão)",
                text=textos_reg,
                textposition="top center",
                textfont=dict(size=10, color=COR_VERMELHO, family="Inter"),
                line=dict(color=COR_VERMELHO, width=3),
                marker=dict(size=8, color=COR_VERMELHO),
            )
        )

    if ritmo_med is not None and not ritmo_med.empty:
        textos_med = [fmt_qtd(float(v)) for v in ritmo_med["qtd"]]
        fig.add_trace(
            go.Scatter(
                x=ritmo_med["dia"],
                y=ritmo_med["qtd"],
                mode="lines+markers+text",
                name="Meta diária (médias)",
                text=textos_med,
                textposition="bottom center",
                textfont=dict(size=10, color="#0f766e", family="Inter"),
                line=dict(color="#0f766e", width=3, dash="dash"),
                marker=dict(size=8, symbol="diamond", color="#0f766e"),
            )
        )

    fig.add_vline(
        x=dia_hoje, line_width=1, line_dash="dot", line_color="#64748b",
        annotation_text="Hoje", annotation_position="top",
        annotation_font=dict(color=COR_TEXTO_PRETO, size=11, family="Inter"),
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.10, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        hovermode="x unified",
        height=460,
    )
    fig.update_xaxes(
        title_text="Dia do mês",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
        range=[0.5, proj["ultimo_dia"] + 0.5],
    )
    fig.update_yaxes(
        title_text="Qtd. no dia",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
        key=_proj_plot_key(proj, "meta_diaria"),
    )


def _plot_comparativo_realizado_projetado(proj: Dict[str, Any]) -> None:
    """Gráfico: projetado × realizado por dia no mês atual (regressão e médias)."""
    st.markdown("##### Projetado × realizado por dia (mês atual)")

    dia_hoje = proj["hoje"].day
    df_reg = proj.get("comparativo_diario_reg", pd.DataFrame())
    df_med = proj.get("comparativo_diario_medias", pd.DataFrame())
    if (df_reg is None or df_reg.empty) and (df_med is None or df_med.empty):
        st.info("Sem dados para o comparativo diário.")
        return

    fig = go.Figure()

    # Realizado (uma série; igual em reg e médias)
    base = df_reg if (df_reg is not None and not df_reg.empty) else df_med
    real = base.dropna(subset=["realizado"]) if "realizado" in base.columns else pd.DataFrame()
    if not real.empty:
        textos_real = [fmt_qtd(float(v)) for v in real["realizado"]]
        fig.add_trace(
            go.Scatter(
                x=real["dia"],
                y=real["realizado"],
                mode="lines+markers+text",
                name="Realizado",
                text=textos_real,
                textposition="top center",
                textfont=dict(size=10, color=COR_AZUL_ESC, family="Inter"),
                line=dict(color=COR_AZUL_ESC, width=3),
                marker=dict(size=8, color=COR_AZUL_ESC),
            )
        )

    if df_reg is not None and not df_reg.empty:
        textos_reg = [fmt_qtd(float(v)) for v in df_reg["projetado"]]
        fig.add_trace(
            go.Scatter(
                x=df_reg["dia"],
                y=df_reg["projetado"],
                mode="lines+markers+text",
                name="Projetado (regressão)",
                text=textos_reg,
                textposition="bottom center",
                textfont=dict(size=10, color=COR_VERMELHO, family="Inter"),
                line=dict(color=COR_VERMELHO, width=3, dash="dash"),
                marker=dict(size=8, color=COR_VERMELHO),
            )
        )

    if df_med is not None and not df_med.empty:
        textos_med = [fmt_qtd(float(v)) for v in df_med["projetado"]]
        fig.add_trace(
            go.Scatter(
                x=df_med["dia"],
                y=df_med["projetado"],
                mode="lines+markers+text",
                name="Projetado (médias)",
                text=textos_med,
                textposition="top center",
                textfont=dict(size=10, color="#0f766e", family="Inter"),
                line=dict(color="#0f766e", width=3, dash="dot"),
                marker=dict(size=8, symbol="diamond", color="#0f766e"),
            )
        )

    fig.add_vline(
        x=dia_hoje, line_width=1, line_dash="dot", line_color="#64748b",
        annotation_text="Hoje", annotation_position="top",
        annotation_font=dict(color=COR_TEXTO_PRETO, size=11, family="Inter"),
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.10, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        hovermode="x unified",
        height=460,
    )
    fig.update_xaxes(
        title_text="Dia do mês",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
        range=[0.5, proj["ultimo_dia"] + 0.5],
    )
    fig.update_yaxes(
        title_text="Qtd. no dia",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
        key=_proj_plot_key(proj, "comparativo"),
    )


def render_projecao_vendas(
    proj: Dict[str, Any],
    titulo: Optional[str] = None,
) -> None:
    """Seção Streamlit: cartões + gráficos de projeção e gráfico único de meta diária."""
    incluir_mes = bool(proj.get("incluir_mes", True))
    if titulo is None:
        titulo = (
            "Projeção de Vendas"
            if incluir_mes
            else "Projeção de Vendas (sem efeito de mês)"
        )

    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader(titulo)

    st.markdown(
        f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi">
                <div class="lbl">Realizado MTD</div>
                <div class="val">{fmt_qtd(proj['qtd_mtd'])}</div>
            </div>
            <div class="vel-kpi">
                <div class="lbl">Qtd. (regressão)</div>
                <div class="val">{fmt_qtd(proj['qtd_projetada_mes'])}</div>
            </div>
            <div class="vel-kpi">
                <div class="lbl">Qtd. (médias)</div>
                <div class="val">{fmt_qtd(proj.get('qtd_projetada_medias', 0))}</div>
            </div>
            <div class="vel-kpi">
                <div class="lbl">VGV (regressão)</div>
                <div class="val val--red">{fmt_br_milhoes(proj['vgv_projetado'])}</div>
            </div>
            <div class="vel-kpi">
                <div class="lbl">VGV (médias)</div>
                <div class="val val--red">{fmt_br_milhoes(proj.get('vgv_projetado_medias', 0))}</div>
            </div>
            <div class="vel-kpi">
                <div class="lbl">% VGV reg. / médias</div>
                <div class="val">{fmt_pct_valor(proj['pct_vgv_meta'])} / {fmt_pct_valor(proj.get('pct_vgv_medias', 0))}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    _plot_projecao_mes(
        "Projeção por regressão",
        f"R² treino: {fmt_num(proj['r2_treino'])}",
        proj,
        proj["diaria"],
        proj["acumulado"],
        chart_id="reg",
    )

    med = proj.get("medias") or {}
    mu = float(med.get("mu") or 0.0)
    _plot_projecao_mes(
        "Projeção por médias sazonais",
        f"μ: {fmt_num(mu)}",
        proj,
        proj.get("diaria_medias", pd.DataFrame()),
        proj.get("acumulado_medias", pd.DataFrame()),
        chart_id="med",
    )

    st.markdown("<br/>", unsafe_allow_html=True)
    _plot_comparativo_realizado_projetado(proj)

    st.markdown("<br/>", unsafe_allow_html=True)
    st.subheader(
        "Meta diária para bater a quantidade"
        if incluir_mes
        else "Meta diária para bater a quantidade (sem efeito de mês)"
    )
    _plot_meta_diaria_integrada(proj)


def _plot_barra_efeitos(
    titulo: str,
    df: pd.DataFrame,
    cor: str,
    referencia: str,
    chart_key: Optional[str] = None,
) -> None:
    """Barra de efeitos relativos (índice; referência = 1.0)."""
    st.markdown(f"##### {titulo}")
    st.caption(f"Referência: {referencia} = 1,00")
    if df is None or df.empty:
        return

    textos = [fmt_num(float(v)) for v in df["indice"]]
    cores = [COR_TEXTO_PRETO if abs(float(v) - 1.0) < 1e-9 else cor for v in df["indice"]]

    fig = go.Figure(
        go.Bar(
            x=df["categoria"],
            y=df["indice"],
            text=textos,
            textposition="outside",
            textfont=dict(size=10, color=COR_TEXTO_PRETO, family="Inter"),
            marker_color=cores,
            name="Índice",
            hovertemplate="%{x}<br>Índice: %{y:.2f}<br>Efeito (qtd): %{customdata:.2f}<extra></extra>",
            customdata=df["efeito"],
        )
    )
    fig.add_hline(
        y=1.0, line_width=1, line_dash="dot", line_color="#64748b",
        annotation_text="Ref. = 1", annotation_position="top left",
        annotation_font=dict(color=COR_TEXTO_PRETO, size=10),
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        showlegend=False,
        height=360,
        bargap=0.25,
    )
    fig.update_xaxes(
        title_text="",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter", size=11),
    )
    fig.update_yaxes(
        title_text="Índice (ref. = 1)",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
        key=chart_key or _plotly_key("efeitos_sazonais", titulo),
    )


def render_efeitos_sazonais(efeitos: Dict[str, Any]) -> None:
    """Seção: efeitos de dia da semana, dia do mês e mês (relativos à referência)."""
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader("Efeitos sazonais (regressão)")
    st.caption(
        f"R²: {fmt_num(float(efeitos.get('r2', 0)))} · "
        f"Baseline (segunda + dia 1 + janeiro): {fmt_num(float(efeitos.get('intercepto', 0)))} vendas/dia"
    )

    _plot_barra_efeitos(
        "Efeito de dia da semana",
        efeitos.get("dia_semana", pd.DataFrame()),
        COR_AZUL_ESC,
        "segunda-feira",
        chart_key=_plotly_key("efeitos_sazonais", "dia_semana"),
    )
    _plot_barra_efeitos(
        "Efeito de dia do mês",
        efeitos.get("dia_mes", pd.DataFrame()),
        COR_VERMELHO,
        "dia 1",
        chart_key=_plotly_key("efeitos_sazonais", "dia_mes"),
    )
    _plot_barra_efeitos(
        "Efeito de mês",
        efeitos.get("mes", pd.DataFrame()),
        "#0f766e",
        "janeiro",
        chart_key=_plotly_key("efeitos_sazonais", "mes"),
    )


# -----------------------------------------------------------------------------
# Funil comercial: agendamentos → visitas → pastas → pastas aprovadas → vendas
# -----------------------------------------------------------------------------

def _aplicar_secrets_salesforce() -> None:
    """Copia [salesforce] dos secrets Streamlit para variáveis de ambiente."""
    try:
        if hasattr(st, "secrets") and "salesforce" in st.secrets:
            sec = st.secrets["salesforce"]
            if sec.get("USER"):
                os.environ["SALESFORCE_USER"] = str(sec["USER"]).strip()
            if sec.get("PASSWORD"):
                os.environ["SALESFORCE_PASSWORD"] = str(sec["PASSWORD"]).strip()
            if sec.get("TOKEN"):
                os.environ["SALESFORCE_TOKEN"] = str(sec["TOKEN"]).strip()
            dom = str(sec.get("DOMAIN") or sec.get("domain") or "").strip()
            if dom:
                os.environ["SALESFORCE_DOMAIN"] = dom
    except Exception:
        pass


def conectar_salesforce_app() -> Tuple[Any, Optional[str]]:
    """Conecta ao Salesforce via simple_salesforce. Retorna (cliente, erro)."""
    import inspect

    for fr in inspect.stack()[1:10]:
        fn = (fr.filename or "").replace("\\", "/")
        if "velocimetro_precompute" in fn or "sf_sync" in fn or "run_precompute" in fn:
            break
    else:
        return None, "Salesforce desabilitado no painel — use Google Sheets."

    _aplicar_secrets_salesforce()
    try:
        from simple_salesforce import Salesforce, SalesforceAuthenticationFailed
    except ImportError:
        return None, "Pacote simple-salesforce não instalado (pip install simple-salesforce)."

    username = (os.environ.get("SALESFORCE_USER") or "").strip()
    password = (os.environ.get("SALESFORCE_PASSWORD") or "").strip()
    token = (os.environ.get("SALESFORCE_TOKEN") or "").strip()
    domain = (os.environ.get("SALESFORCE_DOMAIN") or "login").strip() or "login"
    if not username or not password:
        return None, "Credenciais Salesforce ausentes ([salesforce] USER/PASSWORD/TOKEN nos secrets)."
    try:
        kwargs: Dict[str, Any] = {
            "username": username,
            "password": password,
            "domain": domain,
        }
        if token:
            kwargs["security_token"] = token
        return Salesforce(**kwargs), None
    except SalesforceAuthenticationFailed as e:
        return None, f"Autenticação Salesforce recusada: {e}"
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"


@st.cache_resource(ttl=3300, show_spinner=False)
def _cliente_salesforce_cache():
    """Reutiliza sessão SF entre consultas (evita 3+ logins por refresh)."""
    if _painel_apenas_cache():
        raise CachePainelIndisponivel(
            "Consulta Salesforce bloqueada — painel em modo cache (abas «Cache · *»)."
        )
    sf, err = conectar_salesforce_app()
    if sf is None:
        raise RuntimeError(err or "Falha ao conectar no Salesforce.")
    return sf


# Analytics API: teto duro de detalhe por chamada. CSV Details: ~100k.
SF_ANALYTICS_ROW_CAP = 2000
SF_CSV_SOFT_CAP = 95_000


def _sf_session_bits(sf):
    """Retorna (base_url, session_id, headers) para export/API."""
    base = (getattr(sf, "sf_instance", None) or "").rstrip("/")
    if base and not base.startswith("http"):
        base = f"https://{base}"
    sid = str(getattr(sf, "session_id", "") or "")
    headers = dict(getattr(sf, "headers", {}) or {})
    if sid and "Authorization" not in headers:
        headers["Authorization"] = f"Bearer {sid}"
    return base, sid, headers


def _relatorio_sf_via_csv(sf, report_id: str):
    """
    Export CSV do relatório (Details). Contorna o teto de 2k da Analytics API,
    mas o Salesforce ainda pode cortar ~100k linhas em exports grandes.
    """
    import requests
    from io import StringIO

    rid = (report_id or "").strip()
    if not rid:
        return pd.DataFrame()

    base, sid, headers = _sf_session_bits(sf)
    if not base or not sid:
        raise ValueError("Sessão Salesforce sem instance/session_id.")

    urls = [
        f"{base}/{rid}?isdtp=p1&export=1&enc=UTF-8&xf=csv",
        f"{base}/{rid}?isdtp=p1&export=1&enc=UTF-8&xf=csv&detailsOnly=1",
        f"{base}/servlet/PrintableViewDownloadServlet?isdtp=p1&reportid={rid}",
    ]
    erros = []
    for url in urls:
        try:
            resp = requests.get(
                url,
                headers=headers,
                cookies={"sid": sid},
                timeout=900,
                allow_redirects=True,
            )
            resp.raise_for_status()
            text = resp.content.decode("utf-8-sig", errors="replace")
            sample = text.lstrip()[:300].lower()
            if (
                sample.startswith("<!doctype")
                or sample.startswith("<html")
                or "<table" in sample[:800]
                or ("login" in sample[:400] and "password" in sample[:800])
            ):
                erros.append(f"HTML em {url.split('?')[0][-40:]}")
                continue
            df = pd.read_csv(StringIO(text), low_memory=False)
            if df is None or df.empty:
                erros.append("CSV vazio")
                continue
            return df
        except Exception as e:
            erros.append(f"{type(e).__name__}: {e}")
    raise ValueError("Export CSV falhou: " + " | ".join(erros[:4]))


def _analytics_raw_to_df(raw):
    meta = (raw.get("reportMetadata") or {})
    cols_meta = meta.get("detailColumns") or []
    ext = ((raw.get("reportExtendedMetadata") or {}).get("detailColumnInfo") or {})
    headers = []
    for c in cols_meta:
        info = ext.get(c) or {}
        headers.append(str(info.get("label") or c))

    def _cell_scalar(cell):
        if not isinstance(cell, dict):
            return cell
        v = cell.get("value")
        # value às vezes vem como dict/OrderedDict (lookup) — usa label
        if isinstance(v, (dict, list, tuple)):
            v = cell.get("label")
        elif v is None:
            v = cell.get("label")
        return v

    rows_out = []
    fact = (raw.get("factMap") or {}).get("T!T") or {}
    for row in fact.get("rows") or []:
        cells = row.get("dataCells") or []
        vals = [_cell_scalar(cell) for cell in cells]
        if len(vals) < len(headers):
            vals = vals + [None] * (len(headers) - len(vals))
        rows_out.append(vals[: len(headers)])
    if not headers:
        return pd.DataFrame()
    return pd.DataFrame(rows_out, columns=headers)


def _analytics_run(sf, report_id: str, report_metadata=None, tentativas: int = 8):
    """Executa relatório; em rate-limit (500/hora) espera e tenta de novo."""
    import time as _time

    rid = (report_id or "").strip()
    ultimo = None
    for i in range(max(1, int(tentativas))):
        try:
            if report_metadata is None:
                return sf.restful(f"analytics/reports/{rid}", params={"includeDetails": "true"})
            return sf.restful(
                f"analytics/reports/{rid}",
                method="POST",
                json={"reportMetadata": report_metadata},
            )
        except Exception as e:
            ultimo = e
            msg = str(e).lower()
            rate = (
                "500" in msg and "relat" in msg
            ) or "forbidden" in msg and ("60 minuto" in msg or "60 minute" in msg or "rate" in msg)
            rate = rate or ("não é possível executar mais de 500" in msg) or (
                "nao e possivel executar mais de 500" in msg
            )
            if not rate or i >= tentativas - 1:
                raise
            # backoff crescente (rate limit SF: 500 sync / 60 min)
            espera = min(90, 15 * (i + 1))
            _time.sleep(espera)
    raise ultimo  # pragma: no cover



def _analytics_pick_date_column(meta, ext):
    sdf = meta.get("standardDateFilter") or {}
    col = sdf.get("column")
    if col:
        return str(col)
    for c in meta.get("detailColumns") or []:
        info = ext.get(c) or {}
        data_type = str(info.get("dataType") or "").lower()
        label = str(info.get("label") or c).lower()
        api = str(c).lower()
        if data_type in ("date", "datetime") or "date" in api or "data" in label:
            return str(c)
    return None


def _analytics_pick_id_column(meta, ext):
    """Escolhe coluna estável p/ keyset — evita falso positivo (ex.: 'credito' contém 'id')."""
    cols = list(meta.get("detailColumns") or [])
    scored = []
    for c in cols:
        info = ext.get(c) or {}
        api = str(c)
        label = str(info.get("label") or "")
        api_l = api.lower()
        lab_l = label.lower()
        score = 0
        # Id clássico / lookup Id
        if api_l == "id" or api_l.endswith(".id") or api_l.endswith("_id") or api_l.endswith("id__c"):
            score += 100
        if re.search(r"(^|[._])id($|[._])", api_l):
            score += 80
        if "identificador" in api_l or "identificador" in lab_l:
            score += 90
        if "codigo_do_agendamento" in api_l or "código do agendamento" in lab_l or "codigo do agendamento" in lab_l:
            score += 95
        if "código" in lab_l or "codigo" in lab_l:
            score += 50
        # Nome da avaliação de crédito (pastas) — chave de dedup
        if "nome da avalia" in lab_l or "nome_da_avalia" in api_l.replace("__", "_").lower():
            score += 70
        # Evita campos de status/texto genéricos
        if any(x in lab_l for x in ("status", "contrato", "valor", "telefone", "email")):
            score -= 60
        if score > 0:
            scored.append((score, api))
    if scored:
        scored.sort(key=lambda t: (t[0], t[1]), reverse=True)
        return scored[0][1]
    return cols[0] if cols else None


def _analytics_apply_date_filter(meta, date_col: str, ini, fim):
    m = copy.deepcopy(meta)
    m["standardDateFilter"] = {
        "column": date_col,
        "durationValue": "CUSTOM",
        "startDate": ini.isoformat(),
        "endDate": fim.isoformat(),
    }
    return m


def _analytics_keyset_pages(sf, report_id: str, base_meta, date_col: str, id_col: str, ini, fim):
    """Pagina um intervalo (ex.: 1 dia) com filtro greaterThan no Id ordenado."""
    partes = []
    last_val = None
    vistos = set()
    for _ in range(500):
        meta = _analytics_apply_date_filter(base_meta, date_col, ini, fim)
        meta["sortBy"] = [{"sortColumn": id_col, "sortOrder": "Asc"}]
        filtros = [
            f for f in (meta.get("reportFilters") or [])
            if not (f.get("column") == id_col and f.get("operator") == "greaterThan")
        ]
        if last_val is not None:
            filtros.append({"column": id_col, "operator": "greaterThan", "value": str(last_val)})
        meta["reportFilters"] = filtros
        orig_bool = str(base_meta.get("reportBooleanFilter") or "").strip()
        if last_val is not None:
            idx = len(filtros)
            meta["reportBooleanFilter"] = f"({orig_bool}) AND {idx}" if orig_bool else str(idx)
        else:
            meta["reportBooleanFilter"] = orig_bool or None

        raw = _analytics_run(sf, report_id, meta)
        df = _analytics_raw_to_df(raw)
        if df.empty:
            break
        ext = ((raw.get("reportExtendedMetadata") or {}).get("detailColumnInfo") or {})
        id_label = str((ext.get(id_col) or {}).get("label") or id_col)
        col_id = id_label if id_label in df.columns else (id_col if id_col in df.columns else df.columns[0])
        novos = df[~df[col_id].astype(str).isin(vistos)] if col_id in df.columns else df
        if novos.empty:
            break
        partes.append(novos)
        vals = novos[col_id].astype(str).tolist()
        vistos.update(vals)
        last_val = vals[-1]
        all_data = bool(raw.get("allData", False))
        if all_data or len(df) < SF_ANALYTICS_ROW_CAP:
            break
    if not partes:
        return pd.DataFrame()
    return pd.concat(partes, ignore_index=True)


def _analytics_fetch_range(sf, report_id: str, base_meta, date_col: str, id_col, ini, fim, profundidade: int = 0):
    """Baixa um intervalo; se bater 2k, divide a janela (ou usa keyset no dia)."""
    if ini > fim:
        return pd.DataFrame()
    meta = _analytics_apply_date_filter(base_meta, date_col, ini, fim)
    meta.pop("sortBy", None)

    raw = _analytics_run(sf, report_id, meta)
    df = _analytics_raw_to_df(raw)
    n = len(df)
    all_data = bool(raw.get("allData", True)) if n < SF_ANALYTICS_ROW_CAP else bool(raw.get("allData", False))

    if n == 0:
        return df
    if n < SF_ANALYTICS_ROW_CAP or all_data:
        return df

    if ini == fim:
        if id_col:
            return _analytics_keyset_pages(sf, report_id, base_meta, date_col, id_col, ini, fim)
        return df

    if profundidade > 40:
        if id_col:
            return _analytics_keyset_pages(sf, report_id, base_meta, date_col, id_col, ini, fim)
        return df

    mid = ini + timedelta(days=(fim - ini).days // 2)
    if mid >= fim:
        mid = fim - timedelta(days=1)
    if mid < ini:
        mid = ini

    esq = _analytics_fetch_range(sf, report_id, base_meta, date_col, id_col, ini, mid, profundidade + 1)
    dir_ = _analytics_fetch_range(
        sf, report_id, base_meta, date_col, id_col, mid + timedelta(days=1), fim, profundidade + 1
    )
    if esq.empty:
        return dir_
    if dir_.empty:
        return esq
    return pd.concat([esq, dir_], ignore_index=True)


def _relatorio_sf_via_analytics(sf, report_id: str):
    """Uma chamada Analytics (máx. ~2k) — só para metadados / fallback mínimo."""
    raw = _analytics_run(sf, report_id)
    return _analytics_raw_to_df(raw)


def _relatorio_sf_via_analytics_chunked(sf, report_id: str, anos_historico: int = 4, chunk_dias: int = 14):
    """
    Contorna o limite de 2.000 linhas da Analytics API:
    fatia por data (standardDateFilter) e, se um dia ainda passar de 2k, pagina por Id.
    """
    rid = (report_id or "").strip()
    raw0 = _analytics_run(sf, rid)
    base_meta = copy.deepcopy(raw0.get("reportMetadata") or {})
    ext = ((raw0.get("reportExtendedMetadata") or {}).get("detailColumnInfo") or {})
    date_col = _analytics_pick_date_column(base_meta, ext)
    if not date_col:
        return _analytics_raw_to_df(raw0)

    id_col = _analytics_pick_id_column(base_meta, ext)
    hoje = date.today()
    ini_global = date(hoje.year - int(anos_historico), 1, 1)
    sdf = base_meta.get("standardDateFilter") or {}
    try:
        sd = sdf.get("startDate")
        if sd:
            ini_meta = date.fromisoformat(str(sd)[:10])
            if ini_meta < ini_global:
                ini_global = ini_meta
    except Exception:
        pass

    partes = []
    cursor = ini_global
    while cursor <= hoje:
        fim_chunk = min(cursor + timedelta(days=chunk_dias - 1), hoje)
        parte = _analytics_fetch_range(sf, rid, base_meta, date_col, id_col, cursor, fim_chunk)
        if not parte.empty:
            partes.append(parte)
        cursor = fim_chunk + timedelta(days=1)

    if not partes:
        return pd.DataFrame()
    out = pd.concat(partes, ignore_index=True)
    # Dedup seguro (células podem ter tipos mistos)
    try:
        return out.drop_duplicates().reset_index(drop=True)
    except TypeError:
        return out.astype(str).drop_duplicates().reset_index(drop=True)


def _sf_rel_name(val):
    if isinstance(val, dict):
        return val.get("Name") or val.get("name")
    return None


def _sf_janela_painel_vendas(ref: Optional[date] = None) -> date:
    """Janela enxuta do painel: últimos PAINEL_MESES_VENDAS meses (inclui MTD)."""
    ref = ref or date.today()
    y, m = ref.year, ref.month
    m -= int(PAINEL_MESES_VENDAS) - 1
    while m <= 0:
        m += 12
        y -= 1
    return date(y, m, 1)


def _sf_janela_12_meses_fechados(ref: Optional[date] = None) -> Tuple[date, date]:
    """12 meses-calendário anteriores, excluindo o mês atual."""
    ref = ref or date.today()
    inicio = date(ref.year - 1, ref.month, 1)
    fim = date(ref.year, ref.month, 1) - timedelta(days=1)
    return inicio, fim


def _sf_janela_36_meses_fechados(ref: Optional[date] = None) -> Tuple[date, date]:
    """36 meses-calendário anteriores, excluindo o mês atual."""
    ref = ref or date.today()
    inicio = date(ref.year - FUNIL_ANOS_TREINO, ref.month, 1)
    fim = date(ref.year, ref.month, 1) - timedelta(days=1)
    return inicio, fim


def _sf_inicio_producao(ref: Optional[date] = None) -> date:
    """1º dia do mês atual menos o buffer de lags (produção leve)."""
    ref = ref or date.today()
    inicio_mes = date(ref.year, ref.month, 1)
    return inicio_mes - timedelta(days=int(FUNIL_SOQL_BUFFER_LAGS))


def _sf_soql_desde(modo_janela: str = "producao") -> str:
    """
    Início da janela SOQL.
    - painel: PAINEL_MESES_VENDAS (24m) — KPI + projeção OLS no painel web
    - treino / historico: 36 meses (script offline)
    - producao: 28 dias antes do 1º dia do mês atual (MTD + lags do funil)
    """
    modo = (modo_janela or "producao").strip().lower()
    if modo == "painel":
        ini = _sf_janela_painel_vendas()
    elif modo in ("treino", "historico"):
        ini, _ = _sf_janela_36_meses_fechados()
    else:
        ini = _sf_inicio_producao()
    return f"{ini.isoformat()}T00:00:00Z"


def _sf_soql_agendamentos(sf, modo_janela: str = "producao") -> pd.DataFrame:
    """
    Extrai agendamentos/visitas via SOQL em Event (queryMore).
    Contorna Analytics 2k e cota de 500 relatórios/hora.
    """
    desde = _sf_soql_desde(modo_janela)
    desde_date = desde[:10]
    # Produção: também puxa visitas cuja criação é anterior ao buffer.
    filtro_tempo = (
        f"(CreatedDate >= {desde} OR Data_da_Visita__c >= {desde_date})"
        if (modo_janela or "").strip().lower() != "treino"
        else f"CreatedDate >= {desde}"
    )
    soql = (
        "SELECT Codigo_do_agendamento__c, CreatedDate, Data_da_Visita__c, "
        "Nome_do_empreendimento__c, WhatId, AccountId "
        "FROM Event "
        "WHERE Unidade_de_negocio__c = 'Direcional' "
        "AND Regional__c = 'RJ' "
        "AND PDV__r.Regional_Comercial__c = 'RJ' "
        "AND PDV__r.UnidadeDeNegocio__c = 'Direcional' "
        "AND Empreendimento_de_interesse__c != null "
        "AND Account.Regional_Comercial__c = 'RJ' "
        f"AND {filtro_tempo}"
    )
    res = sf.query_all(soql)
    rows = [
        {
            "Código do agendamento": r.get("Codigo_do_agendamento__c"),
            "Data de criação": r.get("CreatedDate"),
            "Data da visita": r.get("Data_da_Visita__c"),
            "Empreendimento": (r.get("Nome_do_empreendimento__c") or "").strip(),
            "WhatId": r.get("WhatId"),
            "AccountId": r.get("AccountId"),
        }
        for r in (res.get("records") or [])
    ]
    return pd.DataFrame(rows)


_PASTAS_PC_CAMPOS = (
    "Name, CreatedDate, dataPrimeiroEnvioAnalise__c, dataAprovacaoSAFI__c, "
    "Empreendimento__r.Name, Tipo__c, Oportunidade__c, Conta__c, "
    "Renda__c, Valor_da_Renda__c, RendaApurada__c, "
    "Valor_FGTS__c, FGTS_apurado__c, "
    "Valor_Financiamento__c, Valor_de_Subsidio__c"
)


def _sf_soql_pastas(sf, modo_janela: str = "producao") -> pd.DataFrame:
    """Extrai avaliações de crédito (pastas) via SOQL."""
    desde = _sf_soql_desde(modo_janela)
    desde_date = desde[:10]
    filtro_tempo = (
        f"(CreatedDate >= {desde} OR dataPrimeiroEnvioAnalise__c >= {desde_date} "
        f"OR dataAprovacaoSAFI__c >= {desde_date})"
        if (modo_janela or "").strip().lower() != "treino"
        else f"CreatedDate >= {desde}"
    )
    soql = (
        f"SELECT {_PASTAS_PC_CAMPOS} "
        "FROM Avaliacao_credito__c "
        "WHERE Empreendimento__r.Regional__c = 'RJ' "
        "AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
        f"AND {filtro_tempo}"
    )
    res = sf.query_all(soql)
    rows = [
        {
            "Nome da Avaliação de crédito": r.get("Name"),
            "Data de criação": r.get("CreatedDate"),
            "Data Primeiro Envio Análise": r.get("dataPrimeiroEnvioAnalise__c"),
            "Data Aprovação SAFI": r.get("dataAprovacaoSAFI__c"),
            "Empreendimento": _sf_rel_name(r.get("Empreendimento__r")),
            "Tipo": r.get("Tipo__c"),
            "Oportunidade": r.get("Oportunidade__c"),
            "Conta": r.get("Conta__c"),
            "Renda": r.get("Renda__c"),
            "Valor da Renda": r.get("Valor_da_Renda__c"),
            "Renda Apurada": r.get("RendaApurada__c"),
            "Valor FGTS": r.get("Valor_FGTS__c"),
            "FGTS apurado": r.get("FGTS_apurado__c"),
            "Valor do Financiamento": r.get("Valor_Financiamento__c"),
            "Subsídio": r.get("Valor_de_Subsidio__c"),
            "Valor do Subsidio": r.get("Valor_de_Subsidio__c"),
            "Valor de Subsidio": r.get("Valor_de_Subsidio__c"),
        }
        for r in (res.get("records") or [])
    ]
    return pd.DataFrame(rows)


def _sf_soql_vendas(sf, modo_janela: str = "producao") -> pd.DataFrame:
    """
    Equivalente ao relatório 00O3Z000005ZsPmUAK sem o filtro de tempo próprio:
    vendas comerciais Direcional/RJ. Em treino usa 36 meses; em produção, buffer.
    """
    desde = _sf_soql_desde(modo_janela)[:10]  # ContratoGeradoEm__c é Date, não DateTime
    soql = (
        "SELECT Id, Name, Empreendimento__r.Name, Empreendimento__r.Regional__c, "
        "Valor_Real_de_Venda__c, Owner.Name, DirecionalVendas__c, ContratoGeradoEm__c, "
        "DataVenda__c, Termo_de_reserva__c, Ranking__c, M_s_Venda__c, Ano_da_Venda__c, "
        "Imobiliaria__r.Name, Contato_Corretor_Proprietario1__r.Name, Gerente_regional__c, "
        "Diretor_de_vendas__c, Regional__c, OrigemConta__c, LeadSource, "
        "Venda_Futura__c, GeradoComunicadoVenda__c "
        "FROM Opportunity "
        "WHERE DirecionalVendas__c = true "
        "AND Empreendimento__r.Regional__c = 'RJ' "
        "AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
        f"AND ContratoGeradoEm__c >= {desde}"
    )
    res = sf.query_all(soql)
    rows = []
    for r in res.get("records") or []:
        emp = r.get("Empreendimento__r") if isinstance(r.get("Empreendimento__r"), dict) else {}
        regiao = (emp.get("Regional__c") if emp else None) or r.get("Regional__c")
        imob = _sf_rel_name(r.get("Imobiliaria__r"))
        rows.append({
            "ID da Oportunidade": r.get("Id"),
            "Nome da oportunidade": r.get("Name"),
            "Empreendimento": _sf_rel_name(emp),
            "Região": regiao,
            "Imobiliária": imob,
            "Canal": canal_de_imobiliaria(imob),
            "Valor Real de Venda": r.get("Valor_Real_de_Venda__c"),
            "Proprietário da oportunidade": _sf_rel_name(r.get("Owner")),
            "Venda Comercial?": r.get("DirecionalVendas__c"),
            "Venda facilitada": r.get("Termo_de_reserva__c"),
            "Ranking": r.get("Ranking__c"),
            "Data da venda": r.get("DataVenda__c"),
            "Mês Venda": r.get("M_s_Venda__c"),
            "Ano da Venda": r.get("Ano_da_Venda__c"),
            "Contato Corretor Proprietario": _sf_rel_name(
                r.get("Contato_Corretor_Proprietario1__r")
            ),
            "Contrato gerado em": r.get("ContratoGeradoEm__c"),
            "Gerente regional": r.get("Gerente_regional__c"),
            "Diretor de vendas": r.get("Diretor_de_vendas__c"),
            "Regional": r.get("Regional__c"),
            "Origem da Conta": r.get("OrigemConta__c"),
            "Origem do lead": r.get("LeadSource"),
            "Venda futura": r.get("Venda_Futura__c"),
            "Venda comunicada": r.get("GeradoComunicadoVenda__c"),
        })
    return pd.DataFrame(rows)


def _sf_soql_cotacoes(sf, modo_janela: str = "producao") -> pd.DataFrame:
    """Cotações (pro soluto, VCX, sinal) vinculadas a oportunidades RJ."""
    desde = _sf_soql_desde(modo_janela)[:10]
    soql_com_sinal = (
        "SELECT Id, OpportunityId, Opportunity.Empreendimento__r.Name, "
        "VoltaAoCaixa__c, PercentualdoProSoluto__c, TemProSoluto__c, "
        "TotalSinalCom__c, CreatedDate "
        "FROM Quote "
        "WHERE Opportunity.Empreendimento__r.Regional__c = 'RJ' "
        "AND Opportunity.Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
        f"AND CreatedDate >= {desde}T00:00:00Z"
    )
    soql_base = (
        "SELECT Id, OpportunityId, Opportunity.Empreendimento__r.Name, "
        "VoltaAoCaixa__c, PercentualdoProSoluto__c, TemProSoluto__c, CreatedDate "
        "FROM Quote "
        "WHERE Opportunity.Empreendimento__r.Regional__c = 'RJ' "
        "AND Opportunity.Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
        f"AND CreatedDate >= {desde}T00:00:00Z"
    )
    com_sinal = True
    try:
        res = sf.query_all(soql_com_sinal)
    except Exception:
        com_sinal = False
        try:
            res = sf.query_all(soql_base)
        except Exception:
            return pd.DataFrame()
    rows = []
    for r in res.get("records") or []:
        opp = r.get("Opportunity") if isinstance(r.get("Opportunity"), dict) else {}
        emp = _sf_rel_name((opp or {}).get("Empreendimento__r"))
        rows.append({
            "ID Cotação": r.get("Id"),
            "ID da Oportunidade": r.get("OpportunityId"),
            "Empreendimento": emp,
            "Volta ao caixa": r.get("VoltaAoCaixa__c"),
            "Percentual Pro Soluto": r.get("PercentualdoProSoluto__c"),
            "Tem Pro Soluto": r.get("TemProSoluto__c"),
            "Total Sinal Com": r.get("TotalSinalCom__c") if com_sinal else None,
            "Data de criação": r.get("CreatedDate"),
        })
    return pd.DataFrame(rows)


def _sf_soql_estoque_empreendimento(sf) -> pd.DataFrame:
    """
    Estoque por unidade (Produto__c) — equivalente ao relatório SF_REPORT_ESTOQUE_ID.
    Status: Disponível, Mirror, Fora de venda, Fora de Venda - Comercial.
    """
    statuses = ", ".join(f"'{s}'" for s in ESTOQUE_STATUS_TODOS)
    soql = (
        "SELECT Id, StatusUnidade__c, Empreendimento__r.Name, "
        "Identificador__c, ValorFinalComKit__c, Valor_de_Avalia_o_Banc_ria__c, "
        "Valor_Folga__c, B_nus_Adimpl_ncia__c, Area__c, "
        "Tipologia__c, Empreendimento__r.Possui_Investidor__c, "
        "Empreendimento__r.DataExpedicaoHabitese__c "
        "FROM Produto__c "
        "WHERE Empreendimento__r.Regional__c = 'RJ' "
        "AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
        f"AND StatusUnidade__c IN ({statuses})"
    )
    res = sf.query_all(soql)
    rows = []
    for r in res.get("records") or []:
        emp = _sf_rel_name(r.get("Empreendimento__r"))
        emp_rel = r.get("Empreendimento__r") if isinstance(r.get("Empreendimento__r"), dict) else {}
        hab_emp = emp_rel.get("DataExpedicaoHabitese__c") if emp_rel else None
        rows.append({
            "ID da Unidade": r.get("Id"),
            "StatusUnidade__c": r.get("StatusUnidade__c"),
            "Empreendimento": emp,
            "Identificador": r.get("Identificador__c"),
            "Valor Final com Kit": r.get("ValorFinalComKit__c"),
            "Valor de Avaliação Bancária": r.get("Valor_de_Avalia_o_Banc_ria__c"),
            "Valor Folga": r.get("Valor_Folga__c"),
            "Bônus Adimplência": r.get("B_nus_Adimpl_ncia__c"),
            "Area": r.get("Area__c"),
            "Habite-se": hab_emp,
            "Tipologia": r.get("Tipologia__c"),
            "Possui Investidor": emp_rel.get("Possui_Investidor__c") if emp_rel else None,
        })
    return pd.DataFrame(rows)


def _sf_soql_contagem_unidades_total_por_emp(sf) -> Dict[str, int]:
    """Contagem de todas as unidades (Produto__c) por empreendimento, qualquer status."""
    soql = (
        "SELECT Id, Empreendimento__r.Name "
        "FROM Produto__c "
        "WHERE Empreendimento__r.Regional__c = 'RJ' "
        "AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional'"
    )
    try:
        res = sf.query_all(soql)
    except Exception:
        return {}
    out: Dict[str, int] = {}
    for r in res.get("records") or []:
        emp = _limpar_emp(_sf_rel_name(r.get("Empreendimento__r")))
        if emp:
            out[emp] = out.get(emp, 0) + 1
    return out


def _sf_inicio_funil_empreendimento(hoje: Optional[date] = None) -> date:
    """Desde o 1º dia do mês ou 29 dias atrás (cobre MTD + rolling 30d)."""
    hoje = hoje or date.today()
    ini_mes = date(hoje.year, hoje.month, 1)
    return min(ini_mes, hoje - timedelta(days=29))


def _sf_soql_agendamentos_empreendimento(sf, hoje: Optional[date] = None) -> pd.DataFrame:
    """Agendamentos/visitas com empreendimento (aba Funil por Empreendimento)."""
    hoje = hoje or date.today()
    ini = _sf_inicio_funil_empreendimento(hoje)
    desde = f"{ini.isoformat()}T00:00:00Z"
    desde_date = ini.isoformat()
    filtro_tempo = f"(CreatedDate >= {desde} OR Data_da_Visita__c >= {desde_date})"
    soql = (
        "SELECT Codigo_do_agendamento__c, CreatedDate, Data_da_Visita__c, "
        "Nome_do_empreendimento__c "
        "FROM Event "
        "WHERE Unidade_de_negocio__c = 'Direcional' "
        "AND Regional__c = 'RJ' "
        "AND PDV__r.Regional_Comercial__c = 'RJ' "
        "AND PDV__r.UnidadeDeNegocio__c = 'Direcional' "
        "AND Empreendimento_de_interesse__c != null "
        "AND Account.Regional_Comercial__c = 'RJ' "
        f"AND {filtro_tempo}"
    )
    res = sf.query_all(soql)
    rows = []
    for r in res.get("records") or []:
        emp = (r.get("Nome_do_empreendimento__c") or "").strip()
        rows.append({
            "Código do agendamento": r.get("Codigo_do_agendamento__c"),
            "Data de criação": r.get("CreatedDate"),
            "Data da visita": r.get("Data_da_Visita__c"),
            "Empreendimento": emp,
        })
    return pd.DataFrame(rows)


def _sf_soql_pastas_empreendimento(sf, hoje: Optional[date] = None) -> pd.DataFrame:
    """Pastas com empreendimento (aba Funil por Empreendimento)."""
    hoje = hoje or date.today()
    ini = _sf_inicio_funil_empreendimento(hoje)
    desde = f"{ini.isoformat()}T00:00:00Z"
    desde_date = ini.isoformat()
    filtro_tempo = (
        f"(CreatedDate >= {desde} OR dataPrimeiroEnvioAnalise__c >= {desde_date} "
        f"OR dataAprovacaoSAFI__c >= {desde_date})"
    )
    soql = (
        f"SELECT {_PASTAS_PC_CAMPOS} "
        "FROM Avaliacao_credito__c "
        "WHERE Empreendimento__r.Regional__c = 'RJ' "
        "AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
        f"AND {filtro_tempo}"
    )
    res = sf.query_all(soql)
    rows = []
    for r in res.get("records") or []:
        emp = _sf_rel_name(r.get("Empreendimento__r"))
        rows.append({
            "Nome da Avaliação de crédito": r.get("Name"),
            "Data de criação": r.get("CreatedDate"),
            "Data Primeiro Envio Análise": r.get("dataPrimeiroEnvioAnalise__c"),
            "Data Aprovação SAFI": r.get("dataAprovacaoSAFI__c"),
            "Empreendimento": emp,
            "Tipo": r.get("Tipo__c"),
            "Oportunidade": r.get("Oportunidade__c"),
            "Conta": r.get("Conta__c"),
            "Renda": r.get("Renda__c"),
            "Valor da Renda": r.get("Valor_da_Renda__c"),
            "Renda Apurada": r.get("RendaApurada__c"),
            "Valor FGTS": r.get("Valor_FGTS__c"),
            "FGTS apurado": r.get("FGTS_apurado__c"),
            "Valor do Financiamento": r.get("Valor_Financiamento__c"),
            "Subsídio": r.get("Valor_de_Subsidio__c"),
            "Valor do Subsidio": r.get("Valor_de_Subsidio__c"),
            "Valor de Subsidio": r.get("Valor_de_Subsidio__c"),
        })
    return pd.DataFrame(rows)


def _sf_soql_tabela_comprometimento(sf, modo_janela: str = "painel") -> pd.DataFrame:
    """Tabela Comprometimento de Renda — pro soluto máximo por oportunidade."""
    desde = _sf_soql_desde(modo_janela)[:10]
    soql = (
        "SELECT Oportunidade__c, Renda__c, ProSoluto__c, ComprometimentoDeRenda__c, "
        "ComprometimentoDeRendaParcial__c, numParcelas__c "
        "FROM TabelaComprometimentoDeRenda__c "
        "WHERE Oportunidade__r.Empreendimento__r.Regional__c = 'RJ' "
        "AND Oportunidade__r.Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
        f"AND CreatedDate >= {desde}T00:00:00Z"
    )
    try:
        res = sf.query_all(soql)
    except Exception:
        return pd.DataFrame()
    rows = []
    for r in res.get("records") or []:
        rows.append({
            "Oportunidade__c": r.get("Oportunidade__c"),
            "ProSoluto__c": r.get("ProSoluto__c"),
            "Renda__c": r.get("Renda__c"),
            "ComprometimentoDeRenda__c": r.get("ComprometimentoDeRenda__c"),
            "ComprometimentoDeRendaParcial__c": r.get("ComprometimentoDeRendaParcial__c"),
            "numParcelas__c": r.get("numParcelas__c"),
        })
    return pd.DataFrame(rows)


def _sf_soql_oportunidades_empreendimento(sf, hoje: Optional[date] = None) -> pd.DataFrame:
    """Oportunidades abertas (pipeline) + criadas/movimentadas na janela rolling 7d."""
    hoje = hoje or date.today()
    ini = _sf_inicio_funil_empreendimento(hoje)
    desde = f"{ini.isoformat()}T00:00:00Z"
    soql = (
        "SELECT Id, StageName, CreatedDate, LastStageChangeDate, IsClosed, IsWon, "
        "Empreendimento__r.Name, OrigemConta__c, LeadSource "
        "FROM Opportunity "
        "WHERE Empreendimento__r.Regional__c = 'RJ' "
        "AND Empreendimento__r.UnidadeDeNegocio__c = 'Direcional' "
        "AND Empreendimento__c != null "
        f"AND (IsClosed = false OR CreatedDate >= {desde} OR LastStageChangeDate >= {desde})"
    )
    res = sf.query_all(soql)
    rows = []
    for r in res.get("records") or []:
        rows.append({
            "ID da Oportunidade": r.get("Id"),
            "Fase": r.get("StageName"),
            "Data de criação": r.get("CreatedDate"),
            "Data mudança fase": r.get("LastStageChangeDate"),
            "Empreendimento": _sf_rel_name(r.get("Empreendimento__r")),
            "Origem da Conta": r.get("OrigemConta__c"),
            "Origem do lead": r.get("LeadSource"),
            "Fechada": r.get("IsClosed"),
            "Ganha": r.get("IsWon"),
        })
    return pd.DataFrame(rows)


@st.cache_data(ttl=600, show_spinner="Carregando funil por empreendimento (Google Sheets)…")
def carregar_funil_empreendimento_sf(forcar_sf: Optional[bool] = None) -> Dict[str, Any]:
    """Pacote MTD+7d: agendamentos, pastas, oportunidades por empreendimento."""
    cf = _cred_fp_atual()
    df_ag, o_ag = _ler_dado_painel("funil_emp_ag", cf)
    if df_ag.empty:
        df_ag, o_ag = _ler_dado_painel("funil_ag", cf)
    df_pas, o_pas = _ler_dado_painel("funil_emp_pastas", cf)
    if df_pas.empty:
        df_pas, o_pas = _ler_dado_painel("funil_pastas", cf)
    df_opps, _ = _ler_dado_painel("funil_emp_opps", cf)
    df_ven, o_ven = _ler_dado_painel("funil_emp_ven", cf)
    if df_ven.empty:
        df_ven, o_ven = _ler_dado_painel("vendas_raw", cf)
    df_est, o_est = _ler_dado_painel("funil_emp_est", cf)
    if df_est.empty:
        df_est, o_est = _ler_dado_painel("estoque", cf)
    if not df_ven.empty:
        df_ven = filtrar_vendas_comerciais(df_ven)
        ini = _sf_inicio_funil_empreendimento(date.today())
        col_c = achar_coluna(df_ven, ALIASES_CONTRATO_GERADO)
        if col_c:
            dt = parse_data_serie(df_ven[col_c])
            df_ven = df_ven.loc[dt.notna() & (dt.dt.date >= ini)].copy()
    origem = " · ".join(x for x in (o_ag, o_pas, o_ven, o_est) if x) or "Google Sheets"
    return {
        "agendamentos": df_ag,
        "pastas": df_pas,
        "oportunidades": df_opps if df_opps is not None else pd.DataFrame(),
        "vendas": df_ven,
        "estoque": df_est,
        "inicio_janela": _sf_inicio_funil_empreendimento(date.today()).isoformat(),
        "timings": {"total_s": 0.0},
        "origem": origem,
    }


def _carregar_funil_empreendimento_sf_live() -> Dict[str, Any]:
    t0 = time.perf_counter()
    sf = _cliente_salesforce_cache()
    hoje = date.today()
    df_ag = normalizar_colunas(_sf_soql_agendamentos_empreendimento(sf, hoje))
    df_pas = normalizar_colunas(_sf_soql_pastas_empreendimento(sf, hoje))
    df_opps = normalizar_colunas(_sf_soql_oportunidades_empreendimento(sf, hoje))
    df_ven = normalizar_colunas(_sf_soql_vendas(sf, modo_janela="producao"))
    df_est = normalizar_colunas(_sf_soql_estoque_empreendimento(sf))
    df_ven = filtrar_vendas_comerciais(df_ven)
    ini = _sf_inicio_funil_empreendimento(hoje)
    if not df_ven.empty:
        col_c = achar_coluna(df_ven, ALIASES_CONTRATO_GERADO)
        if col_c:
            dt = parse_data_serie(df_ven[col_c])
            df_ven = df_ven.loc[dt.notna() & (dt.dt.date >= ini)].copy()
    t_total = time.perf_counter() - t0
    return {
        "agendamentos": df_ag if df_ag is not None else pd.DataFrame(),
        "pastas": df_pas if df_pas is not None else pd.DataFrame(),
        "oportunidades": df_opps if df_opps is not None else pd.DataFrame(),
        "vendas": df_ven if df_ven is not None else pd.DataFrame(),
        "estoque": df_est if df_est is not None else pd.DataFrame(),
        "inicio_janela": ini.isoformat(),
        "timings": {"total_s": t_total},
    }


def _sf_soql_por_relatorio(sf, report_id: str, rotulo: str, modo_janela: str = "producao"):
    """Tenta SOQL para relatórios grandes. Retorna (df, origem) ou (None, None)."""
    rid = (report_id or "").strip()
    rotulo_l = (rotulo or "").lower()
    modo = (modo_janela or "producao").strip().lower()
    try:
        if rid == SF_REPORT_AGENDAMENTOS_ID or "agendamento" in rotulo_l:
            df = _sf_soql_agendamentos(sf, modo_janela=modo)
            if df is not None and not df.empty:
                origem = (
                    f"Salesforce SOQL · Event · {rotulo} · janela={modo} · "
                    f"{len(df):,} linhas".replace(",", ".")
                )
                return df, origem
        if rid == SF_REPORT_PASTAS_ID or "pasta" in rotulo_l:
            df = _sf_soql_pastas(sf, modo_janela=modo)
            if df is not None and not df.empty:
                origem = (
                    f"Salesforce SOQL · Avaliacao_credito__c · {rotulo} · janela={modo} · "
                    f"{len(df):,} linhas".replace(",", ".")
                )
                return df, origem
        if rid == SF_REPORT_VENDAS_ID or "venda" in rotulo_l:
            df = _sf_soql_vendas(sf, modo_janela=modo)
            if df is not None and not df.empty:
                origem = (
                    f"Salesforce SOQL · Opportunity · {rotulo} "
                    f"(relatório {SF_REPORT_VENDAS_ID}, janela={modo}) · "
                    f"{len(df):,} linhas".replace(",", ".")
                )
                return df, origem
    except Exception:
        return None, None
    return None, None


def _recortar_vendas_painel(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica recorte temporal de vendas do painel (24m)."""
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    col_data = achar_coluna(df, ALIASES_CONTRATO_GERADO)
    if not col_data:
        return df
    inicio_hist = _sf_janela_painel_vendas()
    dt = parse_data_serie(df[col_data])
    return df.loc[dt.notna() & (dt.dt.date >= inicio_hist)].copy()


@st.cache_data(ttl=600, show_spinner="Carregando vendas (Google Sheets)…")
def carregar_vendas_painel_sf(forcar_sf: Optional[bool] = None) -> Dict[str, Any]:
    """Vendas comerciais — cache ou aba BD Vendas Completa."""
    cf = _cred_fp_atual()
    df, origem = _ler_dado_painel("vendas_raw", cf)
    n_v = len(df)
    return {
        "vendas": df,
        "timings": {"total_s": 0.0},
        "origem_vendas": origem or f"Sheets · {WS_VENDAS}",
    }


def _carregar_vendas_painel_sf_live() -> Dict[str, Any]:
    """Vendas comerciais — janela curta (24m). Bloqueia só o KPI inicial."""
    t0 = time.perf_counter()
    sf = _cliente_salesforce_cache()
    t1 = time.perf_counter()
    df_vendas = _sf_soql_vendas(sf, modo_janela="painel")
    t_vendas = time.perf_counter() - t1
    df_vendas = normalizar_colunas(df_vendas) if df_vendas is not None and not df_vendas.empty else pd.DataFrame()
    df_vendas = _recortar_vendas_painel(df_vendas)
    t_total = time.perf_counter() - t0
    n_v = len(df_vendas)
    return {
        "vendas": df_vendas,
        "timings": {"vendas_s": t_vendas, "total_s": t_total},
        "origem_vendas": (
            f"Salesforce SOQL · Opportunity · vendas (painel {PAINEL_MESES_VENDAS}m) · "
            f"{n_v:,} linhas".replace(",", ".")
        ),
    }


@st.cache_data(ttl=600, show_spinner="Carregando poder de compra (Google Sheets)…")
def carregar_pacote_poder_compra_sf(forcar_sf: Optional[bool] = None) -> Dict[str, Any]:
    """Pastas aprovadas + tabela comprometimento de renda."""
    cf = _cred_fp_atual()
    df_pas, _ = _ler_dado_painel("pc_pastas", cf)
    df_tab, _ = _ler_dado_painel("pc_tabela", cf)
    return {
        "pastas_aprovadas": df_pas,
        "tabela_comprometimento": df_tab,
    }


def _carregar_pacote_poder_compra_sf_live() -> Dict[str, Any]:
    sf = _cliente_salesforce_cache()
    df_pas = normalizar_colunas(_sf_soql_pastas(sf, modo_janela="painel"))
    df_pas = deduplicar_pastas_aprovadas_funil(df_pas) if not df_pas.empty else df_pas
    df_tab = normalizar_colunas(_sf_soql_tabela_comprometimento(sf, modo_janela="painel"))
    return {
        "pastas_aprovadas": df_pas,
        "tabela_comprometimento": df_tab,
    }


@st.cache_data(ttl=600, show_spinner="Carregando cotações (Google Sheets)…")
def carregar_cotacoes_painel_sf(forcar_sf: Optional[bool] = None) -> pd.DataFrame:
    cf = _cred_fp_atual()
    df, _ = _ler_dado_painel("cotacoes", cf)
    return df if df is not None else pd.DataFrame()


def _carregar_cotacoes_painel_sf_live() -> pd.DataFrame:
    sf = _cliente_salesforce_cache()
    df = _sf_soql_cotacoes(sf, modo_janela="painel")
    return normalizar_colunas(df) if df is not None and not df.empty else pd.DataFrame()


@st.cache_data(ttl=600, show_spinner="Carregando estoque (Google Sheets)…")
def carregar_estoque_painel_sf(forcar_sf: Optional[bool] = None) -> pd.DataFrame:
    """Estoque Produto__c (RJ · Direcional) para KPI e tabela analítica."""
    cf = _cred_fp_atual()
    df, _ = _ler_dado_painel("estoque", cf)
    return df if df is not None else pd.DataFrame()


def _carregar_estoque_painel_sf_live() -> pd.DataFrame:
    sf = _cliente_salesforce_cache()
    df = _sf_soql_estoque_empreendimento(sf)
    return normalizar_colunas(df) if df is not None and not df.empty else pd.DataFrame()


@st.cache_data(ttl=600, show_spinner="Carregando totais de unidades (Google Sheets)…")
def carregar_total_unidades_por_emp_sf() -> Dict[str, int]:
    """Total de unidades por empreendimento."""
    info = _info_gsheets_atual()
    m = vc.ler_manifest(info) if info else {}
    tot = _total_unidades_do_manifest(m) if m else {}
    if tot:
        return tot
    tot_est = _total_unidades_do_estoque_cache()
    if tot_est:
        return tot_est
    cf = _cred_fp_atual()
    df, _ = _ler_dado_painel("estoque", cf)
    if df is not None and not df.empty and "Empreendimento" in df.columns:
        cont = df.groupby(df["Empreendimento"].map(_limpar_emp)).size()
        return {str(k): int(v) for k, v in cont.items()}
    return {}


@st.cache_data(ttl=600, show_spinner="Carregando funil (Google Sheets)…")
def carregar_funil_painel_sf(forcar_sf: Optional[bool] = None) -> Dict[str, Any]:
    """Agendamentos + pastas — janela produção (~2 meses)."""
    cf = _cred_fp_atual()
    df_ag, o_ag = _ler_dado_painel("funil_ag", cf)
    df_pastas, o_pas = _ler_dado_painel("funil_pastas", cf)
    n_a, n_p = len(df_ag), len(df_pastas)
    return {
        "agendamentos": df_ag,
        "pastas": df_pastas,
        "timings": {"total_s": 0.0},
        "origem_ag": o_ag or f"Sheets · {ABA_AGENDAMENTOS_VISITAS} · {n_a:,} linhas".replace(",", "."),
        "origem_pastas": o_pas or f"Sheets · pastas · {n_p:,} linhas".replace(",", "."),
    }


def _carregar_funil_painel_sf_live() -> Dict[str, Any]:
    t0 = time.perf_counter()
    sf = _cliente_salesforce_cache()
    t1 = time.perf_counter()
    df_ag = _sf_soql_agendamentos(sf, modo_janela="producao")
    df_pastas = _sf_soql_pastas(sf, modo_janela="producao")
    t_funil = time.perf_counter() - t1
    df_ag = normalizar_colunas(df_ag) if df_ag is not None and not df_ag.empty else pd.DataFrame()
    df_pastas = normalizar_colunas(df_pastas) if df_pastas is not None and not df_pastas.empty else pd.DataFrame()
    t_total = time.perf_counter() - t0
    n_a, n_p = len(df_ag), len(df_pastas)
    return {
        "agendamentos": df_ag,
        "pastas": df_pastas,
        "timings": {"funil_s": t_funil, "total_s": t_total},
        "origem_ag": (
            f"Salesforce SOQL · Event · agendamentos (produção) · "
            f"{n_a:,} linhas".replace(",", ".")
        ),
        "origem_pastas": (
            f"Salesforce SOQL · Avaliacao_credito__c · pastas (produção) · "
            f"{n_p:,} linhas".replace(",", ".")
        ),
    }


@st.cache_data(ttl=600, show_spinner="Carregando histórico funil (Google Sheets)…")
def carregar_funil_historico_painel_sf(forcar_sf: Optional[bool] = None) -> Dict[str, Any]:
    """Agendamentos + pastas — janela painel (24m) para comparativos MTD."""
    cf = _cred_fp_atual()
    df_ag, _ = _ler_dado_painel("funil_hist_ag", cf)
    if df_ag.empty:
        df_ag, _ = _ler_dado_painel("funil_ag", cf)
    df_pastas, _ = _ler_dado_painel("funil_hist_pastas", cf)
    if df_pastas.empty:
        df_pastas, _ = _ler_dado_painel("funil_pastas", cf)
    return {
        "agendamentos": df_ag,
        "pastas": df_pastas,
        "timings": {"total_s": 0.0},
    }


def _carregar_funil_historico_painel_sf_live() -> Dict[str, Any]:
    t0 = time.perf_counter()
    sf = _cliente_salesforce_cache()
    df_ag = _sf_soql_agendamentos(sf, modo_janela="painel")
    df_pastas = _sf_soql_pastas(sf, modo_janela="painel")
    df_ag = normalizar_colunas(df_ag) if df_ag is not None and not df_ag.empty else pd.DataFrame()
    df_pastas = normalizar_colunas(df_pastas) if df_pastas is not None and not df_pastas.empty else pd.DataFrame()
    return {
        "agendamentos": df_ag,
        "pastas": df_pastas,
        "timings": {"total_s": time.perf_counter() - t0},
    }


@st.cache_data(ttl=3600, show_spinner="Carregando bases Salesforce…")
def carregar_pacote_painel_sf() -> Dict[str, Any]:
    """Pacote completo (benchmark). Produção usa vendas + funil separados."""
    vendas = carregar_vendas_painel_sf()
    funil = carregar_funil_painel_sf()
    tv = (vendas.get("timings") or {}).get("total_s", 0.0)
    tf = (funil.get("timings") or {}).get("total_s", 0.0)
    return {
        "vendas": vendas["vendas"],
        "agendamentos": funil["agendamentos"],
        "pastas": funil["pastas"],
        "timings": {
            "vendas_s": (vendas.get("timings") or {}).get("vendas_s", tv),
            "funil_s": (funil.get("timings") or {}).get("funil_s", tf),
            "total_s": float(tv) + float(tf),
        },
        "origem_vendas": vendas["origem_vendas"],
        "origem_ag": funil["origem_ag"],
        "origem_pastas": funil["origem_pastas"],
    }


@st.cache_data(ttl=3600, show_spinner="Baixando dados Salesforce (SOQL)…")
def carregar_relatorio_salesforce(
    report_id: str,
    rotulo: str = "relatório",
    modo_janela: str = "producao",
):
    """
    Baixa dados Salesforce sem teto de 2k:
    1) SOQL query_all (Event / Avaliacao_credito) — preferencial p/ volumes grandes
    2) CSV do relatório
    3) Analytics fatiada (fallback; cota 500/h)

    modo_janela:
      - painel: vendas 24 meses (painel web)
      - producao: buffer curto (MTD + 28d de lags)
      - treino: 36 meses fechados (script offline)
    """
    if _painel_apenas_cache():
        _bloquear_sf_live(f"relatorio:{rotulo}")
    try:
        sf = _cliente_salesforce_cache()
    except CachePainelIndisponivel:
        raise
    except Exception:
        sf, err = conectar_salesforce_app()
        if sf is None:
            raise RuntimeError(err or "Falha ao conectar no Salesforce.")

    rid = (report_id or "").strip()
    if not rid:
        raise RuntimeError(f"Report ID vazio ({rotulo}).")

    tentativas = []
    modo = (modo_janela or "producao").strip().lower() or "producao"

    # 1) SOQL direto (melhor caminho para ~180k)
    df_soql, origem_soql = _sf_soql_por_relatorio(sf, rid, rotulo, modo_janela=modo)
    if df_soql is not None and not df_soql.empty:
        df = normalizar_colunas(df_soql)
        return df, origem_soql or f"Salesforce SOQL · {rotulo} · {len(df)} linhas"

    # 2) CSV
    df_csv = None
    try:
        df_csv = _relatorio_sf_via_csv(sf, rid)
    except Exception as e_csv:
        tentativas.append(f"CSV: {e_csv}")

    rotulo_l = (rotulo or "").lower()
    precisa_chunk = False
    if df_csv is None or df_csv.empty:
        precisa_chunk = True
    else:
        n_csv = len(df_csv)
        if SF_ANALYTICS_ROW_CAP - 5 <= n_csv <= SF_ANALYTICS_ROW_CAP + 50:
            precisa_chunk = True
        if n_csv >= SF_CSV_SOFT_CAP:
            precisa_chunk = True
        if "agendamento" in rotulo_l and n_csv < 120_000:
            precisa_chunk = True
        if "pasta" in rotulo_l and n_csv <= SF_ANALYTICS_ROW_CAP + 50:
            precisa_chunk = True

    df = df_csv
    origem = f"Salesforce CSV · {rotulo} · {rid}" if df_csv is not None and not df_csv.empty else ""

    # 3) Analytics fatiada
    if precisa_chunk:
        try:
            df_chunk = _relatorio_sf_via_analytics_chunked(sf, rid)
            if df_chunk is not None and not df_chunk.empty:
                if df is None or df.empty or len(df_chunk) > len(df):
                    df = df_chunk
                    origem = (
                        f"Salesforce Analytics fatiada · {rotulo} · {rid} · "
                        f"{len(df_chunk):,} linhas".replace(",", ".")
                    )
                else:
                    origem = (
                        f"Salesforce CSV · {rotulo} · {rid} · "
                        f"{len(df):,} linhas (chunk≤CSV)".replace(",", ".")
                    )
        except Exception as e_an:
            tentativas.append(f"Analytics fatiada: {e_an}")
            if df is None or df.empty:
                try:
                    df = _relatorio_sf_via_analytics(sf, rid)
                    origem = f"Salesforce Analytics · {rotulo} · {rid} (truncado ~2k)"
                except Exception as e_an2:
                    tentativas.append(f"Analytics: {e_an2}")
                    raise RuntimeError(
                        f"Não foi possível baixar o {rotulo} ({rid}). "
                        + " | ".join(tentativas)
                    ) from e_an2

    if df is None or df.empty:
        raise RuntimeError(
            f"Relatório {rotulo} ({rid}) retornou vazio. " + " | ".join(tentativas)
        )

    df = normalizar_colunas(df)
    if rid == SF_REPORT_VENDAS_ID:
        col_data = achar_coluna(df, ALIASES_CONTRATO_GERADO)
        if col_data:
            if modo == "painel":
                inicio_hist = _sf_janela_painel_vendas()
            else:
                inicio_hist, _ = _sf_janela_36_meses_fechados()
            dt = parse_data_serie(df[col_data])
            df = df.loc[dt.notna() & (dt.dt.date >= inicio_hist)].copy()
    if not origem:
        origem = f"Salesforce · {rotulo} · {rid} · {len(df):,} linhas".replace(",", ".")
    elif "linhas" not in origem.lower():
        origem = f"{origem} · {len(df):,} linhas".replace(",", ".")
    return df, origem

def carregar_agendamentos_visitas_salesforce(
    report_id: str = SF_REPORT_AGENDAMENTOS_ID,
) -> Tuple[pd.DataFrame, str]:
    """Compatibilidade: delega para carregar_relatorio_salesforce."""
    return carregar_relatorio_salesforce(report_id, rotulo="agendamentos/visitas")


def contar_eventos_por_dia(df: pd.DataFrame, aliases: List[str]) -> Dict[date, float]:
    """Conta ocorrências por dia a partir de uma coluna de data."""
    col = achar_coluna(df, aliases)
    if not col or df.empty:
        return {}
    return contar_eventos_por_coluna(df, col)


def contar_eventos_por_coluna(df: pd.DataFrame, col: str) -> Dict[date, float]:
    """Conta ocorrências por dia a partir de uma coluna já resolvida."""
    if not col or col not in df.columns or df.empty:
        return {}
    dt = parse_data_serie(df[col]).dropna()
    if dt.empty:
        return {}
    vc = dt.dt.normalize().value_counts()
    return {pd.Timestamp(k).date(): float(v) for k, v in vc.items()}


def deduplicar_pastas_aprovadas_funil(df: pd.DataFrame) -> pd.DataFrame:
    """
    Uma linha por Nome da Avaliação de crédito, mantendo a Data Aprovação SAFI
    mais recente.
    """
    col_safi = achar_coluna_aprovacao_safi(df)
    if not col_safi:
        return df if df is not None else pd.DataFrame()
    return deduplicar_por_chave_mais_recente(
        df, ALIASES_NOME_AVALIACAO_CREDITO, [col_safi]
    )


def deduplicar_por_chave_mais_recente(
    df: pd.DataFrame,
    aliases_chave: List[str],
    aliases_data: List[str],
) -> pd.DataFrame:
    """
    Remove duplicatas pela chave, mantendo a linha com a data mais recente.
    Se chave ou data não existirem, devolve o DataFrame original.
    """
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    col_chave = achar_coluna(df, aliases_chave)
    col_data = achar_coluna(df, aliases_data)
    if not col_chave or not col_data:
        return df
    out = df.copy()
    out["_dedup_dt"] = parse_data_serie(out[col_data])
    out["_dedup_key"] = out[col_chave].astype(str).str.strip()
    # Chaves vazias/NaN não entram na deduplicação entre si
    mask_ok = out["_dedup_key"].ne("") & out["_dedup_key"].str.lower().ne("nan")
    ok = out.loc[mask_ok].sort_values("_dedup_dt", ascending=False, na_position="last")
    ok = ok.drop_duplicates(subset=["_dedup_key"], keep="first")
    resto = out.loc[~mask_ok]
    out = pd.concat([ok, resto], ignore_index=True)
    return out.drop(columns=["_dedup_dt", "_dedup_key"], errors="ignore")


ALIASES_ID_OPORTUNIDADE = [
    "ID da Oportunidade", "Id da Oportunidade", "Opportunity ID", "Opportunity Id",
    "ID Oportunidade", "Id Oportunidade",
]
ALIASES_CONTRATO_GERADO = [
    "Contrato gerado em", "Contrato Gerado em", "Contrato gerado",
    "Data do Contrato", "Data Contrato", "Close Date", "Data da venda",
    "Data Venda",
]
ALIASES_NOME_AVALIACAO_CREDITO = [
    "Nome da Avaliação de crédito", "Nome da Avaliacao de credito",
    "Nome da Avaliação de Crédito", "Nome da Avaliacao de Credito",
    "Nome Avaliação de crédito", "Nome Avaliacao de credito",
]
ALIASES_CODIGO_AGENDAMENTO = [
    "Código do agendamento", "Codigo do agendamento",
    "Código do Agendamento", "Codigo do Agendamento",
    "Código Agendamento", "Codigo Agendamento",
]
ALIASES_DATA_CRIACAO = [
    "Data de criação", "Data de criacao", "Data Criação", "Data Criacao",
    "Created Date", "Data de Criação", "Criado em", "Data criação",
]


def deduplicar_vendas_funil(df: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por ID da Oportunidade (Contrato gerado em mais recente)."""
    return deduplicar_por_chave_mais_recente(df, ALIASES_ID_OPORTUNIDADE, ALIASES_CONTRATO_GERADO)


def deduplicar_pastas_funil(df: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por Nome da Avaliação (Data Primeiro Envio Análise mais recente)."""
    col_envio = achar_coluna_primeiro_envio_analise(df)
    if not col_envio:
        return df if df is not None else pd.DataFrame()
    return deduplicar_por_chave_mais_recente(
        df, ALIASES_NOME_AVALIACAO_CREDITO, [col_envio]
    )


def deduplicar_agendamentos_funil(df: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por Código do agendamento (Data de criação mais recente)."""
    return deduplicar_por_chave_mais_recente(
        df, ALIASES_CODIGO_AGENDAMENTO, ALIASES_DATA_CRIACAO
    )


def montar_mapa_funil_diario(
    df_ag_vis: pd.DataFrame,
    df_pastas: pd.DataFrame,
    serie_vendas: Optional[pd.DataFrame] = None,
    df_vendas: Optional[pd.DataFrame] = None,
) -> Dict[str, Dict[date, float]]:
    """
    Séries diárias do funil:
      agendamentos ← Data de criação
      visitas ← Data da visita
      pastas ← Data Primeiro Envio Análise (relatório SF 00OU600000FEOoDMAX)
      pastas_aprovadas ← Data Aprovação SAFI
      vendas ← Contrato gerado em, somente vendas comerciais

    Deduplicação antes da contagem:
      vendas → ID da Oportunidade (Contrato gerado em mais recente)
      pastas → Nome da Avaliação (Data Primeiro Envio Análise mais recente)
      pastas aprovadas → Nome da Avaliação (Data Aprovação SAFI mais recente)
      agendamentos → Código do agendamento (Data de criação mais recente)
    """
    df_ag = deduplicar_agendamentos_funil(
        df_ag_vis if df_ag_vis is not None else pd.DataFrame()
    )
    df_pas_raw = df_pastas if df_pastas is not None else pd.DataFrame()
    df_pas = deduplicar_pastas_funil(df_pas_raw)
    df_pas_aprov = deduplicar_pastas_aprovadas_funil(df_pas_raw)

    mapa_vendas: Dict[date, float] = {}
    if df_vendas is not None and not df_vendas.empty:
        df_ven = deduplicar_vendas_funil(filtrar_vendas_comerciais(df_vendas))
        mapa_vendas = contar_eventos_por_dia(df_ven, ALIASES_CONTRATO_GERADO)
    elif serie_vendas is not None and not serie_vendas.empty:
        for _, r in serie_vendas.iterrows():
            mapa_vendas[r["data"]] = float(r["qtd"])

    col_envio = achar_coluna_primeiro_envio_analise(df_pas)
    col_safi_aprov = achar_coluna_aprovacao_safi(df_pas_aprov)
    mapa_pastas = (
        contar_eventos_por_coluna(df_pas, col_envio) if col_envio else {}
    )
    mapa_aprov = (
        contar_eventos_por_coluna(df_pas_aprov, col_safi_aprov) if col_safi_aprov else {}
    )

    return {
        "agendamentos": contar_eventos_por_dia(df_ag, ALIASES_DATA_CRIACAO),
        "visitas": contar_eventos_por_dia(
            df_ag,
            [
                "Data da visita", "Data da Visita", "Data visita", "Data Visita",
                "Activity Date", "Data da Atividade", "Data do agendamento",
                "Data Agendamento", "Start Date Time", "Data/Hora",
            ],
        ),
        "pastas": mapa_pastas,
        "pastas_aprovadas": mapa_aprov,
        "vendas": mapa_vendas,
    }


def calendario_funil_diario(
    inicio: date,
    fim: date,
    mapas: Dict[str, Dict[date, float]],
    lags: Tuple[int, ...] = FUNIL_LAGS,
) -> pd.DataFrame:
    """Calendário diário do funil com dummies, lags, conversões e força de trabalho."""
    idx = pd.date_range(inicio, fim, freq="D")
    cal = pd.DataFrame({"data": [d.date() for d in idx]})
    for etapa in FUNIL_ETAPAS:
        m = mapas.get(etapa) or {}
        cal[etapa] = cal["data"].map(lambda d, _m=m: float(_m.get(d, 0.0)))
    cal["dia_mes"] = cal["data"].map(lambda d: d.day)
    cal["dia_semana"] = cal["data"].map(lambda d: DIAS_SEMANA_PT[d.weekday()])
    cal["mes"] = cal["data"].map(lambda d: MESES_PT[d.month])
    # Soma dos volumes em 4 semanas anteriores: 1–7, 8–14, 15–21, 22–28.
    lag_cols: Dict[str, pd.Series] = {}
    for etapa in FUNIL_ETAPAS:
        for ini_lag, fim_lag in FUNIL_LAG_BLOCOS:
            janela = fim_lag - ini_lag + 1
            lag_cols[col_lag_bloco(etapa, ini_lag, fim_lag)] = (
                cal[etapa].shift(ini_lag).rolling(janela, min_periods=1).sum().fillna(0.0)
            )
    if lag_cols:
        cal = pd.concat([cal, pd.DataFrame(lag_cols, index=cal.index)], axis=1)
    cal = adicionar_conversoes_funil(cal)
    return adicionar_forca_trabalho(cal)


def adicionar_forca_trabalho(
    cal: pd.DataFrame,
    janela: int = FUNIL_JANELA_FORCA,
) -> pd.DataFrame:
    """
    Indicador de força de trabalho: atividade recente em TODAS as etapas do funil.
    Captura que o time trabalha em paralelo (pastas podem 'vazar' para visitas etc.).
    Usa média móvel das etapas (drivers + vendas) deslocada 1 dia (sem vazamento do dia).
    """
    out = cal.copy()
    partes: List[pd.Series] = []
    for e in FUNIL_ETAPAS:
        if e not in out.columns:
            continue
        s = pd.to_numeric(out[e], errors="coerce").fillna(0.0).astype(float)
        mu = float(s.mean()) if len(s) else 0.0
        sd = float(s.std()) if len(s) else 0.0
        if sd < 1e-9:
            z = s * 0.0
        else:
            z = (s - mu) / sd
        partes.append(z)
    if not partes:
        out["forca_trabalho"] = 0.0
        out["atividade_bruta"] = 0.0
        return out
    z_mean = sum(partes) / float(len(partes))
    cols_drv = [e for e in FUNIL_DRIVERS if e in out.columns]
    out["atividade_bruta"] = out[cols_drv].sum(axis=1) if cols_drv else 0.0
    # força = média móvel da atividade padronizada, sem o dia corrente
    out["forca_trabalho"] = (
        z_mean.rolling(janela, min_periods=1).mean().shift(1).fillna(0.0)
    )
    return out


def _atualizar_forca_trabalho_linha(
    cal: pd.DataFrame,
    i: int,
    janela: int = FUNIL_JANELA_FORCA,
) -> None:
    """Recalcula força de trabalho na linha i a partir do histórico até i-1."""
    try:
        pos = int(cal.index.get_loc(i))
    except Exception:
        pos = int(i)
    i0 = max(0, pos - janela)
    hist = cal.iloc[i0:pos]
    if hist.empty:
        cal.at[i, "forca_trabalho"] = 0.0
        cal.at[i, "atividade_bruta"] = 0.0
        return
    z_vals: List[float] = []
    for e in FUNIL_ETAPAS:
        if e not in hist.columns:
            continue
        s_all = pd.to_numeric(cal[e], errors="coerce").fillna(0.0)
        mu = float(s_all.mean()) if len(s_all) else 0.0
        sd = float(s_all.std()) if len(s_all) else 0.0
        s_h = pd.to_numeric(hist[e], errors="coerce").fillna(0.0)
        if sd < 1e-9:
            z_vals.append(0.0)
        else:
            z_vals.append(float(((s_h - mu) / sd).mean()))
    cal.at[i, "forca_trabalho"] = float(np.mean(z_vals)) if z_vals else 0.0
    cal.at[i, "atividade_bruta"] = float(
        sum(float(cal.at[i, e]) for e in FUNIL_DRIVERS if e in cal.columns)
    )


def _safe_ratio(num: pd.Series, den: pd.Series) -> pd.Series:
    """Razão segura; 0 quando denominador ≤ 0. Clip para estabilidade do modelo."""
    n = pd.to_numeric(num, errors="coerce").fillna(0.0).astype(float)
    d = pd.to_numeric(den, errors="coerce").fillna(0.0).astype(float)
    out = np.where(d > 0, n / d, 0.0)
    return pd.Series(np.clip(out, 0.0, 5.0), index=num.index)


def col_lag_bloco(etapa: str, ini_lag: int, fim_lag: int) -> str:
    return f"{etapa}_lag{ini_lag}_{fim_lag}"


def col_conv_etapa(origem: str, destino: str, janela: int) -> str:
    return f"conv_{origem}_{destino}_{int(janela)}d"


def cols_conversoes_funil() -> List[str]:
    """7 séries de conversão × janelas móveis de 7, 14, 21 e 28 dias."""
    return [
        col_conv_etapa(a, b, janela)
        for a, b in FUNIL_CONVERSOES
        for janela in FUNIL_JANELAS_CONV
    ]


def adicionar_conversoes_funil(
    cal: pd.DataFrame,
) -> pd.DataFrame:
    """
    7 conversões em janelas 7/14/21/28 dias, todas encerradas ontem.
    O shift(1) impede que o alvo do próprio dia entre nas explicativas.
    """
    out = cal.copy()
    conv_cols: Dict[str, pd.Series] = {}
    for a, b in FUNIL_CONVERSOES:
        for janela in FUNIL_JANELAS_CONV:
            min_periods = min(5, janela)
            num = out[b].rolling(janela, min_periods=min_periods).sum()
            den = out[a].rolling(janela, min_periods=min_periods).sum()
            conv_cols[col_conv_etapa(a, b, janela)] = (
                _safe_ratio(num, den).shift(1).fillna(0.0)
            )
    if conv_cols:
        out = pd.concat([out, pd.DataFrame(conv_cols, index=out.index)], axis=1)
    return out


def _atualizar_conversoes_linha(
    cal: pd.DataFrame,
    i: int,
) -> None:
    """Recalcula as 28 conversões usando somente os dias anteriores a i."""
    try:
        pos = int(cal.index.get_loc(i))
    except Exception:
        pos = int(i)
    for a, b in FUNIL_CONVERSOES:
        for janela in FUNIL_JANELAS_CONV:
            hist = cal.iloc[max(0, pos - janela):pos]
            den = float(hist[a].sum()) if not hist.empty else 0.0
            num = float(hist[b].sum()) if not hist.empty else 0.0
            cal.at[i, col_conv_etapa(a, b, janela)] = (
                float(np.clip(num / den, 0.0, 5.0)) if den > 0 else 0.0
            )


def taxa_conversao(origem: float, destino: float) -> Optional[float]:
    """Taxa percentual origem→destino; None se origem ≤ 0."""
    if origem is None or destino is None:
        return None
    o = float(origem)
    if o <= 0:
        return None
    return 100.0 * float(destino) / o


def calcular_conversoes_totais(totais: Dict[str, float]) -> Dict[str, Any]:
    """Conversões a partir de totais (MTD ou projetado do mês)."""
    etapa_a_etapa: List[Dict[str, Any]] = []
    for a, b in FUNIL_PARES_ETAPA:
        taxa = taxa_conversao(totais.get(a, 0.0), totais.get(b, 0.0))
        etapa_a_etapa.append({
            "origem": a,
            "destino": b,
            "label": f"{FUNIL_LABELS.get(a, a)} → {FUNIL_LABELS.get(b, b)}",
            "taxa": taxa,
            "origem_qtd": float(totais.get(a, 0.0)),
            "destino_qtd": float(totais.get(b, 0.0)),
        })
    para_venda: List[Dict[str, Any]] = []
    for a in FUNIL_DRIVERS:
        taxa = taxa_conversao(totais.get(a, 0.0), totais.get("vendas", 0.0))
        para_venda.append({
            "origem": a,
            "destino": "vendas",
            "label": f"{FUNIL_LABELS.get(a, a)} → Vendas",
            "taxa": taxa,
            "origem_qtd": float(totais.get(a, 0.0)),
            "destino_qtd": float(totais.get("vendas", 0.0)),
        })
    return {"etapa_a_etapa": etapa_a_etapa, "para_venda": para_venda}


def _funil_gap_vendas(
    gap_vendas: float,
    taxas_hist_frac: Dict[str, Optional[float]],
    totais_hist: Dict[str, float],
    gap_vendas_mes: Optional[float] = None,
    funil_necessario: Optional[Dict[str, float]] = None,
    totais_mtd: Optional[Dict[str, float]] = None,
) -> Dict[str, float]:
    """
    Funil por etapa a partir de um gap de vendas (restante para a meta).
    Usa conversão histórica indicador→venda; fallback proporcional ao gap do mês.
    """
    gap_v = max(0.0, float(gap_vendas))
    gap_mes = max(0.0, float(gap_vendas_mes if gap_vendas_mes is not None else gap_v))
    funil: Dict[str, float] = {"vendas": gap_v}
    for etapa in FUNIL_DRIVERS:
        t = taxas_hist_frac.get(etapa)
        if gap_v > 0 and t is not None and t > 1e-9:
            funil[etapa] = float(math.ceil(gap_v / t))
            continue
        v_h = float(totais_hist.get("vendas", 0.0))
        i_h = float(totais_hist.get(etapa, 0.0))
        if gap_v > 0 and v_h > 0 and i_h > 0:
            funil[etapa] = float(math.ceil(gap_v * (i_h / v_h)))
            continue
        if (
            gap_v > 0
            and gap_mes > 1e-9
            and funil_necessario is not None
            and totais_mtd is not None
        ):
            gap_etapa_mes = max(
                0.0,
                float(funil_necessario.get(etapa, 0.0)) - float(totais_mtd.get(etapa, 0.0)),
            )
            funil[etapa] = float(math.ceil(gap_etapa_mes * (gap_v / gap_mes)))
        else:
            funil[etapa] = 0.0
    return funil


def _pesos_distribuicao_gap(
    dias_distrib: List[date],
    resultados: Dict[str, Any],
) -> np.ndarray:
    """Pesos diários para distribuir o gap de vendas (projeção reg. de vendas)."""
    df_ven = (resultados.get("vendas") or {}).get("diaria", pd.DataFrame())
    pesos: List[float] = []
    for d in dias_distrib:
        p = 1.0
        if df_ven is not None and not df_ven.empty and "dia" in df_ven.columns:
            row = df_ven.loc[df_ven["dia"] == d.day]
            if not row.empty and "projetado_reg" in row.columns:
                p = max(float(row["projetado_reg"].iloc[0]), 0.0)
        pesos.append(p if p > 0 else 1.0)
    return np.asarray(pesos, dtype=float)


def _meta_vendas_por_pesos(
    gap_vendas: float,
    dias_alvo: List[date],
    dias_distrib: List[date],
    pesos: np.ndarray,
) -> float:
    """Parcela do gap de vendas alocada a um conjunto de dias (pesos normalizados)."""
    if gap_vendas <= 0 or not dias_alvo or not dias_distrib:
        return 0.0
    w = np.maximum(np.asarray(pesos, dtype=float), 0.0)
    soma = float(w.sum())
    if soma <= 1e-9:
        w = np.ones(len(dias_distrib), dtype=float)
        soma = float(len(dias_distrib))
    mapa = {dias_distrib[i]: float(w[i] / soma) for i in range(len(dias_distrib))}
    return float(gap_vendas) * sum(mapa.get(d, 0.0) for d in dias_alvo)


def _razoes_entre_funis(
    numerador: Dict[str, float],
    denominador: Dict[str, float],
) -> Dict[str, Optional[float]]:
    """Razão numerador/denominador por etapa (None se denominador ≤ 0)."""
    out: Dict[str, Optional[float]] = {}
    for etapa in FUNIL_ETAPAS:
        den = float(denominador.get(etapa, 0.0))
        num = float(numerador.get(etapa, 0.0))
        out[etapa] = (num / den) if den > 1e-9 else None
    return out


def _somar_funis(*funis: Dict[str, float]) -> Dict[str, float]:
    out: Dict[str, float] = {e: 0.0 for e in FUNIL_ETAPAS}
    for f in funis:
        for e in FUNIL_ETAPAS:
            out[e] += float((f or {}).get(e, 0.0))
    return out


def _taxas_cascata_historicas(totais_hist: Dict[str, float]) -> Dict[Tuple[str, str], float]:
    """Taxas fracionárias etapa→etapa do histórico, com limites operacionais."""
    taxas: Dict[Tuple[str, str], float] = {}
    for a, b in FUNIL_PARES_ETAPA:
        origem = float((totais_hist or {}).get(a, 0.0) or 0.0)
        destino = float((totais_hist or {}).get(b, 0.0) or 0.0)
        taxa = (destino / origem) if origem > 1e-9 else 0.0
        # Visitas→Pastas pode superar 100% por defasagem/backlog entre bases.
        limite = 2.5 if (a, b) == ("visitas", "pastas") else 0.98
        taxas[(a, b)] = float(np.clip(taxa, 1e-4, limite))
    return taxas


def _reconciliar_funil_hibrido(
    previsoes: Dict[str, float],
    taxas: Dict[Tuple[str, str], float],
    r2s: Optional[Dict[str, float]] = None,
    forca_conversao: float = 1.8,
) -> Dict[str, float]:
    """
    Reconcilia as cinco previsões próprias como um único funil.

    Resolve, em escala logarítmica, um mínimos quadrados com dois objetivos:
    preservar a estimativa de CADA etapa e manter as conversões próximas do
    histórico. O peso de cada previsão própria cresce com seu R².

    Depois aplica apenas um corredor de segurança às conversões (±40% do
    histórico). Isso impede 100% artificial em Aprovadas→Vendas sem transformar
    as etapas em simples múltiplos dos agendamentos.
    """
    eps = 1e-6
    n = len(FUNIL_ETAPAS)
    idx = {e: i for i, e in enumerate(FUNIL_ETAPAS)}
    linhas: List[np.ndarray] = []
    alvos: List[float] = []

    # Todos os modelos entram na solução — nenhum é descartado.
    for etapa in FUNIL_ETAPAS:
        r2 = float((r2s or {}).get(etapa, 0.0) or 0.0)
        peso = math.sqrt(float(np.clip(0.35 + max(r2, 0.0), 0.35, 1.35)))
        row = np.zeros(n, dtype=float)
        row[idx[etapa]] = peso
        linhas.append(row)
        alvos.append(peso * math.log(max(float(previsoes.get(etapa, 0.0) or 0.0), eps)))

    # Conversões ligam as previsões num sistema de funil, sem impor igualdade.
    peso_conv = math.sqrt(max(float(forca_conversao), 0.0))
    for a, b in FUNIL_PARES_ETAPA:
        taxa = max(float(taxas.get((a, b), eps)), eps)
        row = np.zeros(n, dtype=float)
        row[idx[b]] = peso_conv
        row[idx[a]] = -peso_conv
        linhas.append(row)
        alvos.append(peso_conv * math.log(taxa))

    A = np.vstack(linhas)
    y = np.asarray(alvos, dtype=float)
    x, *_ = np.linalg.lstsq(A, y, rcond=None)
    out = {e: max(0.0, float(math.exp(x[idx[e]]))) for e in FUNIL_ETAPAS}

    # Corredor largo: conserva complexidade, mas bloqueia inversões absurdas.
    for a, b in FUNIL_PARES_ETAPA:
        origem = out[a]
        if origem <= eps:
            out[b] = 0.0
            continue
        taxa_hist = float(taxas.get((a, b), 0.0))
        min_taxa = max(0.0, taxa_hist * 0.60)
        max_taxa = taxa_hist * 1.40
        if (a, b) != ("visitas", "pastas"):
            max_taxa = min(max_taxa, 0.98)
        taxa_atual = out[b] / origem
        out[b] = origem * float(np.clip(taxa_atual, min_taxa, max_taxa))
    return out


def _reconciliar_linha_cal_funil(
    cal: pd.DataFrame,
    i: int,
    taxas: Dict[Tuple[str, str], float],
    r2s: Optional[Dict[str, float]] = None,
    lags: Tuple[int, ...] = (),
) -> None:
    """Reconcilia numa linha as previsões próprias de todas as etapas."""
    bruto = {e: float(cal.at[i, e]) for e in FUNIL_ETAPAS}
    reconciliado = _reconciliar_funil_hibrido(bruto, taxas, r2s=r2s)
    for e, v in reconciliado.items():
        cal.at[i, e] = v
        if 0 in lags:
            col_l0 = f"{e}_lag0"
            if col_l0 in cal.columns:
                cal.at[i, col_l0] = v


def _domingo_semana_iso(d: date) -> date:
    """Último dia (domingo) da semana ISO que contém d."""
    return d + timedelta(days=(7 - d.isoweekday()))


def _cols_lag_funil(
    lags: Tuple[int, ...] = FUNIL_LAGS,
    alvo: Optional[str] = None,
    etapas: Optional[Tuple[str, ...]] = None,
) -> List[str]:
    """Quatro blocos semanais de volume para cada etapa."""
    etapas_uso = etapas if etapas is not None else FUNIL_ETAPAS
    return [
        col_lag_bloco(e, ini_lag, fim_lag)
        for e in etapas_uso
        for ini_lag, fim_lag in FUNIL_LAG_BLOCOS
    ]


def feature_names_funil_cal_lags(incluir_mes: bool = True) -> List[str]:
    """Ordem canônica das 70 features do modelo ElasticNet de produção."""
    n_cal = 31 + 7 + (12 if incluir_mes else 0)
    return [f"cal_{i}" for i in range(n_cal)] + _cols_lag_funil()


def _matriz_funil_cal_lags(
    df: pd.DataFrame,
    incluir_mes: bool = True,
) -> Tuple[np.ndarray, List[str]]:
    """
    Matriz do modelo de produção: calendário (sem intercepto) + 20 lags + intercepto.
    Total: 71 colunas (70 features + 1.0).
    """
    X_cal = _matriz_explicativas(df, incluir_mes=incluir_mes)
    lag_cols = _cols_lag_funil()
    X_lags = np.zeros((len(df), len(lag_cols)), dtype=float)
    for j, c in enumerate(lag_cols):
        if c in df.columns:
            X_lags[:, j] = pd.to_numeric(df[c], errors="coerce").fillna(0.0).values
    nomes = feature_names_funil_cal_lags(incluir_mes=incluir_mes)
    X = np.hstack([X_cal[:, :-1], X_lags, X_cal[:, -1:]])
    return X, nomes


def _validar_modelo_funil_producao(
    modelo: Optional[Dict[str, Any]],
    incluir_mes: bool = True,
) -> Tuple[bool, str]:
    if not modelo or not isinstance(modelo, dict):
        return False, "FUNIL_MODELO_PRODUCAO ausente — rode treinar_modelo_funil.py"
    if str(modelo.get("schema_version") or "") != FUNIL_MODELO_SCHEMA:
        return False, (
            f"schema incompatível: {modelo.get('schema_version')!r} "
            f"(esperado {FUNIL_MODELO_SCHEMA!r})"
        )
    esperados = feature_names_funil_cal_lags(incluir_mes=incluir_mes)
    nomes = list(modelo.get("feature_names") or [])
    if nomes != esperados:
        return False, "feature_names fora da ordem/quantidade canônica"
    coefs = modelo.get("coefs") or {}
    for etapa in FUNIL_ETAPAS:
        c = coefs.get(etapa)
        if c is None or len(c) != len(esperados) + 1:
            return False, f"coeficientes inválidos para {etapa}"
    for chave in ("r2s", "r2s_medias", "medias", "totais_hist"):
        if chave not in modelo:
            return False, f"campo obrigatório ausente: {chave}"
    return True, ""


def carregar_coefs_funil_producao(
    incluir_mes: bool = True,
) -> Tuple[Dict[str, np.ndarray], Dict[str, Any]]:
    """Carrega coeficientes embutidos; levanta ValueError se inválidos."""
    ok, msg = _validar_modelo_funil_producao(FUNIL_MODELO_PRODUCAO, incluir_mes=incluir_mes)
    if not ok:
        raise ValueError(msg)
    modelo = FUNIL_MODELO_PRODUCAO or {}
    coefs = {
        e: np.asarray(modelo["coefs"][e], dtype=float)
        for e in FUNIL_ETAPAS
    }
    return coefs, modelo


def _efeitos_lags_from_dict(raw: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """Reconstrói DataFrames dos perfis gravados no artefato."""
    if not raw or not isinstance(raw, dict):
        return None
    out = dict(raw)
    perfis_raw = out.get("perfis") or {}
    perfis: Dict[str, pd.DataFrame] = {}
    for etapa, rows in perfis_raw.items():
        if isinstance(rows, pd.DataFrame):
            perfis[etapa] = rows
        elif isinstance(rows, list):
            perfis[etapa] = pd.DataFrame(rows)
        elif isinstance(rows, dict):
            perfis[etapa] = pd.DataFrame(rows)
    out["perfis"] = perfis
    return out


def _matriz_funil_explicativas(
    df: pd.DataFrame,
    incluir_mes: bool = True,
    lags: Tuple[int, ...] = FUNIL_LAGS,
    alvo: Optional[str] = None,
    etapas_lag: Optional[Tuple[str, ...]] = None,
    modelo_vendas_completo: bool = False,
    efeitos_cruzados: bool = True,
) -> Tuple[np.ndarray, List[str]]:
    """
    Calendário + 4 blocos semanais de todas as etapas + força de trabalho
    + 7 conversões em janelas 7/14/21/28d.

    Não usa níveis do próprio dia das outras etapas: isso seria informação que
    ainda não existe no instante da previsão e inflaria artificialmente o R².
    """
    X_cal = _matriz_explicativas(df, incluir_mes=incluir_mes)
    lag_etapas = etapas_lag if etapas_lag is not None else FUNIL_ETAPAS
    lag_cols = _cols_lag_funil(lags, alvo=alvo, etapas=lag_etapas)

    extra_cols: List[str] = list(lag_cols)

    if efeitos_cruzados or modelo_vendas_completo:
        if "forca_trabalho" in df.columns:
            extra_cols.append("forca_trabalho")
        # As 7 conversões × 4 janelas explicam todas as etapas, não só vendas.
        for c in cols_conversoes_funil():
            if c in df.columns:
                extra_cols.append(c)

    # remove duplicatas preservando ordem
    seen = set()
    extra_unique: List[str] = []
    for c in extra_cols:
        if c not in seen:
            seen.add(c)
            extra_unique.append(c)
    extra_cols = extra_unique

    X_extra = np.zeros((len(df), len(extra_cols)), dtype=float)
    for j, c in enumerate(extra_cols):
        if c in df.columns:
            X_extra[:, j] = pd.to_numeric(df[c], errors="coerce").fillna(0.0).values
    X = np.hstack([X_cal[:, :-1], X_extra, X_cal[:, -1:]])
    return X, extra_cols


def _pesos_temporais_treino(datas: pd.Series) -> np.ndarray:
    """Peso 1,00 / 0,55 / 0,30 conforme a idade dentro dos três anos."""
    dt = pd.to_datetime(datas, errors="coerce")
    if dt.notna().any():
        ref = dt.max()
        idade_dias = (ref - dt).dt.days.fillna(365 * FUNIL_ANOS_TREINO)
    else:
        idade_dias = pd.Series(np.zeros(len(datas)), index=datas.index)
    ano_idade = np.floor(np.maximum(idade_dias.to_numpy(dtype=float), 0.0) / 365.25).astype(int)
    pesos = np.asarray(
        [FUNIL_PESOS_ANUAIS[min(int(a), len(FUNIL_PESOS_ANUAIS) - 1)] for a in ano_idade],
        dtype=float,
    )
    return pesos


def _media_ponderada(valores: pd.Series, pesos: np.ndarray) -> float:
    v = pd.to_numeric(valores, errors="coerce").fillna(0.0).to_numpy(dtype=float)
    w = np.asarray(pesos, dtype=float)
    soma_w = float(w.sum())
    return float(np.dot(v, w) / soma_w) if soma_w > 1e-9 else 0.0


def _ajustar_ridge_coef_original(
    X: np.ndarray,
    y: np.ndarray,
    alpha: float = FUNIL_RIDGE_ALPHA,
    sample_weight: Optional[np.ndarray] = None,
) -> np.ndarray:
    """
    Ridge com padronização, devolvendo coeficientes na escala original de X.
    A última coluna de X é o intercepto e não recebe penalização.
    """
    if X.ndim != 2 or X.shape[0] == 0:
        return np.zeros(X.shape[1] if X.ndim == 2 else 0, dtype=float)
    Xv = np.asarray(X[:, :-1], dtype=float)
    w = (
        np.asarray(sample_weight, dtype=float)
        if sample_weight is not None
        else np.ones(len(Xv), dtype=float)
    )
    w = np.clip(w, 1e-9, None)
    w = w / float(w.mean())
    mu = np.average(Xv, axis=0, weights=w)
    var = np.average((Xv - mu) ** 2, axis=0, weights=w)
    sd = np.sqrt(var)
    sd = np.where(sd > 1e-9, sd, 1.0)
    Z = np.hstack([(Xv - mu) / sd, np.ones((len(Xv), 1), dtype=float)])
    raiz_w = np.sqrt(w)[:, None]
    Zw = Z * raiz_w
    yw = np.asarray(y, dtype=float) * raiz_w[:, 0]
    penal = np.eye(Z.shape[1], dtype=float) * max(float(alpha), 0.0)
    penal[-1, -1] = 0.0
    try:
        beta_z = np.linalg.solve(Zw.T @ Zw + penal, Zw.T @ yw)
    except np.linalg.LinAlgError:
        beta_z, *_ = np.linalg.lstsq(Zw.T @ Zw + penal, Zw.T @ yw, rcond=None)
    beta = beta_z[:-1] / sd
    intercepto = float(beta_z[-1] - np.dot(mu, beta))
    return np.concatenate([beta, np.array([intercepto])])


def treinar_regressao_funil(
    treino: pd.DataFrame,
    alvo: str,
    incluir_mes: bool = True,
    lags: Tuple[int, ...] = FUNIL_LAGS,
) -> np.ndarray:
    X, _ = _matriz_funil_explicativas(
        treino,
        incluir_mes=incluir_mes,
        lags=lags,
        alvo=alvo,
        modelo_vendas_completo=(alvo == "vendas"),
        efeitos_cruzados=True,
    )
    y = treino[alvo].astype(float).values
    pesos = _pesos_temporais_treino(treino["data"])
    return _ajustar_ridge_coef_original(X, y, sample_weight=pesos)


def _r2_funil(
    treino: pd.DataFrame,
    coef: np.ndarray,
    alvo: str,
    incluir_mes: bool = True,
    lags: Tuple[int, ...] = FUNIL_LAGS,
) -> float:
    X, _ = _matriz_funil_explicativas(
        treino,
        incluir_mes=incluir_mes,
        lags=lags,
        alvo=alvo,
        modelo_vendas_completo=(alvo == "vendas"),
        efeitos_cruzados=True,
    )
    y = treino[alvo].astype(float).values
    y_hat = X @ coef
    w = _pesos_temporais_treino(treino["data"])
    media_y = float(np.average(y, weights=w))
    ss_res = float(np.sum(w * (y - y_hat) ** 2))
    ss_tot = float(np.sum(w * (y - media_y) ** 2))
    if ss_tot <= 0:
        return 0.0
    return 1.0 - (ss_res / ss_tot)


def calcular_medias_funil(
    treino: pd.DataFrame,
    incluir_mes: bool = True,
    lags: Tuple[int, ...] = FUNIL_LAGS,
) -> Dict[str, Any]:
    """Médias sazonais ponderadas por recência + blocos + força de trabalho."""
    pesos = _pesos_temporais_treino(treino["data"])

    def medias_por_grupo(col_grupo: str, col_valor: str) -> Dict[Any, float]:
        resultado: Dict[Any, float] = {}
        for chave, indices in treino.groupby(col_grupo).groups.items():
            pos = treino.index.get_indexer(indices)
            pos = pos[pos >= 0]
            resultado[chave] = _media_ponderada(treino.loc[indices, col_valor], pesos[pos])
        return resultado

    forca_mu = (
        _media_ponderada(treino["forca_trabalho"], pesos)
        if "forca_trabalho" in treino.columns and len(treino)
        else 1.0
    )
    if abs(forca_mu) < 1e-9:
        forca_mu = 1e-9
    out: Dict[str, Any] = {
        "incluir_mes": incluir_mes,
        "lags": lags,
        "forca_mu": forca_mu,
        "etapas": {},
    }
    for etapa in FUNIL_ETAPAS:
        mu = _media_ponderada(treino[etapa], pesos) if len(treino) else 0.0
        if mu <= 0:
            mu = 1e-9
        etapa_info: Dict[str, Any] = {
            "mu": mu,
            "media_dia_semana": {
                k: float(v) for k, v in medias_por_grupo("dia_semana", etapa).items()
            },
            "media_dia_mes": {
                int(k): float(v) for k, v in medias_por_grupo("dia_mes", etapa).items()
            },
        }
        if incluir_mes:
            etapa_info["media_mes"] = {
                k: float(v) for k, v in medias_por_grupo("mes", etapa).items()
            }
        # lags de TODAS as etapas (efeitos cruzados na intensidade)
        lag_mu: Dict[str, float] = {}
        for e2 in FUNIL_ETAPAS:
            for ini_lag, fim_lag in FUNIL_LAG_BLOCOS:
                col = col_lag_bloco(e2, ini_lag, fim_lag)
                if col in treino.columns:
                    m = _media_ponderada(treino[col], pesos)
                    lag_mu[col] = m if m > 0 else mu
        etapa_info["lag_mu"] = lag_mu
        out["etapas"][etapa] = etapa_info
    return out


def _pred_sazonal_etapa(d: date, info: Dict[str, Any], incluir_mes: bool) -> float:
    mu = float(info["mu"])
    a = float(info["media_dia_semana"].get(DIAS_SEMANA_PT[d.weekday()], mu))
    b = float(info["media_dia_mes"].get(d.day, mu))
    if incluir_mes and info.get("media_mes"):
        c = float(info["media_mes"].get(MESES_PT[d.month], mu))
        return max((a * b * c) / (mu * mu), 0.0)
    return max((a * b) / mu, 0.0)


def _intensidade_lags_linha(row: pd.Series, info: Dict[str, Any]) -> float:
    """Fator multiplicativo pela intensidade dos lags (próprios + cruzados)."""
    lag_mu = info.get("lag_mu") or {}
    if not lag_mu:
        return 1.0
    ratios: List[float] = []
    for col, mu_l in lag_mu.items():
        v = float(row.get(col, 0.0) or 0.0)
        mu_l = float(mu_l) if float(mu_l) > 1e-9 else 1e-9
        ratios.append(v / mu_l)
    if not ratios:
        return 1.0
    ratios = [max(r, 0.05) for r in ratios]
    geo = float(np.exp(np.mean(np.log(ratios))))
    return float(np.clip(geo, 0.25, 3.0))


def _intensidade_forca_linha(row: pd.Series, medias: Dict[str, Any]) -> float:
    """Fator pela força de trabalho vs média histórica."""
    mu = float(medias.get("forca_mu") or 0.0)
    v = float(row.get("forca_trabalho", 0.0) or 0.0)
    fat = 1.0 + 0.35 * (v - mu)
    return float(np.clip(fat, 0.35, 2.5))


def _prever_linha_reg_funil(
    coef: np.ndarray,
    row_df: pd.DataFrame,
    incluir_mes: bool,
    lags: Tuple[int, ...],
    alvo: str,
    usar_cal_lags: bool = True,
) -> float:
    if usar_cal_lags:
        X, _ = _matriz_funil_cal_lags(row_df, incluir_mes=incluir_mes)
    else:
        X, _ = _matriz_funil_explicativas(
            row_df,
            incluir_mes=incluir_mes,
            lags=lags,
            alvo=alvo,
            modelo_vendas_completo=(alvo == "vendas"),
            efeitos_cruzados=True,
        )
    return float(max((X @ coef)[0], 0.0))


def _prever_linha_medias_funil(
    row: pd.Series,
    d: date,
    etapa: str,
    medias: Dict[str, Any],
    incluir_mes: bool,
) -> float:
    info = medias["etapas"][etapa]
    saz = _pred_sazonal_etapa(d, info, incluir_mes)
    intens = _intensidade_lags_linha(row, info)
    forca = _intensidade_forca_linha(row, medias)
    return max(saz * intens * forca, 0.0)


def _combinar_previsoes_funil(
    previsao_reg: float,
    previsao_sazonal: float,
    r2: float,
) -> float:
    """
    Ensemble robusto por etapa.

    A regressão contém lags/efeitos cruzados, mas pode superajustar com muitos
    regressores. Seu peso varia de 35% a 70%; a base sazonal sempre participa.
    """
    peso_reg = float(np.clip(0.35 + 0.35 * max(float(r2 or 0.0), 0.0), 0.35, 0.70))
    reg = max(float(previsao_reg or 0.0), 0.0)
    saz = max(float(previsao_sazonal or 0.0), 0.0)
    return peso_reg * reg + (1.0 - peso_reg) * saz


def _r2_medias_funil(
    treino: pd.DataFrame,
    medias: Dict[str, Any],
    alvo: str,
    incluir_mes: bool = True,
) -> float:
    if treino.empty or alvo not in treino.columns:
        return 0.0
    y = treino[alvo].astype(float).values
    y_hat = np.array([
        _prever_linha_medias_funil(row, row["data"], alvo, medias, incluir_mes)
        for _, row in treino.iterrows()
    ], dtype=float)
    w = _pesos_temporais_treino(treino["data"])
    media_y = float(np.average(y, weights=w))
    ss_res = float(np.sum(w * (y - y_hat) ** 2))
    ss_tot = float(np.sum(w * (y - media_y) ** 2))
    if ss_tot <= 0:
        return 0.0
    return 1.0 - (ss_res / ss_tot)


def _atualizar_lags_linha(cal: pd.DataFrame, i: int, lags: Tuple[int, ...]) -> None:
    """Atualiza os quatro blocos de volume anteriores à linha i."""
    for etapa in FUNIL_ETAPAS:
        for ini_lag, fim_lag in FUNIL_LAG_BLOCOS:
            j_ini = max(0, i - fim_lag)
            j_fim = max(0, i - ini_lag + 1)
            valor = float(cal.iloc[j_ini:j_fim][etapa].sum()) if j_fim > j_ini else 0.0
            cal.at[i, col_lag_bloco(etapa, ini_lag, fim_lag)] = valor


def _resumir_perfil_lag_vendas(df_p: pd.DataFrame, etapa: str) -> Dict[str, Any]:
    """
    Resume o perfil semanal (1–7, 8–14, 15–21, 22–28) sobre vendas.
    """
    df_p = df_p[df_p["lag"] >= 1].copy()
    if df_p.empty:
        return {
            "etapa": etapa,
            "label": FUNIL_LABELS.get(etapa, etapa),
            "lag_pico": 7,
            "efeito_pico": 0.0,
            "lag_meia_vida": 7,
            "efeito_lag1": 0.0,
            "efeito_acum": 0.0,
        }

    df_p = df_p.sort_values("lag").reset_index(drop=True)
    df_p["acumulado"] = df_p["efeito"].cumsum()

    i_peak = int(df_p["acumulado"].idxmax())
    lag_pico = int(df_p.loc[i_peak, "lag"])
    efeito_pico = float(df_p.loc[i_peak, "acumulado"])

    acum_final = float(df_p["acumulado"].iloc[-1])
    lag_meia = lag_pico
    if abs(acum_final) > 1e-9:
        alvo = 0.5 * acum_final
        lag_meia = int(df_p["lag"].iloc[-1])
        for _, r in df_p.iterrows():
            ac = float(r["acumulado"])
            if acum_final > 0 and ac >= alvo:
                lag_meia = int(r["lag"])
                break
            if acum_final < 0 and ac <= alvo:
                lag_meia = int(r["lag"])
                break
    else:
        lag_meia = max(7, lag_pico)

    efeito_lag1 = float(df_p.iloc[0]["efeito"]) if not df_p.empty else 0.0
    return {
        "etapa": etapa,
        "label": FUNIL_LABELS.get(etapa, etapa),
        "lag_pico": lag_pico,
        "efeito_pico": efeito_pico,
        "lag_meia_vida": lag_meia,
        "efeito_lag1": efeito_lag1,
        "efeito_acum": acum_final,
    }


def estimar_efeitos_lags_sobre_vendas(
    treino: pd.DataFrame,
    lags: Tuple[int, ...] = FUNIL_LAGS_PERFIL,
    incluir_mes: bool = True,
) -> Optional[Dict[str, Any]]:
    """
    Perfil de tempo até efeito nas vendas em quatro blocos semanais.

    Ridge por etapa: vendas ~ calendário + volumes 1–7, 8–14, 15–21 e 22–28d.
    """
    if treino.empty or float(treino["vendas"].sum()) <= 0 or len(treino) < 40:
        return None

    cal = treino.copy()
    for etapa in FUNIL_DRIVERS:
        if etapa not in cal.columns:
            return None
        for ini_lag, fim_lag in FUNIL_LAG_BLOCOS:
            col = col_lag_bloco(etapa, ini_lag, fim_lag)
            if col not in cal.columns:
                janela = fim_lag - ini_lag + 1
                cal[col] = (
                    cal[etapa].shift(ini_lag).rolling(janela, min_periods=1).sum().fillna(0.0)
                )

    y = cal["vendas"].astype(float).values
    pesos_perfil = _pesos_temporais_treino(cal["data"])
    perfis: Dict[str, pd.DataFrame] = {}
    resumo: List[Dict[str, Any]] = []
    r2s_etapa: List[float] = []

    for etapa in FUNIL_DRIVERS:
        lag_cols = [
            col_lag_bloco(etapa, ini_lag, fim_lag)
            for ini_lag, fim_lag in FUNIL_LAG_BLOCOS
        ]
        X_cal = _matriz_explicativas(cal, incluir_mes=incluir_mes)
        X_lag = np.zeros((len(cal), len(lag_cols)), dtype=float)
        for j, c in enumerate(lag_cols):
            X_lag[:, j] = pd.to_numeric(cal[c], errors="coerce").fillna(0.0).values
        X = np.hstack([X_cal[:, :-1], X_lag, X_cal[:, -1:]])

        coef = _ajustar_ridge_coef_original(X, y, sample_weight=pesos_perfil)
        n_cal = X_cal.shape[1] - 1
        coef_lags = coef[n_cal:n_cal + len(lag_cols)]

        efeitos = [float(coef_lags[j]) for j in range(len(lag_cols))]
        df_p = pd.DataFrame({
            "lag": [fim_lag for _ini_lag, fim_lag in FUNIL_LAG_BLOCOS],
            "inicio_lag": [ini_lag for ini_lag, _fim_lag in FUNIL_LAG_BLOCOS],
            "efeito": efeitos,
        })
        df_p["acumulado"] = df_p["efeito"].cumsum()
        perfis[etapa] = df_p
        resumo.append(_resumir_perfil_lag_vendas(df_p, etapa))

        y_hat = X @ coef
        media_y = float(np.average(y, weights=pesos_perfil))
        ss_res = float(np.sum(pesos_perfil * (y - y_hat) ** 2))
        ss_tot = float(np.sum(pesos_perfil * (y - media_y) ** 2))
        r2s_etapa.append((1.0 - ss_res / ss_tot) if ss_tot > 0 else 0.0)

    r2 = float(np.mean(r2s_etapa)) if r2s_etapa else 0.0

    return {
        "r2": r2,
        "r2s_etapa": {e: r2s_etapa[i] for i, e in enumerate(FUNIL_DRIVERS)},
        "lags": tuple(fim_lag for _ini_lag, fim_lag in FUNIL_LAG_BLOCOS),
        "perfis": perfis,
        "resumo": resumo,
    }


def projetar_funil_mes_atual(
    mapas: Dict[str, Dict[date, float]],
    hoje: Optional[date] = None,
    incluir_mes: bool = True,
    lags: Tuple[int, ...] = FUNIL_LAGS,
    meta_qtd_mes: float = 0.0,
) -> Optional[Dict[str, Any]]:
    """
    Projeção híbrida e reconciliada do funil completo.

    Usa coeficientes ElasticNet pré-treinados (cal + lags) embutidos em
    FUNIL_MODELO_PRODUCAO — sem retreino no open do relatório. Os mapas
    precisam cobrir apenas o mês atual + buffer de lags (SOQL produção).

    meta_qtd_mes: meta de vendas do mês (funil necessário e metas diárias).
    """
    hoje = hoje or date.today()
    try:
        coefs, modelo = carregar_coefs_funil_producao(incluir_mes=incluir_mes)
    except ValueError as e:
        raise RuntimeError(str(e)) from e

    inicio = date.fromisoformat(str(modelo["treino_inicio"]))
    fim_treino = date.fromisoformat(str(modelo["treino_fim"]))
    max_lag = max(lags) if lags else 0
    # Produção: calendário curto (buffer de lags + mês/semana corrente).
    inicio_cal = _sf_inicio_producao(hoje) - timedelta(
        days=max(0, max_lag - FUNIL_SOQL_BUFFER_LAGS)
    )
    inicio_cal = min(inicio_cal, date(hoje.year, hoje.month, 1) - timedelta(days=max_lag))
    # Se mapas têm histórico mais longo (SOQL painel), estende para lags completos.
    datas_mapa: List[date] = []
    for mp in (mapas or {}).values():
        if isinstance(mp, dict):
            datas_mapa.extend(d for d in mp.keys() if isinstance(d, date))
    if datas_mapa:
        inicio_mapa = min(datas_mapa)
        inicio_cal = min(inicio_cal, inicio_mapa)
        inicio_cal = min(inicio_cal, date(hoje.year, hoje.month, 1) - timedelta(days=max_lag))

    ano, mes = hoje.year, hoje.month
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    fim_mes = date(ano, mes, ultimo_dia)
    # Se a semana ISO atravessa o fim do mês, estende o calendário até o domingo
    # para completar a meta da semana com o projetado do mês seguinte.
    fim_semana = _domingo_semana_iso(hoje)
    fim_cal = max(fim_mes, fim_semana)

    cal = calendario_funil_diario(inicio_cal, fim_cal, mapas, lags=lags)
    cal = cal.reset_index(drop=True)
    cal["data"] = cal["data"].map(_as_date_funil)
    cal = cal.dropna(subset=["data"]).reset_index(drop=True)
    if cal.empty:
        return None
    # Mesmo sem volume MTD, segue com projeção (coeficientes + médias históricas).

    r2s = {e: float((modelo.get("r2s") or {}).get(e, 0.0)) for e in FUNIL_ETAPAS}
    r2s_medias = {
        e: float((modelo.get("r2s_medias") or {}).get(e, 0.0)) for e in FUNIL_ETAPAS
    }
    medias = modelo.get("medias") or {}
    totais_hist_treino = {
        e: float((modelo.get("totais_hist") or {}).get(e, 0.0)) for e in FUNIL_ETAPAS
    }
    taxas_cascata = _taxas_cascata_historicas(totais_hist_treino)
    # Treino stub só para fallback de dia da semana (já coberto por medias).
    treino = cal.iloc[0:0].copy()

    # Cópias para projeção (reg e médias)
    cal_reg = cal.copy()
    cal_med = cal.copy()

    dias_mes = [date(ano, mes, d) for d in range(1, ultimo_dia + 1)]
    # Projeta dias futuros do mês atual + dias da semana no mês seguinte (se houver)
    dias_a_projetar = []
    d_cursor = hoje + timedelta(days=1)
    while d_cursor <= fim_cal:
        dias_a_projetar.append(d_cursor)
        d_cursor += timedelta(days=1)
    idx_por_data = _indice_por_data_cal(cal_reg)

    for d in dias_a_projetar:
        i = idx_por_data.get(d)
        if i is None:
            continue
        _atualizar_lags_linha(cal_reg, i, lags)
        _atualizar_lags_linha(cal_med, i, lags)
        _atualizar_forca_trabalho_linha(cal_reg, i)
        _atualizar_forca_trabalho_linha(cal_med, i)

        # Todos os modelos fazem uma previsão própria.
        for etapa in FUNIL_ETAPAS:
            row_reg = cal_reg.loc[[i]]
            row_med = cal_med.loc[[i]]
            pred_reg_puro = _prever_linha_reg_funil(
                coefs[etapa], row_reg, incluir_mes, lags, etapa, usar_cal_lags=True
            )
            pred_m = _prever_linha_medias_funil(
                row_med.iloc[0], d, etapa, medias, incluir_mes
            )
            pred_r = _combinar_previsoes_funil(
                pred_reg_puro, pred_m, r2s.get(etapa, 0.0)
            )
            cal_reg.at[i, etapa] = pred_r
            cal_med.at[i, etapa] = pred_m
            if 0 in lags:
                cal_reg.at[i, f"{etapa}_lag0"] = pred_r
                cal_med.at[i, f"{etapa}_lag0"] = pred_m

        # Reconciliação conjunta: modelos próprios + coerência de conversões.
        _reconciliar_linha_cal_funil(cal_reg, i, taxas_cascata, r2s, lags)
        _reconciliar_linha_cal_funil(cal_med, i, taxas_cascata, r2s_medias, lags)
        _atualizar_conversoes_linha(cal_reg, i)
        _atualizar_conversoes_linha(cal_med, i)

    dias_futuros_mes = [d for d in dias_mes if d > hoje]
    _garantir_previsoes_futuras_funil(
        cal_reg, cal_med, idx_por_data, dias_futuros_mes, treino, medias, coefs, incluir_mes, lags
    )
    # Fallbacks podem preencher etapas isoladas — reconcilia novamente o sistema.
    for d in dias_a_projetar:
        i = idx_por_data.get(d)
        if i is None:
            continue
        _reconciliar_linha_cal_funil(cal_reg, i, taxas_cascata, r2s, lags)
        _reconciliar_linha_cal_funil(cal_med, i, taxas_cascata, r2s_medias, lags)

    # Projeção diária reconciliada (passado contrafactual + futuro).
    proj_dia_reg: Dict[date, Dict[str, float]] = {}
    proj_dia_med: Dict[date, Dict[str, float]] = {}
    for d in dias_mes:
        i = idx_por_data.get(d)
        if i is None:
            continue
        if d <= hoje:
            bruto_r: Dict[str, float] = {}
            bruto_m: Dict[str, float] = {}
            for e in FUNIL_ETAPAS:
                reg_puro = _prever_linha_reg_funil(
                    coefs[e], cal_reg.loc[[i]], incluir_mes, lags, e, usar_cal_lags=True
                )
                saz = _prever_linha_medias_funil(
                    cal_med.iloc[i], d, e, medias, incluir_mes
                )
                bruto_r[e] = _combinar_previsoes_funil(
                    reg_puro, saz, r2s.get(e, 0.0)
                )
                bruto_m[e] = saz
            proj_dia_reg[d] = _reconciliar_funil_hibrido(
                bruto_r, taxas_cascata, r2s
            )
            proj_dia_med[d] = _reconciliar_funil_hibrido(
                bruto_m, taxas_cascata, r2s_medias
            )
        else:
            proj_dia_reg[d] = {e: float(cal_reg.at[i, e]) for e in FUNIL_ETAPAS}
            proj_dia_med[d] = {e: float(cal_med.at[i, e]) for e in FUNIL_ETAPAS}

    # montar resultados por etapa
    resultados: Dict[str, Any] = {}
    for etapa in FUNIL_ETAPAS:
        mtd = 0.0
        proj_mtd_reg = 0.0
        proj_mtd_med = 0.0
        rest_reg = 0.0
        rest_med = 0.0
        diaria: List[Dict[str, Any]] = []
        for d in dias_mes:
            if d not in proj_dia_reg:
                continue
            real = float((mapas.get(etapa) or {}).get(d, 0.0)) if d <= hoje else None
            pred_reg_dia = float(proj_dia_reg[d].get(etapa, 0.0))
            pred_med_dia = float(proj_dia_med[d].get(etapa, 0.0))
            if d <= hoje:
                mtd += real or 0.0
                proj_mtd_reg += pred_reg_dia
                proj_mtd_med += pred_med_dia
            else:
                rest_reg += pred_reg_dia
                rest_med += pred_med_dia
            diaria.append({
                "dia": d.day,
                "realizado": real,
                "projetado_reg": pred_reg_dia,
                "projetado_med": pred_med_dia,
            })
        resultados[etapa] = {
            "mtd": mtd,
            "projetado_mtd_reg": proj_mtd_reg,
            "projetado_mtd_med": proj_mtd_med,
            "projetado_reg": mtd + rest_reg,
            "projetado_med": mtd + rest_med,
            "r2": r2s[etapa],
            "r2_medias": r2s_medias[etapa],
            "diaria": pd.DataFrame(diaria),
        }

    totais_mtd = {e: float((resultados.get(e) or {}).get("mtd", 0)) for e in FUNIL_ETAPAS}
    # O realizado é factual e não pode ser reduzido pela reconciliação.
    # Cada dia futuro já foi reconciliado; o mês = realizado MTD + futuro.
    totais_proj_mtd = {
        e: float((resultados.get(e) or {}).get("projetado_mtd_reg", 0)) for e in FUNIL_ETAPAS
    }
    totais_proj_mtd_med = {
        e: float((resultados.get(e) or {}).get("projetado_mtd_med", 0)) for e in FUNIL_ETAPAS
    }
    totais_proj = {
        e: max(
            totais_mtd[e],
            float((resultados.get(e) or {}).get("projetado_reg", 0)),
        )
        for e in FUNIL_ETAPAS
    }
    totais_proj_med = {
        e: max(
            totais_mtd[e],
            float((resultados.get(e) or {}).get("projetado_med", 0)),
        )
        for e in FUNIL_ETAPAS
    }
    totais_hist = dict(totais_hist_treino)

    # Conversões do mês (MTD / projetado) × histórico (treino: sem mês atual, até 1 ano)
    conversoes = {
        "realizado_mtd": calcular_conversoes_totais(totais_mtd),
        "projetado_mtd": calcular_conversoes_totais(totais_proj_mtd),
        "projetado_mes": calcular_conversoes_totais(totais_proj),
        "historico": calcular_conversoes_totais(totais_hist),
        "inicio_hist": inicio,
        "fim_hist": fim_treino,
    }

    # Funil necessário para bater a meta de vendas (via conversões históricas indicador→venda)
    meta_qtd = float(meta_qtd_mes or 0.0)
    gap_vendas = max(0.0, meta_qtd - totais_mtd.get("vendas", 0.0))
    taxas_hist_frac: Dict[str, Optional[float]] = {}
    for item in (conversoes["historico"].get("para_venda") or []):
        orig = str(item.get("origem") or "")
        taxa = item.get("taxa")
        taxas_hist_frac[orig] = (float(taxa) / 100.0) if taxa is not None else None

    funil_necessario: Dict[str, float] = {"vendas": meta_qtd if meta_qtd > 0 else totais_mtd.get("vendas", 0.0)}
    for etapa in FUNIL_DRIVERS:
        t = taxas_hist_frac.get(etapa)
        if meta_qtd > 0 and t is not None and t > 1e-9:
            funil_necessario[etapa] = float(math.ceil(meta_qtd / t))
        else:
            # fallback: proporção histórica vs vendas
            v_h = totais_hist.get("vendas", 0.0)
            i_h = totais_hist.get(etapa, 0.0)
            if meta_qtd > 0 and v_h > 0 and i_h > 0:
                funil_necessario[etapa] = float(math.ceil(meta_qtd * (i_h / v_h)))
            else:
                funil_necessario[etapa] = float(totais_proj.get(etapa, 0.0))

    # Gap restante por indicador (conversão histórica indicador→venda)
    gap_indicadores = _funil_gap_vendas(
        gap_vendas,
        taxas_hist_frac,
        totais_hist,
        gap_vendas_mes=gap_vendas,
        funil_necessario=funil_necessario,
        totais_mtd=totais_mtd,
    )

    # Distribuição do gap só nos dias restantes do mês atual
    dias_distrib = [d for d in dias_mes if d >= hoje]
    pesos_gap = _pesos_distribuicao_gap(dias_distrib, resultados)
    meta_vendas_dia = _meta_vendas_por_pesos(gap_vendas, [hoje], dias_distrib, pesos_gap)

    semana_iso = hoje.isocalendar()[:2]
    dias_semana_rest = []
    d_sem = hoje
    while d_sem <= fim_semana and d_sem.isocalendar()[:2] == semana_iso:
        dias_semana_rest.append(d_sem)
        d_sem += timedelta(days=1)
    dias_semana_mes = [d for d in dias_semana_rest if d <= fim_mes]
    dias_semana_prox = [d for d in dias_semana_rest if d > fim_mes]

    meta_vendas_semana_mes = _meta_vendas_por_pesos(
        gap_vendas, dias_semana_mes, dias_distrib, pesos_gap
    )
    funil_meta_dia = _funil_gap_vendas(
        meta_vendas_dia,
        taxas_hist_frac,
        totais_hist,
        gap_vendas_mes=gap_vendas,
        funil_necessario=funil_necessario,
        totais_mtd=totais_mtd,
    )
    funil_semana_mes = _funil_gap_vendas(
        meta_vendas_semana_mes,
        taxas_hist_frac,
        totais_hist,
        gap_vendas_mes=gap_vendas,
        funil_necessario=funil_necessario,
        totais_mtd=totais_mtd,
    )
    # Completa a semana com o projetado do mês seguinte (dias após fim do mês)
    funil_semana_prox: Dict[str, float] = {e: 0.0 for e in FUNIL_ETAPAS}
    for d in dias_semana_prox:
        i = idx_por_data.get(d)
        if i is None:
            continue
        for etapa in FUNIL_ETAPAS:
            funil_semana_prox[etapa] += float(cal_reg.at[i, etapa])
    funil_meta_semana = _somar_funis(funil_semana_mes, funil_semana_prox)
    meta_vendas_semana = float(funil_meta_semana.get("vendas", 0.0))

    razoes_real_vs_proj_mtd = _razoes_entre_funis(totais_mtd, totais_proj_mtd)
    razoes_proj_vs_nec = _razoes_entre_funis(totais_proj, funil_necessario)

    dias_futuros = [d for d in dias_mes if d > hoje]
    metas_diarias: Dict[str, Any] = {}
    for etapa in FUNIL_ETAPAS:
        df_d = (resultados.get(etapa) or {}).get("diaria", pd.DataFrame())
        if df_d is None or df_d.empty or not dias_futuros:
            metas_diarias[etapa] = {
                "gap": gap_indicadores.get(etapa, 0.0),
                "ritmo_reg": pd.DataFrame(columns=["dia", "qtd"]),
                "ritmo_med": pd.DataFrame(columns=["dia", "qtd"]),
                "realizado": pd.DataFrame(columns=["dia", "qtd"]),
            }
            continue
        fut = df_d[df_d["dia"] > hoje.day]
        pesos_reg = fut["projetado_reg"].astype(float).values if not fut.empty else np.ones(len(dias_futuros))
        pesos_med = fut["projetado_med"].astype(float).values if not fut.empty else np.ones(len(dias_futuros))
        # Alinha comprimento
        if len(pesos_reg) != len(dias_futuros):
            pesos_reg = np.ones(len(dias_futuros), dtype=float)
        if len(pesos_med) != len(dias_futuros):
            pesos_med = np.ones(len(dias_futuros), dtype=float)
        gap_e = float(gap_indicadores.get(etapa, 0.0))
        ritmo_reg, _ = _distribuir_gap_por_pesos(
            gap_e, pesos_reg, dias_futuros, totais_mtd.get(etapa, 0.0), hoje.day, arredondar_cima=True
        )
        ritmo_med, _ = _distribuir_gap_por_pesos(
            gap_e, pesos_med, dias_futuros, totais_mtd.get(etapa, 0.0), hoje.day, arredondar_cima=True
        )
        real_rows = []
        for d in dias_mes:
            if d > hoje:
                break
            real_rows.append({
                "dia": d.day,
                "qtd": float((mapas.get(etapa) or {}).get(d, 0.0)),
            })
        metas_diarias[etapa] = {
            "gap": gap_e,
            "necessario_mes": float(funil_necessario.get(etapa, 0.0)),
            "ritmo_reg": ritmo_reg,
            "ritmo_med": ritmo_med,
            "realizado": pd.DataFrame(real_rows),
        }

    # Perfil de lags embutido no modelo mensal (sem retreino no open).
    efeitos_lags = _efeitos_lags_from_dict(modelo.get("efeitos_lags_vendas"))

    return {
        "hoje": hoje,
        "inicio_treino": inicio,
        "fim_treino": fim_treino,
        "incluir_mes": incluir_mes,
        "lags": lags,
        "ultimo_dia": ultimo_dia,
        "fim_mes": fim_mes,
        "fim_semana": fim_semana,
        "meta_qtd_mes": meta_qtd,
        "gap_vendas": gap_vendas,
        "r2s": r2s,
        "r2s_medias": r2s_medias,
        "medias": medias,
        "etapas": resultados,
        "totais_mtd": totais_mtd,
        "totais_proj_mtd": totais_proj_mtd,
        "totais_proj": totais_proj,
        "totais_hist": totais_hist,
        "funil_necessario": funil_necessario,
        "funil_meta_dia": funil_meta_dia,
        "funil_meta_semana": funil_meta_semana,
        "funil_semana_mes": funil_semana_mes,
        "funil_semana_prox": funil_semana_prox,
        "meta_vendas_dia": meta_vendas_dia,
        "meta_vendas_semana": meta_vendas_semana,
        "meta_vendas_semana_mes": meta_vendas_semana_mes,
        "dias_semana_prox": dias_semana_prox,
        "razoes_real_vs_proj_mtd": razoes_real_vs_proj_mtd,
        "razoes_proj_vs_nec": razoes_proj_vs_nec,
        "conversoes": conversoes,
        "metas_diarias": metas_diarias,
        "efeitos_lags_vendas": efeitos_lags,
        "modelo_schema": str(modelo.get("schema_version") or FUNIL_MODELO_SCHEMA),
        "modelo_treinado_em": str(modelo.get("treinado_em") or ""),
        "taxas_cascata": {
            f"{a}->{b}": float(taxas_cascata.get((a, b), 0.0))
            for a, b in FUNIL_PARES_ETAPA
        },
    }


def _fmt_taxa_pct(taxa: Optional[float]) -> str:
    if taxa is None:
        return "—"
    return fmt_pct_valor(float(taxa))


def _fmt_razao(r: Optional[float]) -> str:
    if r is None:
        return "—"
    return f"{fmt_num(float(r))}×"


def _render_razoes_funil(
    razoes: Dict[str, Optional[float]],
    titulo: str,
    caption: str = "",
) -> None:
    """Cards com razão entre dois funis por etapa."""
    st.markdown(f"###### {titulo}")
    if caption:
        st.caption(caption)
    _render_kpi_cards([
        {
            "lbl": FUNIL_LABELS.get(etapa, etapa),
            "val": _fmt_razao(razoes.get(etapa)),
            "sub": (
                "acima" if (razoes.get(etapa) or 0) > 1.02
                else ("abaixo" if (razoes.get(etapa) is not None and float(razoes.get(etapa)) < 0.98) else "no ritmo")
            ),
        }
        for etapa in FUNIL_ETAPAS
    ])


def _plot_funil_etapa_comparativo(etapa: str, df: pd.DataFrame, ultimo_dia: int, dia_hoje: int) -> None:
    label = FUNIL_LABELS.get(etapa, etapa)
    st.markdown(f"##### {label}: projetado × realizado")
    if df is None or df.empty or "dia" not in df.columns:
        return
    fig = go.Figure()
    if "realizado" in df.columns:
        real = df.dropna(subset=["realizado"])
        if not real.empty:
            fig.add_trace(
                go.Scatter(
                    x=real["dia"], y=real["realizado"],
                    mode="lines+markers+text",
                    name="Realizado",
                    text=[fmt_qtd(float(v)) for v in real["realizado"]],
                    textposition="top center",
                    textfont=dict(size=10, color=COR_AZUL_ESC, family="Inter"),
                    line=dict(color=COR_AZUL_ESC, width=3),
                    marker=dict(size=7, color=COR_AZUL_ESC),
                )
            )
    if "projetado_reg" in df.columns:
        fig.add_trace(
            go.Scatter(
                x=df["dia"], y=df["projetado_reg"],
                mode="lines+markers+text",
                name="Projetado (híbrido reconciliado)",
                text=[fmt_qtd(float(v)) for v in df["projetado_reg"]],
                textposition="bottom center",
                textfont=dict(size=10, color=COR_VERMELHO, family="Inter"),
                line=dict(color=COR_VERMELHO, width=3, dash="dash"),
                marker=dict(size=7, color=COR_VERMELHO),
            )
        )
    if "projetado_med" in df.columns:
        fig.add_trace(
            go.Scatter(
                x=df["dia"], y=df["projetado_med"],
                mode="lines+markers+text",
                name="Projetado (médias)",
                text=[fmt_qtd(float(v)) for v in df["projetado_med"]],
                textposition="top center",
                textfont=dict(size=10, color="#0f766e", family="Inter"),
                line=dict(color="#0f766e", width=3, dash="dot"),
                marker=dict(size=7, symbol="diamond", color="#0f766e"),
            )
        )
    if not fig.data:
        return
    fig.add_vline(
        x=dia_hoje, line_width=1, line_dash="dot", line_color="#64748b",
        annotation_text="Hoje", annotation_position="top",
        annotation_font=dict(color=COR_TEXTO_PRETO, size=11, family="Inter"),
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=11),
        ),
        hovermode="x unified",
        height=380,
    )
    fig.update_xaxes(
        title_text="Dia do mês",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
        range=[0.5, ultimo_dia + 0.5],
    )
    fig.update_yaxes(
        title_text="Qtd. no dia",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.plotly_chart(
        fig, use_container_width=True, config={"displayModeBar": False},
        key=_plotly_key("funil_etapa_cmp", etapa),
    )


ALIASES_DATA_VISITA = [
    "Data da visita", "Data da Visita", "Data visita", "Data Visita",
    "Activity Date", "Data da Atividade", "Data do agendamento",
    "Data Agendamento", "Start Date Time", "Data/Hora",
]


def _montar_df_comparativo_mtd_parcial(
    df: pd.DataFrame,
    col_data: str,
    dia_atual: int,
    col_qtd: Optional[str] = None,
) -> pd.DataFrame:
    """Soma eventos do dia 1 ao dia_atual de cada mês (janela MTD parcial)."""
    if df is None or df.empty or not col_data or col_data not in df.columns:
        return pd.DataFrame()
    base = df.copy()
    base["_dt"] = parse_data_serie(base[col_data])
    base = base.dropna(subset=["_dt"])
    if base.empty:
        return pd.DataFrame()
    base = base.loc[base["_dt"].dt.day <= int(dia_atual)]
    if col_qtd and col_qtd in base.columns:
        base["_qtd"] = pd.to_numeric(base[col_qtd], errors="coerce").fillna(0.0)
    else:
        base["_qtd"] = 1.0
    base["_ano_c"] = base["_dt"].dt.year
    base["_mes_c"] = base["_dt"].dt.month
    df_comp = base.groupby(["_ano_c", "_mes_c"], as_index=False).agg(QTD=("_qtd", "sum"))
    df_comp = df_comp.sort_values(["_ano_c", "_mes_c"])
    df_comp["Periodo"] = (
        df_comp["_mes_c"].astype(str).str.zfill(2) + "/" + df_comp["_ano_c"].astype(str)
    )
    df_comp["QTD_Formatado"] = df_comp["QTD"].apply(lambda x: fmt_qtd(x))
    return df_comp


def _plot_comparativo_mtd_qtd_linha(
    df_comp: pd.DataFrame,
    titulo: str,
    cor: str = COR_AZUL_ESC,
    chart_key: str = "",
) -> None:
    """Gráfico de linha: quantidade acumulada MTD parcial por período."""
    st.markdown(f"##### {titulo}")
    if df_comp is None or df_comp.empty:
        st.info(f"Sem dados para {titulo}.")
        return
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=df_comp["Periodo"],
            y=df_comp["QTD"],
            mode="lines+markers+text",
            name="Quantidade",
            line=dict(color=cor, width=3),
            marker=dict(size=8, color=cor),
            text=df_comp["QTD_Formatado"],
            textposition="top center",
            textfont=dict(color=cor, size=11, family="Inter"),
        )
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        hovermode="x unified",
        height=380,
    )
    fig.update_xaxes(
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
    )
    fig.update_yaxes(
        title_text="Quantidade",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226, 232, 240, 0.5)",
    )
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
        key=chart_key or None,
    )


def _plot_comparativo_mtd_vendas_vgv(
    df_vendas: pd.DataFrame,
    col_contrato: str,
    dia_atual: int,
) -> None:
    """Comparativo MTD parcial de vendas (QTD + VGV), dia 1 ao dia atual."""
    df_grafico = df_vendas.copy()
    df_grafico["Data_Contrato_DT"] = parse_data_serie(df_grafico[col_contrato])
    df_grafico = df_grafico.dropna(subset=["Data_Contrato_DT"])
    if df_grafico.empty:
        st.info("Não há dados de vendas no período acumulado de eficiência para exibir.")
        return
    df_parcial = df_grafico.loc[df_grafico["Data_Contrato_DT"].dt.day <= dia_atual].copy()
    df_parcial["_ano_c"] = df_parcial["Data_Contrato_DT"].dt.year
    df_parcial["_mes_c"] = df_parcial["Data_Contrato_DT"].dt.month
    df_comp = df_parcial.groupby(["_ano_c", "_mes_c"], as_index=False).agg(
        QTD=("_qtd_venda", "sum"),
        VGV=("_vgv_venda", "sum"),
    ).sort_values(["_ano_c", "_mes_c"])
    df_comp["Periodo"] = (
        df_comp["_mes_c"].astype(str).str.zfill(2) + "/" + df_comp["_ano_c"].astype(str)
    )
    df_comp["VGV_Formatado"] = df_comp["VGV"].apply(lambda x: fmt_br_milhoes(x))
    df_comp["QTD_Formatado"] = df_comp["QTD"].apply(lambda x: fmt_qtd(x))

    fig_linha = make_subplots(specs=[[{"secondary_y": True}]])
    fig_linha.add_trace(
        go.Scatter(
            x=df_comp["Periodo"],
            y=df_comp["QTD"],
            mode="lines+markers+text",
            name="QTD Vendas",
            line=dict(color=COR_AZUL_ESC, width=3),
            marker=dict(size=8, color=COR_AZUL_ESC),
            text=df_comp["QTD_Formatado"],
            textposition="top center",
            textfont=dict(color=COR_AZUL_ESC, size=11, family="Inter"),
        ),
        secondary_y=False,
    )
    fig_linha.add_trace(
        go.Scatter(
            x=df_comp["Periodo"],
            y=df_comp["VGV"],
            mode="lines+markers+text",
            name="VGV Real",
            line=dict(color=COR_VERMELHO, width=3),
            marker=dict(size=8, color=COR_VERMELHO),
            text=df_comp["VGV_Formatado"],
            textposition="bottom center",
            textfont=dict(color=COR_VERMELHO, size=11, family="Inter"),
        ),
        secondary_y=True,
    )
    fig_linha.update_layout(
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        hovermode="x unified",
    )
    fig_linha.update_xaxes(
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
    )
    fig_linha.update_yaxes(
        title_text="Quantidade (Vendas)",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        secondary_y=False,
        showgrid=False,
    )
    fig_linha.update_yaxes(
        title_text="VGV Real (R$)",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        secondary_y=True,
        showgrid=True,
        gridcolor="rgba(226, 232, 240, 0.5)",
    )
    st.plotly_chart(fig_linha, use_container_width=True, config={"displayModeBar": False})


def _comp_mtd_qtd_etapa(
    df: pd.DataFrame,
    col_data: str,
    dia_atual: int,
    col_qtd: Optional[str] = None,
) -> pd.DataFrame:
    """Série MTD parcial por mês para uma etapa do funil."""
    return _montar_df_comparativo_mtd_parcial(df, col_data, dia_atual, col_qtd=col_qtd)


def _montar_df_mtd_funil_agregado(
    df_vendas: pd.DataFrame,
    col_contrato: str,
    df_ag: pd.DataFrame,
    df_pas: pd.DataFrame,
    df_pas_aprov: pd.DataFrame,
    dia_atual: int,
) -> pd.DataFrame:
    """
    Por mês (MTD parcial): quantidades por etapa, soma simples, soma ponderada
    e razões vendas ÷ soma.
    """
    col_ag = achar_coluna(df_ag, ALIASES_DATA_CRIACAO)
    col_vis = achar_coluna(df_ag, ALIASES_DATA_VISITA)
    col_envio = achar_coluna_primeiro_envio_analise(df_pas)
    col_safi = achar_coluna_aprovacao_safi(df_pas_aprov)
    col_qtd_ven = "_qtd_venda" if (df_vendas is not None and "_qtd_venda" in df_vendas.columns) else None

    etapas_cfg: List[Tuple[str, pd.DataFrame, Optional[str], Optional[str]]] = [
        ("agendamentos", df_ag, col_ag, None),
        ("visitas", df_ag, col_vis, None),
        ("pastas", df_pas, col_envio, None),
        ("pastas_aprovadas", df_pas_aprov, col_safi, None),
        ("vendas", df_vendas, col_contrato or None, col_qtd_ven),
    ]

    base: Optional[pd.DataFrame] = None
    chaves = ["Periodo", "_ano_c", "_mes_c"]
    for chave, dff, col_d, col_qtd in etapas_cfg:
        if dff is None or dff.empty or not col_d:
            continue
        comp = _comp_mtd_qtd_etapa(dff, col_d, dia_atual, col_qtd=col_qtd)
        if comp.empty:
            continue
        parte = comp.rename(columns={"QTD": chave})[chaves + [chave]]
        base = parte if base is None else base.merge(parte, on=chaves, how="outer")

    if base is None or base.empty:
        return pd.DataFrame()

    for etapa in FUNIL_ETAPAS:
        if etapa not in base.columns:
            base[etapa] = 0.0
        base[etapa] = pd.to_numeric(base[etapa], errors="coerce").fillna(0.0)

    base["soma_simples"] = sum(base[e] for e in FUNIL_ETAPAS)
    base["soma_ponderada"] = sum(
        float(PESOS_FUNIL_MTD[e]) * base[e] for e in FUNIL_ETAPAS
    )
    base["ratio_ponderada"] = np.where(
        base["soma_ponderada"] > 0, base["vendas"] / base["soma_ponderada"], np.nan,
    )
    base["ratio_simples"] = np.where(
        base["soma_simples"] > 0, base["vendas"] / base["soma_simples"], np.nan,
    )
    return base.sort_values(["_ano_c", "_mes_c"]).reset_index(drop=True)


def _plot_mtd_ratio_vendas_funil(df: pd.DataFrame) -> None:
    """Vendas ÷ soma ponderada e vendas ÷ soma simples por período."""
    st.markdown("##### Vendas ÷ funil (MTD parcial)")
    st.caption(
        "Pesos: agendamentos 1 · visitas 2 · pastas 3 · pastas aprovadas 4 · vendas 5"
    )
    if df is None or df.empty:
        st.info("Sem dados para o comparativo de razões.")
        return
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df["Periodo"],
        y=df["ratio_ponderada"],
        mode="lines+markers+text",
        name="Vendas ÷ soma ponderada",
        line=dict(color=COR_AZUL_ESC, width=3),
        marker=dict(size=8, color=COR_AZUL_ESC),
        text=[fmt_num(v) for v in df["ratio_ponderada"]],
        textposition="top center",
        textfont=dict(size=10, color=COR_AZUL_ESC, family="Inter"),
    ))
    fig.add_trace(go.Scatter(
        x=df["Periodo"],
        y=df["ratio_simples"],
        mode="lines+markers+text",
        name="Vendas ÷ soma simples",
        line=dict(color="#0f766e", width=3, dash="dash"),
        marker=dict(size=8, color="#0f766e", symbol="diamond"),
        text=[fmt_num(v) for v in df["ratio_simples"]],
        textposition="bottom center",
        textfont=dict(size=10, color="#0f766e", family="Inter"),
    ))
    fig.update_layout(
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5),
        hovermode="x unified",
        height=400,
    )
    fig.update_yaxes(
        title_text="Razão (vendas ÷ soma)",
        showgrid=True,
        gridcolor="rgba(226, 232, 240, 0.5)",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False}, key="comp_mtd_ratio_funil")


def _plot_mtd_volumes_funil(df: pd.DataFrame) -> None:
    """Vendas, soma simples e soma ponderada por período."""
    st.markdown("##### Volumes MTD parcial — vendas, soma simples e ponderada")
    if df is None or df.empty:
        st.info("Sem dados para o comparativo de volumes.")
        return
    fig = go.Figure()
    series = [
        ("vendas", "Vendas", COR_VERMELHO),
        ("soma_simples", "Soma simples do funil", COR_AZUL_ESC),
        ("soma_ponderada", "Soma ponderada do funil", "#0f766e"),
    ]
    for col, nome, cor in series:
        fig.add_trace(go.Scatter(
            x=df["Periodo"],
            y=df[col],
            mode="lines+markers+text",
            name=nome,
            line=dict(color=cor, width=3),
            marker=dict(size=8, color=cor),
            text=[fmt_qtd(v) for v in df[col]],
            textposition="top center",
            textfont=dict(size=10, color=cor, family="Inter"),
        ))
    fig.update_layout(
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5),
        hovermode="x unified",
        height=400,
    )
    fig.update_yaxes(title_text="Quantidade", showgrid=True, gridcolor="rgba(226, 232, 240, 0.5)")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False}, key="comp_mtd_vol_funil")


def render_comparativos_mtd_funil(
    df_vendas: pd.DataFrame,
    col_contrato_gerado: Optional[str],
) -> None:
    """Comparativos MTD parciais (dia 1 ao dia atual): vendas + etapas do funil."""
    dia_atual = datetime.now().day
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader(f"Comparativos MTD (Dia 01 ao Dia {dia_atual:02d} do Mês)")
    st.caption(
        "Cada ponto soma apenas eventos do dia 1 até o dia atual de cada mês — "
        "comparação justa de ritmo entre períodos."
    )

    if col_contrato_gerado:
        st.markdown(f"##### Vendas")
        _plot_comparativo_mtd_vendas_vgv(df_vendas, col_contrato_gerado, dia_atual)
    else:
        st.warning("Coluna 'Contrato gerado em' não encontrada — comparativo de vendas indisponível.")

    try:
        pacote = carregar_funil_historico_painel_sf()
        df_ag = deduplicar_agendamentos_funil(_coalesce_df(pacote.get("agendamentos")))
        df_pas_raw = _coalesce_df(pacote.get("pastas"))
        df_pas = deduplicar_pastas_funil(df_pas_raw)
        df_pas_aprov = deduplicar_pastas_aprovadas_funil(df_pas_raw)
        t_sf = float((pacote.get("timings") or {}).get("total_s", 0.0))
        if t_sf:
            st.caption(f"Funil histórico SF ({PAINEL_MESES_VENDAS}m) carregado em {fmt_num(t_sf)}s")

        col_ag = achar_coluna(df_ag, ALIASES_DATA_CRIACAO)
        col_vis = achar_coluna(df_ag, ALIASES_DATA_VISITA)
        col_envio = achar_coluna_primeiro_envio_analise(df_pas)
        col_safi = achar_coluna_aprovacao_safi(df_pas_aprov)

        comparativos = [
            ("Agendamentos", df_ag, col_ag, COR_AZUL_ESC),
            ("Visitas", df_ag, col_vis, "#1e60b3"),
            ("Pastas", df_pas, col_envio, "#0f766e"),
            ("Pastas aprovadas", df_pas_aprov, col_safi, "#b45309"),
        ]
        for titulo, df_etapa, col_data, cor in comparativos:
            if not col_data:
                st.markdown(f"##### {titulo}")
                st.info(f"Coluna de data não encontrada para {titulo}.")
                continue
            df_comp = _montar_df_comparativo_mtd_parcial(df_etapa, col_data, dia_atual)
            _plot_comparativo_mtd_qtd_linha(
                df_comp,
                titulo,
                cor=cor,
                chart_key=f"comp_mtd_{titulo.lower().replace(' ', '_')}",
            )

        if col_contrato_gerado:
            df_funil_agg = _montar_df_mtd_funil_agregado(
                df_vendas, col_contrato_gerado, df_ag, df_pas, df_pas_aprov, dia_atual,
            )
            _plot_mtd_ratio_vendas_funil(df_funil_agg)
            _plot_mtd_volumes_funil(df_funil_agg)
    except Exception as exc:
        st.warning(f"Não foi possível carregar comparativos do funil: {exc}")


def _render_kpi_cards(items: List[Dict[str, str]]) -> None:
    """
    Cards KPI em colunas Streamlit (evita HTML indentado virar bloco de código no markdown).
    items: [{"lbl": ..., "val": ..., "sub": ...}, ...]
    """
    if not items:
        return
    n = len(items)
    cols = st.columns(n)
    for col, it in zip(cols, items):
        lbl = html.escape(str(it.get("lbl", "")))
        val = html.escape(str(it.get("val", "")))
        sub = html.escape(str(it.get("sub", "")))
        with col:
            st.markdown(
                (
                    f'<div class="vel-kpi" style="width:100%;box-sizing:border-box;">'
                    f'<div class="lbl">{lbl}</div>'
                    f'<div class="val">{val}</div>'
                    f'<div class="lbl" style="margin-top:6px;opacity:0.75;">{sub}</div>'
                    f"</div>"
                ),
                unsafe_allow_html=True,
            )


def _render_conversoes_funil(conversoes: Dict[str, Any]) -> None:
    """Conversões do mês atual × histórico (sem mês atual, até 1 ano)."""
    if not conversoes:
        return
    mes = conversoes.get("realizado_mtd") or {}
    hist = conversoes.get("historico") or {}
    ini = conversoes.get("inicio_hist")
    fim = conversoes.get("fim_hist")
    periodo = ""
    if ini and fim:
        periodo = f"{ini.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}"

    st.markdown("##### Referência: conversões MTD × histórico")
    st.caption(f"Mês atual (MTD) × histórico ({periodo or 'janela de treino'})")
    pares_m = mes.get("etapa_a_etapa") or []
    pares_h = {r["label"]: r for r in (hist.get("etapa_a_etapa") or [])}
    fins_m = mes.get("para_venda") or []
    fins_h = {r["label"]: r for r in (hist.get("para_venda") or [])}

    labels_e = [str(r.get("label", "")) for r in pares_m]
    y_mes_e = [float(r["taxa"]) if r.get("taxa") is not None else 0.0 for r in pares_m]
    y_hist_e = [float((pares_h.get(lbl) or {}).get("taxa") or 0.0) for lbl in labels_e]
    labels_f = [str(r.get("label", "")) for r in fins_m]
    y_mes_f = [float(r["taxa"]) if r.get("taxa") is not None else 0.0 for r in fins_m]
    y_hist_f = [float((fins_h.get(lbl) or {}).get("taxa") or 0.0) for lbl in labels_f]

    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=("Etapa → etapa", "Indicador → venda"),
        horizontal_spacing=0.08,
    )
    fig.add_trace(
        go.Bar(
            name="Mês atual (MTD)", x=labels_e, y=y_mes_e,
            text=[_fmt_taxa_pct(v) for v in y_mes_e], textposition="outside",
            marker_color=COR_AZUL_ESC, textfont=dict(color=COR_TEXTO_PRETO, size=10),
            legendgroup="mes",
        ),
        row=1, col=1,
    )
    fig.add_trace(
        go.Bar(
            name="Histórico (sem mês atual)", x=labels_e, y=y_hist_e,
            text=[_fmt_taxa_pct(v) for v in y_hist_e], textposition="outside",
            marker_color="#0f766e", textfont=dict(color=COR_TEXTO_PRETO, size=10),
            legendgroup="hist",
        ),
        row=1, col=1,
    )
    fig.add_trace(
        go.Bar(
            name="Mês atual (MTD)", x=labels_f, y=y_mes_f,
            text=[_fmt_taxa_pct(v) for v in y_mes_f], textposition="outside",
            marker_color=COR_AZUL_ESC, textfont=dict(color=COR_TEXTO_PRETO, size=10),
            legendgroup="mes", showlegend=False,
        ),
        row=1, col=2,
    )
    fig.add_trace(
        go.Bar(
            name="Histórico (sem mês atual)", x=labels_f, y=y_hist_f,
            text=[_fmt_taxa_pct(v) for v in y_hist_f], textposition="outside",
            marker_color="#0f766e", textfont=dict(color=COR_TEXTO_PRETO, size=10),
            legendgroup="hist", showlegend=False,
        ),
        row=1, col=2,
    )
    fig.update_layout(
        barmode="group",
        margin=dict(l=20, r=20, t=50, b=80),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.12, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        height=420,
    )
    fig.update_yaxes(
        title_text="Taxa (%)",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    fig.update_xaxes(tickfont=dict(color=COR_TEXTO_PRETO, family="Inter", size=10), tickangle=-25)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def _plot_conversoes_barras_simples(
    conv: Dict[str, Any],
    titulo: str = "",
    chart_key: str = "",
) -> None:
    """Gráfico de barras das conversões de um funil (sem cards)."""
    pares = conv.get("etapa_a_etapa") or []
    diretas = conv.get("para_venda") or []
    if not pares and not diretas:
        return
    if titulo:
        st.markdown(f"###### {titulo}")

    ncols = (1 if pares else 0) + (1 if diretas else 0)
    if ncols == 0:
        return
    fig = make_subplots(
        rows=1, cols=ncols,
        subplot_titles=tuple(
            t for t, ok in (("Etapa → etapa", bool(pares)), ("Indicador → venda", bool(diretas))) if ok
        ),
        horizontal_spacing=0.08,
    )
    col_idx = 1
    if pares:
        labels = [str(r.get("label", "")) for r in pares]
        y = [float(r["taxa"]) if r.get("taxa") is not None else 0.0 for r in pares]
        fig.add_trace(
            go.Bar(
                x=labels, y=y,
                text=[_fmt_taxa_pct(v) for v in y],
                textposition="outside",
                marker_color=COR_AZUL_ESC,
                textfont=dict(color=COR_TEXTO_PRETO, size=10),
                showlegend=False,
            ),
            row=1, col=col_idx,
        )
        col_idx += 1
    if diretas:
        labels = [str(r.get("label", "")) for r in diretas]
        y = [float(r["taxa"]) if r.get("taxa") is not None else 0.0 for r in diretas]
        fig.add_trace(
            go.Bar(
                x=labels, y=y,
                text=[_fmt_taxa_pct(v) for v in y],
                textposition="outside",
                marker_color="#0f766e",
                textfont=dict(color=COR_TEXTO_PRETO, size=10),
                showlegend=False,
            ),
            row=1, col=col_idx,
        )
    fig.update_layout(
        margin=dict(l=20, r=20, t=50, b=80),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        height=380,
    )
    fig.update_yaxes(
        title_text="Taxa (%)",
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    fig.update_xaxes(tickangle=-25)
    st.plotly_chart(
        fig, use_container_width=True, config={"displayModeBar": False},
        key=chart_key or _plotly_key("conv_barras", titulo),
    )


def _plot_conversoes_par_barras(
    conv_a: Dict[str, Any],
    conv_b: Dict[str, Any],
    label_a: str,
    label_b: str,
    chart_key: str = "",
) -> None:
    """Comparativo em barras agrupadas entre dois funis."""
    pares_a = conv_a.get("etapa_a_etapa") or []
    pares_b = {r["label"]: r for r in (conv_b.get("etapa_a_etapa") or [])}
    fins_a = conv_a.get("para_venda") or []
    fins_b = {r["label"]: r for r in (conv_b.get("para_venda") or [])}

    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=("Etapa → etapa", "Indicador → venda"),
        horizontal_spacing=0.08,
    )
    if pares_a:
        labels = [str(r.get("label", "")) for r in pares_a]
        ya = [float(r["taxa"]) if r.get("taxa") is not None else 0.0 for r in pares_a]
        yb = [float((pares_b.get(lbl) or {}).get("taxa") or 0.0) for lbl in labels]
        fig.add_trace(go.Bar(name=label_a, x=labels, y=ya, marker_color=COR_AZUL_ESC, legendgroup="a"), row=1, col=1)
        fig.add_trace(go.Bar(name=label_b, x=labels, y=yb, marker_color="#0f766e", legendgroup="b"), row=1, col=1)
    if fins_a:
        labels = [str(r.get("label", "")) for r in fins_a]
        ya = [float(r["taxa"]) if r.get("taxa") is not None else 0.0 for r in fins_a]
        yb = [float((fins_b.get(lbl) or {}).get("taxa") or 0.0) for lbl in labels]
        fig.add_trace(go.Bar(name=label_a, x=labels, y=ya, marker_color=COR_AZUL_ESC, showlegend=False), row=1, col=2)
        fig.add_trace(go.Bar(name=label_b, x=labels, y=yb, marker_color="#0f766e", showlegend=False), row=1, col=2)

    fig.update_layout(
        barmode="group",
        margin=dict(l=20, r=20, t=50, b=80),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(orientation="h", yanchor="bottom", y=1.12, xanchor="center", x=0.5),
        height=400,
    )
    fig.update_yaxes(title_text="Taxa (%)", showgrid=True, gridcolor="rgba(226,232,240,0.5)")
    fig.update_xaxes(tickangle=-25)
    st.plotly_chart(
        fig, use_container_width=True, config={"displayModeBar": False},
        key=chart_key or _plotly_key("conv_par", label_a, label_b),
    )


def _render_conversoes_de_totais(
    totais: Dict[str, float],
    titulo: str = "",
) -> None:
    """Conversões etapa→etapa e diretas (indicador→venda) de um funil — só gráficos."""
    conv = calcular_conversoes_totais(ceil_funil_totais(totais))
    _plot_conversoes_barras_simples(conv, titulo=titulo, chart_key=_plotly_key("conv_tot", titulo))


def _render_conversoes_par_funis(
    totais_a: Dict[str, float],
    label_a: str,
    totais_b: Dict[str, float],
    label_b: str,
) -> None:
    """Conversões etapa→etapa e diretas dos dois funis — só gráficos."""
    st.markdown("###### Conversões etapa → etapa e diretas")
    conv_a = calcular_conversoes_totais(ceil_funil_totais(totais_a))
    conv_b = calcular_conversoes_totais(ceil_funil_totais(totais_b))
    _plot_conversoes_par_barras(conv_a, conv_b, label_a, label_b, chart_key=_plotly_key("conv_par", label_a, label_b))


def _plot_funil_go(
    titulo: str,
    totais: Dict[str, float],
    altura: int = 350,
    chart_key: str = "",
) -> None:
    """Funil estilo marketing (go.Funnel); volumes arredondados para cima."""
    labels = [FUNIL_LABELS[e] for e in FUNIL_ETAPAS]
    ceil_tot = ceil_funil_totais(totais)
    vals = [float(ceil_tot.get(e, 0.0)) for e in FUNIL_ETAPAS]
    fig = _criar_fig_funil(labels, vals, titulo=titulo, altura=altura)
    st.plotly_chart(
        fig, use_container_width=True, config={"displayModeBar": False},
        key=_plotly_key("funil_go", chart_key or titulo),
    )


def _plot_meta_diaria_indicador(
    etapa: str,
    meta_info: Dict[str, Any],
    ultimo_dia: int,
    dia_hoje: int,
) -> None:
    """Realizado + meta diária (reg/médias) para bater a necessidade do indicador."""
    label = FUNIL_LABELS.get(etapa, etapa)
    gap = float(meta_info.get("gap", 0.0))
    nec = float(meta_info.get("necessario_mes", 0.0))
    st.markdown(f"##### Meta diária — {label}")
    st.caption(
        f"Necessário no mês: {fmt_qtd(nec)} · faltam {fmt_qtd(gap)} · "
        "pesos = projeção diária (dia da semana / dia do mês / mês)"
    )
    real = meta_info.get("realizado", pd.DataFrame())
    ritmo_reg = meta_info.get("ritmo_reg", pd.DataFrame())
    ritmo_med = meta_info.get("ritmo_med", pd.DataFrame())

    fig = go.Figure()
    if real is not None and not real.empty:
        fig.add_trace(go.Scatter(
            x=real["dia"], y=real["qtd"],
            mode="lines+markers+text",
            name="Realizado",
            text=[fmt_qtd(float(v)) for v in real["qtd"]],
            textposition="top center",
            textfont=dict(size=10, color=COR_AZUL_ESC, family="Inter"),
            line=dict(color=COR_AZUL_ESC, width=3),
            marker=dict(size=7, color=COR_AZUL_ESC),
        ))
    if ritmo_reg is not None and not ritmo_reg.empty:
        fig.add_trace(go.Scatter(
            x=ritmo_reg["dia"], y=ritmo_reg["qtd"],
            mode="lines+markers+text",
            name="Meta diária (reg.)",
            text=[fmt_qtd(float(v)) for v in ritmo_reg["qtd"]],
            textposition="top center",
            textfont=dict(size=10, color=COR_VERMELHO, family="Inter"),
            line=dict(color=COR_VERMELHO, width=3),
            marker=dict(size=7, color=COR_VERMELHO),
        ))
    if ritmo_med is not None and not ritmo_med.empty:
        fig.add_trace(go.Scatter(
            x=ritmo_med["dia"], y=ritmo_med["qtd"],
            mode="lines+markers+text",
            name="Meta diária (médias)",
            text=[fmt_qtd(float(v)) for v in ritmo_med["qtd"]],
            textposition="bottom center",
            textfont=dict(size=10, color="#0f766e", family="Inter"),
            line=dict(color="#0f766e", width=3, dash="dash"),
            marker=dict(size=7, symbol="diamond", color="#0f766e"),
        ))
    fig.add_vline(
        x=dia_hoje, line_width=1, line_dash="dot", line_color="#64748b",
        annotation_text="Hoje", annotation_position="top",
        annotation_font=dict(color=COR_TEXTO_PRETO, size=11, family="Inter"),
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=11),
        ),
        hovermode="x unified",
        height=360,
    )
    fig.update_xaxes(
        title_text="Dia do mês",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
        range=[0.5, ultimo_dia + 0.5],
    )
    fig.update_yaxes(
        title_text="Qtd. no dia",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.plotly_chart(
        fig, use_container_width=True, config={"displayModeBar": False},
        key=_plotly_key("meta_diaria_ind", etapa),
    )


def render_projecao_funil(proj: Dict[str, Any]) -> None:
    """Seção Streamlit: 3 funis, conversões mês×histórico, metas diárias e projeções."""
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader("Projeção do Funil Comercial (modelo completo vs médias)")
    lags = proj.get("lags") or FUNIL_LAGS
    meta_qtd = float(proj.get("meta_qtd_mes") or 0.0)
    st.caption(
        f"ElasticNet pré-treinado (α={FUNIL_ELASTICNET_ALPHA}, "
        f"l1={FUNIL_ELASTICNET_L1_RATIO}) · features calendário + blocos "
        f"1–7/8–14/15–21/22–28d · inferência NumPy sem retreino no open · "
        f"schema {proj.get('modelo_schema') or FUNIL_MODELO_SCHEMA} · "
        f"treino {proj.get('inicio_treino')}→{proj.get('fim_treino')}"
        + (
            f" · gerado em {proj.get('modelo_treinado_em')}"
            if proj.get("modelo_treinado_em")
            else ""
        )
        + (f" · Meta vendas: {fmt_qtd(meta_qtd)}" if meta_qtd > 0 else "")
        + " · SOQL de produção = MTD + 28 dias de lags"
    )

    etapas = proj.get("etapas") or {}
    r2s = proj.get("r2s") or {}
    r2s_med = proj.get("r2s_medias") or {}

    st.markdown("##### R² do modelo mensal (ElasticNet cal+lags)")
    _render_kpi_cards([
        {
            "lbl": FUNIL_LABELS.get(etapa, etapa),
            "val": f"Reg {fmt_num(float(r2s.get(etapa, (etapas.get(etapa) or {}).get('r2', 0))))}",
            "sub": f"Médias {fmt_num(float(r2s_med.get(etapa, (etapas.get(etapa) or {}).get('r2_medias', 0))))}",
        }
        for etapa in FUNIL_ETAPAS
    ])

    _render_kpi_cards([
        {
            "lbl": FUNIL_LABELS.get(etapa, etapa),
            "val": fmt_qtd(float((etapas.get(etapa) or {}).get("mtd", 0))),
            "sub": (
                f"Proj. híbrida {fmt_qtd(float((etapas.get(etapa) or {}).get('projetado_reg', 0)))}"
                f" (+{fmt_qtd(max(0.0, float((etapas.get(etapa) or {}).get('projetado_reg', 0)) - float((etapas.get(etapa) or {}).get('mtd', 0))))})"
                f" · méd {fmt_qtd(float((etapas.get(etapa) or {}).get('projetado_med', 0)))}"
            ),
        }
        for etapa in FUNIL_ETAPAS
    ])

    # -------------------------------------------------------------------------
    # Seção 1 — Realizado até agora × Projetado até agora
    # -------------------------------------------------------------------------
    totais_mtd = proj.get("totais_mtd") or {
        e: float((etapas.get(e) or {}).get("mtd", 0)) for e in FUNIL_ETAPAS
    }
    totais_proj_mtd = proj.get("totais_proj_mtd") or {
        e: float((etapas.get(e) or {}).get("projetado_mtd_reg", 0)) for e in FUNIL_ETAPAS
    }
    totais_proj = proj.get("totais_proj") or {
        e: float((etapas.get(e) or {}).get("projetado_reg", 0)) for e in FUNIL_ETAPAS
    }
    funil_nec = proj.get("funil_necessario") or totais_proj
    razoes_1 = proj.get("razoes_real_vs_proj_mtd") or _razoes_entre_funis(totais_mtd, totais_proj_mtd)
    razoes_2 = proj.get("razoes_proj_vs_nec") or _razoes_entre_funis(totais_proj, funil_nec)

    st.markdown("##### 1) Realizado até agora × Projetado até agora")
    st.caption(
        "Compara o funil realizado (MTD) com o que o modelo previa até hoje. "
        "Razão > 1 = acima do projetado; < 1 = abaixo. Volumes arredondados para cima."
    )
    c1, c2 = st.columns(2)
    with c1:
        _plot_funil_go("Realizado até agora (MTD)", totais_mtd, altura=380, chart_key="sec1_real")
    with c2:
        _plot_funil_go("Projetado até agora (modelo)", totais_proj_mtd, altura=380, chart_key="sec1_proj")
    _render_razoes_funil(
        razoes_1,
        "Razões realizado / projetado até agora",
        "Por etapa: realizado ÷ projetado MTD do modelo.",
    )
    _render_conversoes_par_funis(
        totais_mtd, "Realizado até agora",
        totais_proj_mtd, "Projetado até agora",
    )

    # -------------------------------------------------------------------------
    # Seção 2 — Projetado do mês × Necessário para a meta
    # -------------------------------------------------------------------------
    st.markdown("##### 2) Projetado do mês × Necessário para a meta")
    st.caption(
        "Projetado do mês = realizado até agora + projeção dos dias restantes. "
        "Necessário = volumes para bater a meta de vendas (conversões históricas). "
        "Razão ≥ 1 sugere que, mantendo o previsto, a meta é atingível."
    )
    c3, c4 = st.columns(2)
    with c3:
        _plot_funil_go("Projetado do mês", totais_proj, altura=380, chart_key="sec2_proj_mes")
    with c4:
        _plot_funil_go("Necessário p/ meta", funil_nec, altura=380, chart_key="sec2_nec")
    _render_razoes_funil(
        razoes_2,
        "Razões projetado do mês / necessário para a meta",
        "Por etapa: projetado do mês ÷ necessário.",
    )
    _render_conversoes_par_funis(
        totais_proj, "Projetado do mês",
        funil_nec, "Necessário p/ meta",
    )

    # -------------------------------------------------------------------------
    # Seção 3 — Meta do dia × Meta da semana
    # -------------------------------------------------------------------------
    funil_dia = proj.get("funil_meta_dia") or {}
    funil_sem = proj.get("funil_meta_semana") or {}
    meta_v_dia = float(math.ceil(float(proj.get("meta_vendas_dia") or 0.0)))
    meta_v_sem = float(math.ceil(float(proj.get("meta_vendas_semana") or 0.0)))
    meta_v_sem_mes = float(math.ceil(float(proj.get("meta_vendas_semana_mes") or 0.0)))
    gap_v = float(math.ceil(float(proj.get("gap_vendas") or 0.0)))
    dias_prox = proj.get("dias_semana_prox") or []
    fim_mes = proj.get("fim_mes")
    fim_semana = proj.get("fim_semana")
    nota_prox = ""
    if dias_prox:
        d0 = min(dias_prox)
        d1 = max(dias_prox)
        fim_mes_txt = f" ({fim_mes.strftime('%d/%m')})" if fim_mes else ""
        nota_prox = (
            f" Semana atravessa o fim do mês{fim_mes_txt}: "
            f"dias {d0.strftime('%d/%m')}–{d1.strftime('%d/%m')} "
            f"completados com o projetado do mês seguinte "
            f"(meta do mês na semana: {fmt_qtd(meta_v_sem_mes)} vendas)."
        )

    st.markdown("##### 3) Meta do dia × Meta da semana (para bater a meta de vendas)")
    st.caption(
        f"Restante do mês: {fmt_qtd(gap_v)} vendas · "
        f"hoje: {fmt_qtd(meta_v_dia)} · semana: {fmt_qtd(meta_v_sem)}"
        + (
            f" (até {fim_semana.strftime('%d/%m')})"
            if fim_semana else ""
        )
        + " · gap do mês distribuído pelos pesos da projeção diária."
        + nota_prox
    )
    c5, c6 = st.columns(2)
    with c5:
        _plot_funil_go(
            f"Meta do dia ({fmt_qtd(meta_v_dia)} vendas)", funil_dia, altura=380, chart_key="sec3_dia",
        )
    with c6:
        _plot_funil_go(
            f"Meta da semana ({fmt_qtd(meta_v_sem)} vendas)", funil_sem, altura=380, chart_key="sec3_sem",
        )
    _render_conversoes_par_funis(
        funil_dia, "Meta do dia",
        funil_sem, "Meta da semana",
    )

    # Comparativo histórico de conversões (referência)
    _render_conversoes_funil(proj.get("conversoes") or {})

    # Metas diárias por indicador
    st.markdown("##### Meta diária por indicador (para bater a meta de vendas)")
    st.caption(
        "Vendas faltantes ÷ conversão histórica indicador→venda = gap do indicador; "
        "distribuído nos dias restantes com pesos da projeção (reg / médias)."
    )
    metas_diarias = proj.get("metas_diarias") or {}
    hoje_proj = proj.get("hoje") or date.today()
    if not isinstance(hoje_proj, date):
        try:
            hoje_proj = pd.to_datetime(hoje_proj).date()
        except Exception:
            hoje_proj = date.today()
    dia_hoje = hoje_proj.day
    ultimo = int(
        proj.get("ultimo_dia")
        or calendar.monthrange(hoje_proj.year, hoje_proj.month)[1]
    )
    for etapa in FUNIL_ETAPAS:
        info = metas_diarias.get(etapa) or {}
        if info:
            _plot_meta_diaria_indicador(etapa, info, ultimo, dia_hoje)

    # Comparativo diário projetado × realizado (ambos os modelos)
    st.markdown("##### Projeção diária × realizado (modelos treinados sem o mês atual)")
    st.caption(
        "Linhas de projetado = previsão do modelo (reg / médias) em cada dia. "
        "Realizado só até hoje. Total do mês = realizado MTD + previsão dos dias restantes."
    )
    for etapa in FUNIL_ETAPAS:
        df = (etapas.get(etapa) or {}).get("diaria", pd.DataFrame())
        _plot_funil_etapa_comparativo(etapa, df, ultimo, dia_hoje)


def render_efeitos_lags_sobre_vendas(
    efeitos: Dict[str, Any],
    mostrar_cards: bool = True,
) -> None:
    """Perfil dos quatro blocos semanais das etapas sobre vendas."""
    st.markdown("<br/>", unsafe_allow_html=True)
    st.subheader("Perfil de lags sobre vendas")
    st.caption(
        f"R² médio: {fmt_num(float(efeitos.get('r2', 0)))} · "
        "Ridge por etapa (vendas ~ calendário + blocos 1–7/8–14/15–21/22–28d) · "
        "pico = bloco de maior efeito acumulado"
    )

    if mostrar_cards:
        resumo = efeitos.get("resumo") or []
        _render_kpi_cards([
            {
                "lbl": str(r.get("label", "")),
                "val": f"{int(r.get('lag_pico', 0))}d",
                "sub": (
                    f"Meia-vida {int(r.get('lag_meia_vida', 0))}d"
                    f" · bloco 1 {fmt_num(float(r.get('efeito_lag1', 0)))}"
                    f" · acum {fmt_num(float(r.get('efeito_acum', 0)))}"
                ),
            }
            for r in resumo
        ])

    perfis = efeitos.get("perfis") or {}
    fig = go.Figure()
    for etapa in FUNIL_DRIVERS:
        df = perfis.get(etapa)
        if df is None or df.empty:
            continue
        cor = FUNIL_CORES_DRIVER.get(etapa, COR_AZUL_ESC)
        fig.add_trace(
            go.Scatter(
                x=df["lag"],
                y=df["efeito"],
                mode="lines+markers",
                name=FUNIL_LABELS.get(etapa, etapa),
                line=dict(color=cor, width=3),
                marker=dict(size=7, color=cor),
            )
        )
    fig.add_hline(y=0, line_width=1, line_dash="dot", line_color="#64748b")
    fig.update_layout(
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.10, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        hovermode="x unified",
        height=420,
    )
    fig.update_xaxes(
        title_text="Lag (dias)",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=7,
    )
    fig.update_yaxes(
        title_text="Efeito nas vendas",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.markdown("##### Efeito por bloco semanal (1–28 dias)")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    fig_c = go.Figure()
    for etapa in FUNIL_DRIVERS:
        df = perfis.get(etapa)
        if df is None or df.empty:
            continue
        cor = FUNIL_CORES_DRIVER.get(etapa, COR_AZUL_ESC)
        fig_c.add_trace(
            go.Scatter(
                x=df["lag"],
                y=df["acumulado"],
                mode="lines+markers",
                name=FUNIL_LABELS.get(etapa, etapa),
                line=dict(color=cor, width=3),
                marker=dict(size=7, color=cor),
            )
        )
    fig_c.add_hline(y=0, line_width=1, line_dash="dot", line_color="#64748b")
    fig_c.update_layout(
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.10, xanchor="center", x=0.5,
            font=dict(color=COR_TEXTO_PRETO, family="Inter", size=12),
        ),
        hovermode="x unified",
        height=380,
    )
    fig_c.update_xaxes(
        title_text="Lag (dias)",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        dtick=1,
    )
    fig_c.update_yaxes(
        title_text="Efeito acumulado",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter"),
        tickfont=dict(color=COR_TEXTO_PRETO, family="Inter"),
        showgrid=True,
        gridcolor="rgba(226,232,240,0.5)",
    )
    st.markdown("##### Efeito acumulado até o lag")
    st.plotly_chart(fig_c, use_container_width=True, config={"displayModeBar": False})


def criar_medidor(
    titulo: str,
    realizado: float,
    meta: float,
    vgv: float,
    meta_vgv: float,
    vendas_qtd: float,
    mostrar_vgv: bool = True,
    *,
    metrica: str = "qtd",
) -> None:
    """Gauge Plotly. metrica='vgv' → percentual VGV realizado / meta VGV."""
    if metrica == "vgv":
        meta_f = float(meta_vgv) if meta_vgv and meta_vgv > 0 else 0.0
        real_calc = float(vgv)
    else:
        meta_f = float(meta) if meta and meta > 0 else 0.0
        real_calc = float(realizado)
    true_perc = (real_calc / meta_f * 100.0) if meta_f > 0 else 0.0
    axis_max = 100
    fill_limit = min(true_perc, axis_max)

    gradient_steps = []
    for i in range(100):
        if i >= fill_limit: break
        ratio = i / 100.0
        r = int(203 + (4 - 203) * ratio)
        g = int(9 + (66 - 9) * ratio)
        b = int(53 + (143 - 53) * ratio)
        end_val = min(i + 1.0, fill_limit)
        gradient_steps.append({"range": [i, end_val], "color": f"rgba({r}, {g}, {b}, 0.9)"})
        
    if fill_limit < 100:
        gradient_steps.append({"range": [fill_limit, 100], "color": "rgba(226, 232, 240, 0.4)"})

    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=true_perc,
            number={
                "suffix": "%",
                "font": {"size": 26, "family": "Montserrat", "color": COR_AZUL_ESC},
                "valueformat": ".1f",
            },
            title={
                "text": titulo,
                "font": {"size": 16, "family": "Montserrat", "color": COR_TEXTO_PRETO},
            },
            gauge={
                "axis": {
                    "range": [0, axis_max],
                    "tickwidth": 1,
                    "tickcolor": "#64748b",
                    "tickfont": {"color": COR_TEXTO_PRETO, "family": "Inter"},
                },
                "bar": {"color": "rgba(0,0,0,0)"},
                "bgcolor": "white",
                "borderwidth": 2,
                "bordercolor": COR_BORDA,
                "steps": gradient_steps,
                "threshold": {
                    "line": {"color": COR_AZUL_ESC, "width": 3},
                    "thickness": 0.8,
                    "value": 100,
                },
            },
        )
    )

    fig.update_layout(
        height=300,
        margin=dict(l=24, r=24, t=56, b=24),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    vgv_linha = (
        f"<strong>VGV:</strong> {fmt_br_milhoes(float(vgv))} / {fmt_br_milhoes(float(meta_vgv))}"
        if mostrar_vgv and float(meta_vgv) > 0
        else ""
    )
    st.markdown(
        f"""
        <div style="text-align:center;font-size:0.85rem;color:{COR_TEXTO_PRETO};margin-top:-8px;line-height:1.4;">
            <strong>Qtd:</strong> {fmt_qtd(vendas_qtd)} / {fmt_qtd(meta_f)} <br/>
            {vgv_linha}
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)


def _distribuir_vendas_coordenador(
    df_vendas: pd.DataFrame,
    df_metas: pd.DataFrame,
) -> pd.DataFrame:
    """Expande vendas por coordenador via merge (substitui iterrows)."""
    if df_metas is None or df_metas.empty or "_peso_coord" not in df_metas.columns:
        out = df_vendas.copy()
        out["_peso_coord"] = 1.0
        if "Região" in out.columns:
            out["Regiao_Coord"] = out["Região"].astype(str).str.strip()
        else:
            out["Regiao_Coord"] = "Não Informado"
        return out
    map_emp = df_metas[["Empreendimento", "Regiao_Coord", "_peso_coord"]].drop_duplicates()
    out = df_vendas.merge(map_emp, on="Empreendimento", how="left")
    if "Região" in out.columns:
        out["Regiao_Coord"] = out["Regiao_Coord"].fillna(out["Região"].astype(str).str.strip())
    out["Regiao_Coord"] = out["Regiao_Coord"].fillna("Não Informado")
    if "_peso_coord" not in out.columns:
        out["_peso_coord"] = 1.0
    else:
        out["_peso_coord"] = pd.to_numeric(out["_peso_coord"], errors="coerce").fillna(1.0)
    return out


def _coalesce_df(val: Any) -> pd.DataFrame:
    """Evita `df or pd.DataFrame()` — DataFrame vazio levanta ambiguous truth value."""
    if val is None or not isinstance(val, pd.DataFrame):
        return pd.DataFrame()
    return val


def _coalesce_dict_df(d: Optional[Dict[str, Any]], key: str) -> pd.DataFrame:
    if not d:
        return pd.DataFrame()
    return _coalesce_df(d.get(key))


def _serie_coluna(df: Optional[pd.DataFrame], col: str) -> pd.Series:
    if df is None or df.empty or col not in df.columns:
        return pd.Series(dtype=object)
    return df[col]


def _opcoes_unicas(*series: pd.Series) -> List[str]:
    vals: set = set()
    for s in series:
        if s is None or s.empty:
            continue
        for c in s.dropna().unique():
            t = str(c).strip()
            if t:
                vals.add(t)
    return sorted(vals)


def _limpar_emp(val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if s.lower() in ("", "nan", "none", "null", "-", "n/a"):
        return ""
    return s


def _janelas_funil_emp(hoje: Optional[date] = None) -> Dict[str, date]:
    """
    Janelas de agregação:
      mes / 7d_mes — restritas ao mês corrente
      7d / 30d — rolling, sem restrição de mês
    """
    hoje = hoje or date.today()
    ini_mes = date(hoje.year, hoje.month, 1)
    return {
        "hoje": hoje,
        "ini_mes": ini_mes,
        "fim": hoje,
        "ini_7d_mes": max(ini_mes, hoje - timedelta(days=6)),
        "ini_7d": hoje - timedelta(days=6),
        "ini_30d": hoje - timedelta(days=29),
    }


def _sf_bool(val: Any) -> bool:
    if val is True:
        return True
    if isinstance(val, (int, float)) and not pd.isna(val) and int(val) == 1:
        return True
    return str(val or "").strip().upper() in ("TRUE", "1", "SIM", "YES")


def _origem_e_digital(val: Any) -> bool:
    return _limpar_emp(val).lower() in ORIGENS_NUCLEO_DIGITAL_NORM


def _is_nucleo_digital_row(row: pd.Series) -> bool:
    """Núcleo digital: Atribuição Digital, Última entrada Digital ou origem digital."""
    if _sf_bool(row.get("Atribuição Digital")):
        return True
    ult = row.get("Última entrada Digital")
    if ult is not None and not (isinstance(ult, float) and pd.isna(ult)):
        if str(ult).strip().lower() not in ("", "nan", "none", "nat"):
            return True
    if _origem_e_digital(row.get("Origem da Conta")):
        return True
    if _origem_e_digital(row.get("Origem do lead")):
        return True
    return False


def _filtrar_nucleo_digital(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    return df.loc[df.apply(_is_nucleo_digital_row, axis=1)].copy()


def _filtrar_opps_abertas(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    col = achar_coluna(df, ALIASES_OPP_FECHADA)
    if not col:
        return df.copy()
    fechada = df[col].map(_sf_bool)
    return df.loc[~fechada].copy()


def _filtrar_df_periodo(
    df: pd.DataFrame,
    col_ou_aliases,
    ini: date,
    fim: date,
) -> pd.DataFrame:
    """Filtra linhas cuja coluna de data está em [ini, fim]. Aceita nome de coluna ou lista de aliases."""
    if df is None or df.empty:
        return df.iloc[0:0].copy() if df is not None else pd.DataFrame()
    if isinstance(col_ou_aliases, (list, tuple)):
        col = achar_coluna(df, list(col_ou_aliases))
    else:
        col = col_ou_aliases
    if not col or col not in df.columns:
        return df.iloc[0:0].copy()
    dt = parse_data_serie(df[col])
    ini_ts = pd.Timestamp(ini)
    fim_ts = pd.Timestamp(fim) + pd.Timedelta(hours=23, minutes=59, seconds=59)
    mask = dt.notna() & (dt >= ini_ts) & (dt <= fim_ts)
    return df.loc[mask].copy()


def _filtrar_df_periodo_ou(
    df: pd.DataFrame,
    listas_aliases: List[List[str]],
    ini: date,
    fim: date,
) -> pd.DataFrame:
    """Mantém linhas com qualquer coluna de data dentro de [ini, fim]."""
    if df is None or df.empty:
        return pd.DataFrame()
    ini_ts = pd.Timestamp(ini)
    fim_ts = pd.Timestamp(fim) + pd.Timedelta(hours=23, minutes=59, seconds=59)
    mask_total = pd.Series(False, index=df.index)
    achou_col = False
    for aliases in listas_aliases:
        col = achar_coluna(df, aliases)
        if not col:
            continue
        achou_col = True
        dt = parse_data_serie(df[col])
        mask_total |= dt.notna() & (dt >= ini_ts) & (dt <= fim_ts)
    if not achou_col:
        return pd.DataFrame()
    return df.loc[mask_total].copy()


def _filtrar_df_emp(df: pd.DataFrame, empreendimento: str) -> pd.DataFrame:
    if df is None or df.empty or not empreendimento:
        return pd.DataFrame()
    col = achar_coluna(df, ALIASES_EMPREENDIMENTO)
    if not col:
        return pd.DataFrame()
    emp_norm = _limpar_emp(empreendimento).lower()
    mask = df[col].map(_limpar_emp).str.lower() == emp_norm
    return df.loc[mask].copy()


def _contar_etapa_df(
    df: pd.DataFrame,
    aliases_data: List[str],
    dedup_fn=None,
) -> float:
    if df is None or df.empty:
        return 0.0
    base = dedup_fn(df) if dedup_fn else df
    col = achar_coluna(base, aliases_data)
    if not col:
        return 0.0
    return float(parse_data_serie(base[col]).notna().sum())


def totais_funil_empreendimento(
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    df_vendas: pd.DataFrame,
    empreendimento: str,
    ini: date,
    fim: date,
) -> Dict[str, float]:
    """Conta as 5 etapas do funil comercial para um empreendimento no período."""
    ag = _filtrar_df_emp(_filtrar_df_periodo(df_ag, ALIASES_DATA_CRIACAO, ini, fim), empreendimento)
    pas_raw = _filtrar_df_emp(_filtrar_df_periodo(df_pastas, COLUNAS_PASTAS_ALIASES + ALIASES_DATA_CRIACAO, ini, fim), empreendimento)
    ven = _filtrar_df_emp(_filtrar_df_periodo(df_vendas, ALIASES_CONTRATO_GERADO, ini, fim), empreendimento)

    ag_d = deduplicar_agendamentos_funil(ag) if not ag.empty else ag
    pas_d = deduplicar_pastas_funil(pas_raw) if not pas_raw.empty else pas_raw
    pas_ap = deduplicar_pastas_aprovadas_funil(pas_raw) if not pas_raw.empty else pas_raw
    ven_d = deduplicar_vendas_funil(filtrar_vendas_comerciais(ven)) if not ven.empty else ven

    col_envio = achar_coluna_primeiro_envio_analise(pas_d)
    col_safi = achar_coluna_aprovacao_safi(pas_ap)
    col_visita = achar_coluna(ag_d, [
        "Data da visita", "Data da Visita", "Data visita", "Activity Date",
    ])

    n_ag = _contar_etapa_df(ag_d, ALIASES_DATA_CRIACAO)
    n_vis = float(parse_data_serie(ag_d[col_visita]).notna().sum()) if col_visita and not ag_d.empty else 0.0
    n_pas = float(parse_data_serie(pas_d[col_envio]).notna().sum()) if col_envio and not pas_d.empty else 0.0
    n_ap = float(parse_data_serie(pas_ap[col_safi]).notna().sum()) if col_safi and not pas_ap.empty else 0.0
    n_ven = float(len(ven_d)) if ven_d is not None and not ven_d.empty else 0.0

    return {
        "agendamentos": n_ag,
        "visitas": n_vis,
        "pastas": n_pas,
        "pastas_aprovadas": n_ap,
        "vendas": n_ven,
    }


def _is_oportunidade_nucleo_digital(row: pd.Series) -> bool:
    return _is_nucleo_digital_row(row)


def contagem_fases_oportunidade(
    df_opps: pd.DataFrame,
    empreendimento: str,
    ini: Optional[date] = None,
    fim: Optional[date] = None,
    apenas_digital: bool = False,
    modo: str = "pipeline",
) -> Tuple[Dict[str, int], int]:
    """
    Agrupa oportunidades por fase (StageName).
    - pipeline: oportunidades abertas (snapshot atual)
    - periodo: criadas ou com mudança de fase em [ini, fim]
    """
    if df_opps is None or df_opps.empty:
        return {}, 0
    base = _filtrar_df_emp(df_opps, empreendimento)
    if base.empty:
        return {}, 0
    if modo == "pipeline":
        base = _filtrar_opps_abertas(base)
    elif ini is not None and fim is not None:
        base = _filtrar_df_periodo_ou(
            base,
            [ALIASES_DATA_CRIACAO, ALIASES_OPP_MUDANCA_FASE],
            ini,
            fim,
        )
    if base.empty:
        return {}, 0
    if apenas_digital:
        base = _filtrar_nucleo_digital(base)
    if base.empty:
        return {}, 0
    col_id = achar_coluna(base, ALIASES_ID_OPORTUNIDADE) or "ID da Oportunidade"
    if col_id in base.columns:
        base = base.drop_duplicates(subset=[col_id])
    col_fase = achar_coluna(base, ALIASES_FASE_OPORTUNIDADE) or "Fase"
    if col_fase not in base.columns:
        return {}, 0
    vc = base[col_fase].fillna("Sem fase").astype(str).str.strip().value_counts()
    por_fase = {str(k): int(v) for k, v in vc.items() if str(k).strip()}
    return por_fase, int(sum(por_fase.values()))


def _ordenar_fases_track(por_fase: Dict[str, int]) -> List[Tuple[str, int]]:
    ordem = {f: i for i, f in enumerate(TRACK_FUNIL_FASES)}
    items = list(por_fase.items())
    items.sort(key=lambda x: (ordem.get(x[0], 999), -x[1], x[0]))
    return items


def _criar_fig_funil_com_conversoes(
    totais: Dict[str, float],
    titulo: str = "",
    altura: int = 380,
    chart_key: str = "",
    etapas: Tuple[str, ...] = FUNIL_ETAPAS,
    labels_map: Optional[Dict[str, str]] = None,
    cores: Optional[List[str]] = None,
) -> go.Figure:
    """Funil vertical — volume + conversão fora do bloco, fonte uniforme preta."""
    labels_map = labels_map or FUNIL_LABELS
    labels = [labels_map.get(e, e) for e in etapas]
    ceil_tot = {e: float(math.ceil(max(0.0, float((totais or {}).get(e, 0.0))))) for e in etapas}
    vals = [float(ceil_tot.get(e, 0.0)) for e in etapas]
    fonte = _fonte_funil_plotly()
    textos: List[str] = []
    for i, etapa in enumerate(etapas):
        linha1 = fmt_funil_valor(vals[i])
        if i == 0:
            linha2 = ""
        else:
            prev = etapas[i - 1]
            taxa = taxa_conversao(float(ceil_tot.get(prev, 0)), float(ceil_tot.get(etapa, 0)))
            linha2 = _fmt_taxa_pct(taxa)
        textos.append(f"{linha1}<br>{linha2}" if linha2 else linha1)
    fig = go.Figure(go.Funnel(
        y=labels,
        x=vals,
        text=textos,
        textinfo="text",
        textposition="outside",
        insidetextfont=fonte,
        outsidetextfont=fonte,
        marker={"color": cores or FUNIL_CORES_NIVEIS},
        connector={"fillcolor": "rgba(4, 66, 143, 0.15)"},
    ))
    return _layout_plotly_preto(fig, titulo=titulo, altura=altura, margin_r=160)


def _criar_fig_track_funnel(
    por_fase: Dict[str, int],
    titulo: str = "",
    altura: int = 320,
    total_rotulo: Optional[str] = None,
) -> go.Figure:
    """
    Track Funnel: barras horizontais cuja largura = % do total.
    Número e % ao lado da barra, texto preto uniforme.
    """
    items = _ordenar_fases_track(por_fase)
    fonte = _fonte_funil_plotly()
    if not items:
        fig = go.Figure()
        fig.add_annotation(
            text="Sem oportunidades neste recorte",
            xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False,
            font=fonte,
        )
        return _layout_plotly_preto(fig, titulo=titulo, altura=180, margin_r=40)
    total = max(sum(por_fase.values()), 1)
    fases = [x[0] for x in items]
    qtds = [x[1] for x in items]
    pcts = [100.0 * q / total for q in qtds]
    cores = [TRACK_FUNIL_CORES[i % len(TRACK_FUNIL_CORES)] for i in range(len(fases))]
    labels_y = [f if len(f) <= 22 else f[:21] + "…" for f in fases]
    textos = [f"{fmt_qtd(q)}  ({fmt_pct_valor(p)})" for q, p in zip(qtds, pcts)]
    fig = go.Figure(go.Bar(
        y=labels_y,
        x=pcts,
        orientation="h",
        text=textos,
        textposition="outside",
        textfont=fonte,
        cliponaxis=False,
        marker=dict(color=cores, line=dict(color="rgba(255,255,255,0.6)", width=1)),
        hovertemplate="%{customdata[0]}<br>Qtd: %{customdata[1]}<br>%{x:.1f}%<extra></extra>",
        customdata=list(zip(fases, qtds)),
    ))
    x_max = max(max(pcts) * 1.35, 10.0)
    fig = _layout_plotly_preto(fig, titulo=titulo, altura=max(altura, 48 + len(fases) * 32), margin_r=100)
    fig.update_xaxes(
        title="% do total",
        title_font=dict(color=COR_TEXTO_PRETO, family="Inter", size=FUNIL_FONTE_TAMANHO),
        range=[0, x_max],
        ticksuffix="%",
    )
    fig.update_yaxes(autorange="reversed")
    if total_rotulo:
        fig.add_annotation(
            text=total_rotulo,
            xref="paper", yref="paper", x=0, y=1.08, showarrow=False,
            xanchor="left", font=fonte,
        )
    return fig


def _kpi_gap_projetado(
    meta_qtd: float,
    realizado: float,
    hoje: Optional[date] = None,
) -> Dict[str, float]:
    """Gap vs ritmo projetado linear (meta proporcional ao dia do mês)."""
    hoje = hoje or date.today()
    dias_mes = calendar.monthrange(hoje.year, hoje.month)[1]
    dia = hoje.day
    meta = max(float(meta_qtd), 0.0)
    real = max(float(realizado), 0.0)
    projetado_pace = meta * dia / dias_mes if dias_mes > 0 else 0.0
    gap = projetado_pace - real
    pct_meta = (real / meta * 100.0) if meta > 0 else 0.0
    pct_pace = (real / projetado_pace * 100.0) if projetado_pace > 0 else 0.0
    return {
        "meta": meta,
        "realizado": real,
        "projetado_pace": projetado_pace,
        "gap": gap,
        "pct_meta": pct_meta,
        "pct_pace": pct_pace,
    }


def _empreendimentos_rj_direcional(
    df_metas: pd.DataFrame,
    df_vendas: pd.DataFrame,
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    df_opps: Optional[pd.DataFrame] = None,
) -> List[str]:
    emps: set = set()
    if df_metas is not None and not df_metas.empty and "Empreendimento" in df_metas.columns:
        for e in df_metas["Empreendimento"].dropna():
            s = _limpar_emp(e)
            if s and s.lower() not in ("total", "geral"):
                emps.add(s)
    for df in (df_vendas, df_ag, df_pastas, df_opps):
        if df is None or df.empty:
            continue
        col = achar_coluna(df, ALIASES_EMPREENDIMENTO)
        if col:
            for e in df[col].dropna():
                s = _limpar_emp(e)
                if s:
                    emps.add(s)
    return sorted(emps)


def _mapear_fase_funil(fase: str) -> Optional[str]:
    f = (fase or "").lower()
    if "visita" in f:
        return "visitas"
    if "análise de crédito" in f or "analise de credito" in f or "análise safi" in f or "analise safi" in f:
        return "pastas"
    if "aprovado safi" in f:
        return "pastas_aprovadas"
    if "fechado e ganho" in f or f == "contrato gerado":
        return "vendas"
    if any(x in f for x in ("aguardando", "atendimento", "elaboração", "elaboracao", "proposta")):
        return "agendamentos"
    return None


def totais_funil_digital_oportunidades(
    df_opps: pd.DataFrame,
    df_vendas: pd.DataFrame,
    empreendimento: str,
    ini: date,
    fim: date,
) -> Dict[str, float]:
    """Funil digital via oportunidades (núcleo digital) + vendas com origem digital."""
    out = {e: 0.0 for e in FUNIL_DIGITAL_ETAPAS}
    if df_opps is not None and not df_opps.empty:
        base = _filtrar_df_emp(_filtrar_df_periodo_ou(
            df_opps,
            [ALIASES_DATA_CRIACAO, ALIASES_OPP_MUDANCA_FASE],
            ini,
            fim,
        ), empreendimento)
        if not base.empty:
            base = _filtrar_nucleo_digital(base)
        col_id = achar_coluna(base, ALIASES_ID_OPORTUNIDADE) or "ID da Oportunidade"
        col_fase = achar_coluna(base, ALIASES_FASE_OPORTUNIDADE) or "Fase"
        if not base.empty and col_fase in base.columns:
            base = base.drop_duplicates(subset=[col_id])
            out["leads"] = float(len(base))
            for _, row in base.iterrows():
                etapa = _mapear_fase_funil(str(row.get(col_fase, "")))
                if etapa:
                    out[etapa] += 1.0
    df_ven_d = df_vendas.copy() if df_vendas is not None else pd.DataFrame()
    if not df_ven_d.empty:
        df_ven_d = _filtrar_nucleo_digital(df_ven_d)
        ven = _filtrar_df_emp(_filtrar_df_periodo(df_ven_d, ALIASES_CONTRATO_GERADO, ini, fim), empreendimento)
        ven_d = deduplicar_vendas_funil(filtrar_vendas_comerciais(ven)) if not ven.empty else ven
        out["vendas"] = max(out["vendas"], float(len(ven_d)) if ven_d is not None and not ven_d.empty else 0.0)
    return out


def _meta_qtd_empreendimento(df_metas: pd.DataFrame, emp: str, mes_num: int) -> float:
    if df_metas is None or df_metas.empty or "Mes_Num" not in df_metas.columns:
        return 0.0
    m = df_metas[
        (df_metas["Empreendimento"].map(_limpar_emp) == emp)
        & (df_metas["Mes_Num"] == mes_num)
    ]
    return float(m["Meta_Qtd"].sum()) if not m.empty and "Meta_Qtd" in m.columns else 0.0


def _classificar_status_estoque(status: Any) -> str:
    """Retorna chave interna: disponivel | mirror | fora_venda | fora_comercial | outro."""
    s = str(status or "").strip()
    if not s:
        return "outro"
    if s in ESTOQUE_STATUS_VENDAVEL:
        return "mirror" if s == "Mirror" else "disponivel"
    s_low = s.lower()
    if "fora de venda - comercial" in s_low:
        return "fora_comercial"
    if "fora de venda" in s_low:
        return "fora_venda"
    return "outro"


def resumo_estoque_empreendimentos(df_estoque: pd.DataFrame) -> Dict[str, Dict[str, int]]:
    """Agrega unidades por empreendimento (status do relatório de estoque)."""
    out: Dict[str, Dict[str, int]] = {}
    if df_estoque is None or df_estoque.empty:
        return out
    col_emp = achar_coluna(df_estoque, ALIASES_EMPREENDIMENTO) or "Empreendimento"
    col_st = achar_coluna(df_estoque, ALIASES_STATUS_UNIDADE) or "StatusUnidade__c"
    for _, row in df_estoque.iterrows():
        emp = _limpar_emp(row.get(col_emp))
        if not emp:
            continue
        bucket = out.setdefault(
            emp,
            {
                "disponivel": 0,
                "mirror": 0,
                "fora_venda": 0,
                "fora_comercial": 0,
                "vendavel": 0,
                "total_status": 0,
            },
        )
        bucket["total_status"] += 1
        cls = _classificar_status_estoque(row.get(col_st))
        if cls == "disponivel":
            bucket["disponivel"] += 1
            bucket["vendavel"] += 1
        elif cls == "mirror":
            bucket["mirror"] += 1
            bucket["vendavel"] += 1
        elif cls == "fora_venda":
            bucket["fora_venda"] += 1
        elif cls == "fora_comercial":
            bucket["fora_comercial"] += 1
    return out


def metricas_liberacao_estoque_por_emp(
    estoque_map: Dict[str, Dict[str, int]],
    total_por_emp: Optional[Dict[str, int]] = None,
) -> Dict[str, Dict[str, Any]]:
    """
    Por empreendimento:
    - liberadas = unidades nos 4 status do relatório de estoque
    - disponíveis = status Disponível
    - total = todas as unidades Produto__c (se total_por_emp informado)
    """
    total_por_emp = total_por_emp or {}
    out: Dict[str, Dict[str, Any]] = {}
    for emp, buckets in estoque_map.items():
        liberadas = int(buckets.get("total_status", 0))
        disponiveis = int(buckets.get("disponivel", 0))
        total = int(total_por_emp.get(emp, liberadas))
        out[emp] = {
            "disponivel": disponiveis,
            "liberadas": liberadas,
            "total": total,
            "pct_disp_liberadas": (disponiveis / liberadas * 100.0) if liberadas > 0 else 0.0,
            "pct_liberadas_total": (liberadas / total * 100.0) if total > 0 else 0.0,
        }
    return out


def _media_mensal_vendas_emp(
    df_vendas: pd.DataFrame,
    empreendimento: str,
    col_contrato: str,
    hoje: Optional[date] = None,
    meses: int = 6,
) -> float:
    """Média de vendas/mês nos meses completos anteriores ao mês corrente."""
    hoje = hoje or date.today()
    ven = _filtrar_df_emp(filtrar_vendas_comerciais(df_vendas), empreendimento)
    if ven.empty or not col_contrato or col_contrato not in ven.columns:
        return 0.0
    ven = ven.copy()
    ven["_dt"] = parse_data_serie(ven[col_contrato])
    ven = ven.dropna(subset=["_dt"])
    if ven.empty:
        return 0.0
    ini_mes_atual = date(hoje.year, hoje.month, 1)
    ven = ven.loc[ven["_dt"].dt.date < ini_mes_atual]
    if ven.empty:
        return 0.0
    ven["_mes"] = ven["_dt"].dt.to_period("M")
    por_mes = ven.groupby("_mes").size()
    if por_mes.empty:
        return 0.0
    ultimos = por_mes.sort_index().tail(meses)
    return float(ultimos.mean()) if len(ultimos) else 0.0


def calcular_meta_qtd_empreendimento(
    df_vendas_hist: pd.DataFrame,
    empreendimento: str,
    col_contrato: Optional[str],
    estoque: Optional[Dict[str, int]],
    vendas_mtd: float,
    df_metas: pd.DataFrame,
    mes_num: int,
    hoje: Optional[date] = None,
    incluir_mes: bool = True,
) -> Dict[str, Any]:
    """
    Meta mensal por empreendimento:
      1) Regressão OLS sobre histórico de vendas (mesma lógica do painel geral)
      2) Teto por estoque vendável (Disponível + Mirror): vendas MTD + estoque restante
    Fallback: média mensal recente ou meta da planilha.
    """
    hoje = hoje or date.today()
    tem_estoque = estoque is not None
    est = estoque or {}
    estoque_vendavel = int(est.get("vendavel", 0)) if tem_estoque else None
    meta_planilha = _meta_qtd_empreendimento(df_metas, empreendimento, mes_num)
    vendas_mtd = max(float(vendas_mtd or 0.0), 0.0)
    meta_cap_estoque = (
        vendas_mtd + int(estoque_vendavel)
        if tem_estoque and estoque_vendavel is not None
        else None
    )
    meta_reg = 0.0
    origem = "planilha"
    r2 = None

    col_c = col_contrato or achar_coluna(df_vendas_hist, ALIASES_CONTRATO_GERADO)
    ven_emp = _filtrar_df_emp(filtrar_vendas_comerciais(df_vendas_hist), empreendimento)

    if col_c and not ven_emp.empty:
        serie = serie_diaria_contratos(ven_emp, col_c)
        inicio, fim_treino = janela_treino_meses_exatos(hoje)
        treino = calendario_diario(inicio, fim_treino, serie)
        if len(treino) >= 30 and float(treino["qtd"].sum()) > 0:
            modelo = treinar_regressao_vendas_diarias(treino, incluir_mes=incluir_mes)
            r2 = _r2_treino(treino, modelo, incluir_mes=incluir_mes)
            ultimo_dia = calendar.monthrange(hoje.year, hoje.month)[1]
            dias_mes = [date(hoje.year, hoje.month, d) for d in range(1, ultimo_dia + 1)]
            pred_reg_mes = prever_qtd_dias(modelo, dias_mes, incluir_mes=incluir_mes)
            intensidades_fim = calcular_intensidades_fim_mes(treino, janelas=JANELAS_FIM_MES)
            pred_reg_mes = aplicar_sazonalidade_fim_mes(
                pred_reg_mes, dias_mes, ultimo_dia, intensidades_fim,
            )
            meta_reg = float(np.sum(pred_reg_mes))
            origem = "regressao"
        else:
            media_m = _media_mensal_vendas_emp(ven_emp, empreendimento, col_c, hoje)
            if media_m > 0:
                meta_reg = media_m
                origem = "media_mensal"

    if meta_reg <= 0:
        if meta_planilha > 0:
            meta_reg = meta_planilha
            origem = "planilha"
        else:
            media_m = _media_mensal_vendas_emp(ven_emp, empreendimento, col_c or "", hoje) if col_c else 0.0
            if media_m > 0:
                meta_reg = media_m
                origem = "media_mensal"

    limitado_estoque = False
    if meta_cap_estoque is not None and meta_reg > meta_cap_estoque + 1e-9:
        meta_final = meta_cap_estoque
        limitado_estoque = True
    else:
        meta_final = meta_reg

    meta_final = max(meta_final, vendas_mtd)
    ev = int(estoque_vendavel or 0)
    meses_estoque = (
        ev / (meta_reg / 30.0)
        if meta_reg > 1e-9 and ev > 0
        else None
    )

    return {
        "meta_final": float(math.ceil(max(0.0, meta_final))),
        "meta_regressao": float(math.ceil(max(0.0, meta_reg))),
        "meta_planilha": float(meta_planilha),
        "meta_cap_estoque": float(meta_cap_estoque) if meta_cap_estoque is not None else None,
        "estoque_vendavel": ev if tem_estoque else None,
        "estoque_disponivel": int(est.get("disponivel", 0)),
        "estoque_mirror": int(est.get("mirror", 0)),
        "estoque_fora_venda": int(est.get("fora_venda", 0)),
        "estoque_fora_comercial": int(est.get("fora_comercial", 0)),
        "limitado_estoque": limitado_estoque,
        "meses_estoque_restante": meses_estoque,
        "origem": origem,
        "r2_treino": r2,
        "vendas_mtd": vendas_mtd,
    }


def _caption_meta_empreendimento(meta_info: Dict[str, Any]) -> str:
    """Texto explicativo da meta (regressão × estoque)."""
    origem = meta_info.get("origem", "?")
    meta_reg = meta_info.get("meta_regressao", 0)
    est_v = meta_info.get("estoque_vendavel")
    disp = meta_info.get("estoque_disponivel", 0)
    mir = meta_info.get("estoque_mirror", 0)
    parts = [f"Meta = regressão histórica ({fmt_qtd(meta_reg)})"]
    if est_v is not None:
        parts.append(f"estoque vendável {est_v} (Disp {disp} · Mirror {mir})")
    else:
        parts.append("estoque: sem dados SF para este empreendimento")
    if meta_info.get("limitado_estoque"):
        parts.append(f"limitada a {fmt_qtd(meta_info.get('meta_cap_estoque', 0))} (MTD + estoque)")
    meses_est = meta_info.get("meses_estoque_restante")
    if meses_est is not None and est_v is not None and est_v > 0:
        parts.append(f"~{fmt_num(meses_est)} meses de estoque ao ritmo projetado")
    r2 = meta_info.get("r2_treino")
    if r2 is not None and origem == "regressao":
        parts.append(f"R² treino {fmt_num(r2)}")
    if origem != "regressao":
        parts.append(f"fallback: {origem}")
    return " · ".join(parts)


def _plot_funil_pair(
    tot_a: Dict[str, float],
    titulo_a: str,
    tot_b: Dict[str, float],
    titulo_b: str,
    key_prefix: str,
    etapas: Tuple[str, ...] = FUNIL_ETAPAS,
    labels_map: Optional[Dict[str, str]] = None,
    cores: Optional[List[str]] = None,
) -> None:
    ca, cb = st.columns(2)
    kwargs = {"etapas": etapas, "labels_map": labels_map, "cores": cores}
    with ca:
        st.plotly_chart(
            _criar_fig_funil_com_conversoes(tot_a, titulo=titulo_a, **kwargs),
            use_container_width=True,
            config={"displayModeBar": False},
            key=_plotly_key(key_prefix, "a", titulo_a),
        )
    with cb:
        st.plotly_chart(
            _criar_fig_funil_com_conversoes(tot_b, titulo=titulo_b, **kwargs),
            use_container_width=True,
            config={"displayModeBar": False},
            key=_plotly_key(key_prefix, "b", titulo_b),
        )


def _plot_track_pair(
    fases_a: Dict[str, int],
    titulo_a: str,
    total_a: int,
    fases_b: Dict[str, int],
    titulo_b: str,
    total_b: int,
    key_prefix: str,
) -> None:
    ca, cb = st.columns(2)
    with ca:
        st.caption(f"**{titulo_a}** · {total_a} opp.")
        st.plotly_chart(
            _criar_fig_track_funnel(
                fases_a, titulo=titulo_a, total_rotulo=f"OPORTUNIDADES · {total_a}",
            ),
            use_container_width=True,
            config={"displayModeBar": False},
            key=_plotly_key(key_prefix, "track_a", titulo_a),
        )
    with cb:
        st.caption(f"**{titulo_b}** · {total_b} opp.")
        st.plotly_chart(
            _criar_fig_track_funnel(
                fases_b, titulo=titulo_b, total_rotulo=f"OPORTUNIDADES · {total_b}",
            ),
            use_container_width=True,
            config={"displayModeBar": False},
            key=_plotly_key(key_prefix, "track_b", titulo_b),
        )


def _render_secao_funil_empreendimento(
    emp: str,
    df_metas: pd.DataFrame,
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    df_opps: pd.DataFrame,
    df_ven: pd.DataFrame,
    janelas: Dict[str, date],
    key_slug: str,
    df_vendas_hist: Optional[pd.DataFrame] = None,
    col_contrato: Optional[str] = None,
    estoque_emp: Optional[Dict[str, int]] = None,
    total_unidades_emp: Optional[int] = None,
    sinal_sobre_renda: Optional[float] = None,
) -> None:
    """Uma seção completa por empreendimento (sem nova consulta SOQL)."""
    hoje = janelas["hoje"]
    ini_mes = janelas["ini_mes"]
    fim = janelas["fim"]
    ini_7d_mes = janelas["ini_7d_mes"]
    ini_7d = janelas["ini_7d"]
    ini_30d = janelas["ini_30d"]
    mes_num = hoje.month

    tot_mes = totais_funil_empreendimento(df_ag, df_pastas, df_ven, emp, ini_mes, fim)
    vendas_mtd = float(tot_mes.get("vendas", 0.0))
    meta_info = calcular_meta_qtd_empreendimento(
        df_vendas_hist if df_vendas_hist is not None else df_ven,
        emp,
        col_contrato,
        estoque_emp,
        vendas_mtd,
        df_metas,
        mes_num,
        hoje=hoje,
    )
    meta_qtd = float(meta_info.get("meta_final", 0.0))
    tot_7d_mes = totais_funil_empreendimento(df_ag, df_pastas, df_ven, emp, ini_7d_mes, fim)
    tot_7d = totais_funil_empreendimento(df_ag, df_pastas, df_ven, emp, ini_7d, fim)
    tot_30d = totais_funil_empreendimento(df_ag, df_pastas, df_ven, emp, ini_30d, fim)
    kpi = _kpi_gap_projetado(meta_qtd, tot_mes.get("vendas", 0.0), hoje)

    st.markdown(f"#### {emp}")
    gap_txt = f"{fmt_qtd(kpi['gap'])} abaixo do projetado" if kpi["gap"] > 0.01 else "no ritmo ou acima"
    st.markdown(
        f"""
        <div class="vel-kpi-row aba-funil-kpi">
            <div class="vel-kpi"><div class="lbl">Meta mês</div><div class="val">{int(kpi['meta'])}</div></div>
            <div class="vel-kpi"><div class="lbl">Meta regressão</div><div class="val">{int(meta_info.get('meta_regressao', 0))}</div></div>
            <div class="vel-kpi"><div class="lbl">Estoque vendável</div><div class="val">{
                int(meta_info['estoque_vendavel']) if meta_info.get('estoque_vendavel') is not None else '—'
            }</div></div>
            <div class="vel-kpi"><div class="lbl">Vendas MTD</div><div class="val">{fmt_qtd(kpi['realizado'])}</div></div>
            <div class="vel-kpi"><div class="lbl">Projetado até hoje</div><div class="val">{fmt_qtd(kpi['projetado_pace'])}</div></div>
            <div class="vel-kpi"><div class="lbl">Gap vs projetado</div><div class="val">{fmt_qtd(kpi['gap'])}</div></div>
            <div class="vel-kpi"><div class="lbl">% meta</div><div class="val">{fmt_pct_valor(kpi['pct_meta'])}</div></div>
            <div class="vel-kpi"><div class="lbl">% ritmo</div><div class="val">{fmt_pct_valor(kpi['pct_pace'])}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(f"Ritmo: {gap_txt} · projeção linear ao dia {hoje.day}.")
    st.caption(_caption_meta_empreendimento(meta_info))

    if estoque_emp:
        lib = metricas_liberacao_estoque_por_emp(
            {_limpar_emp(emp): estoque_emp},
            {_limpar_emp(emp): total_unidades_emp} if total_unidades_emp is not None else None,
        ).get(_limpar_emp(emp), {})
        disp = int(lib.get("disponivel", estoque_emp.get("disponivel", 0)))
        liberadas = int(lib.get("liberadas", estoque_emp.get("total_status", 0)))
        total_u = int(lib.get("total", total_unidades_emp or liberadas))
        pct_dl = float(lib.get("pct_disp_liberadas", 0.0))
        pct_lt = float(lib.get("pct_liberadas_total", 0.0))
        sinal_txt = fmt_pct_valor(sinal_sobre_renda * 100) if sinal_sobre_renda is not None else "—"
        st.markdown(
            f"""
            <div class="vel-kpi-row aba-funil-kpi">
                <div class="vel-kpi"><div class="lbl">Disp. / liberadas</div>
                <div class="val">{disp} / {liberadas} ({pct_dl:.0f}%)</div></div>
                <div class="vel-kpi"><div class="lbl">Liberadas / total</div>
                <div class="val">{liberadas} / {total_u} ({pct_lt:.0f}%)</div></div>
                <div class="vel-kpi"><div class="lbl">Σ sinais / Σ renda</div>
                <div class="val">{sinal_txt}</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.caption(
            "Estoque: liberadas = Disponível + Mirror + Fora de venda + Fora de Venda - Comercial · "
            "total = todas as unidades Produto__c · sinal/renda = pastas do período com cotação vinculada."
        )

    st.markdown("##### Mês corrente (MTD)")
    st.caption(
        f"Funil comercial restrito a {ini_mes.strftime('%d/%m')}–{fim.strftime('%d/%m')} · "
        f"7 dias do mês = {ini_7d_mes.strftime('%d/%m')}–{fim.strftime('%d/%m')}"
    )
    _plot_funil_pair(
        tot_mes,
        f"MTD ({ini_mes.strftime('%d/%m')}→{fim.strftime('%d/%m')})",
        tot_7d_mes,
        f"7d no mês ({ini_7d_mes.strftime('%d/%m')}→{fim.strftime('%d/%m')})",
        _plotly_key("funil_mes", key_slug),
    )

    fases_pipeline, n_pipeline = contagem_fases_oportunidade(
        df_opps, emp, modo="pipeline", apenas_digital=False,
    )
    fases_mes, n_mes = contagem_fases_oportunidade(
        df_opps, emp, ini_mes, fim, modo="periodo", apenas_digital=False,
    )
    fases_7d_mes, n_7d_mes = contagem_fases_oportunidade(
        df_opps, emp, ini_7d_mes, fim, modo="periodo", apenas_digital=False,
    )
    st.markdown("##### Oportunidades por fase — mês corrente")
    st.caption(
        "Pipeline aberto (fase atual) · criadas/movimentadas no MTD · criadas/movimentadas nos 7d do mês"
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Pipeline aberto", n_pipeline)
        st.plotly_chart(
            _criar_fig_track_funnel(
                fases_pipeline, titulo="Pipeline aberto", total_rotulo=f"OPORTUNIDADES · {n_pipeline}",
            ),
            use_container_width=True,
            config={"displayModeBar": False},
            key=_plotly_key("pipe", key_slug),
        )
    with c2:
        st.metric("No mês (MTD)", n_mes)
        st.plotly_chart(
            _criar_fig_track_funnel(
                fases_mes, titulo="Movimentadas no mês", total_rotulo=f"OPORTUNIDADES · {n_mes}",
            ),
            use_container_width=True,
            config={"displayModeBar": False},
            key=_plotly_key("fase_mes", key_slug),
        )
    with c3:
        st.metric("7d no mês", n_7d_mes)
        st.plotly_chart(
            _criar_fig_track_funnel(
                fases_7d_mes, titulo="7 dias no mês", total_rotulo=f"OPORTUNIDADES · {n_7d_mes}",
            ),
            use_container_width=True,
            config={"displayModeBar": False},
            key=_plotly_key("fase_7d_mes", key_slug),
        )

    st.markdown("##### Janela rolling (sem restrição de mês)")
    st.caption(
        f"Funil e fases nos últimos 30 dias ({ini_30d.strftime('%d/%m')}→{fim.strftime('%d/%m')}) "
        f"e 7 dias ({ini_7d.strftime('%d/%m')}→{fim.strftime('%d/%m')})"
    )
    _plot_funil_pair(
        tot_30d,
        f"30 dias ({ini_30d.strftime('%d/%m')}→{fim.strftime('%d/%m')})",
        tot_7d,
        f"7 dias ({ini_7d.strftime('%d/%m')}→{fim.strftime('%d/%m')})",
        _plotly_key("funil_roll", key_slug),
    )
    fases_30d, n_30d = contagem_fases_oportunidade(
        df_opps, emp, ini_30d, fim, modo="periodo", apenas_digital=False,
    )
    fases_7d, n_7d = contagem_fases_oportunidade(
        df_opps, emp, ini_7d, fim, modo="periodo", apenas_digital=False,
    )
    _plot_track_pair(
        fases_30d, "30 dias rolling", n_30d,
        fases_7d, "7 dias rolling", n_7d,
        _plotly_key("roll", key_slug),
    )

    st.markdown("##### Marketing digital (núcleo digital)")
    st.caption(
        "Atribuição Digital · Última entrada Digital · origem digital "
        f"({len(ORIGENS_NUCLEO_DIGITAL)} canais mapeados)"
    )
    tot_d_mes = totais_funil_digital_oportunidades(df_opps, df_ven, emp, ini_mes, fim)
    tot_d_7d_mes = totais_funil_digital_oportunidades(df_opps, df_ven, emp, ini_7d_mes, fim)
    _plot_funil_pair(
        tot_d_mes,
        "Digital — MTD",
        tot_d_7d_mes,
        "Digital — 7d no mês",
        _plotly_key("dig_mes", key_slug),
        etapas=FUNIL_DIGITAL_ETAPAS,
        labels_map=FUNIL_DIGITAL_LABELS,
        cores=FUNIL_DIGITAL_CORES,
    )
    tot_d_30d = totais_funil_digital_oportunidades(df_opps, df_ven, emp, ini_30d, fim)
    tot_d_7d = totais_funil_digital_oportunidades(df_opps, df_ven, emp, ini_7d, fim)
    _plot_funil_pair(
        tot_d_30d,
        "Digital — 30d rolling",
        tot_d_7d,
        "Digital — 7d rolling",
        _plotly_key("dig_roll", key_slug),
        etapas=FUNIL_DIGITAL_ETAPAS,
        labels_map=FUNIL_DIGITAL_LABELS,
        cores=FUNIL_DIGITAL_CORES,
    )
    fp_d, n_fp_d = contagem_fases_oportunidade(
        df_opps, emp, modo="pipeline", apenas_digital=True,
    )
    fases_d_mes, n_d_mes = contagem_fases_oportunidade(
        df_opps, emp, ini_mes, fim, modo="periodo", apenas_digital=True,
    )
    _plot_track_pair(
        fp_d, "Digital pipeline aberto", n_fp_d,
        fases_d_mes, "Digital no mês", n_d_mes,
        _plotly_key("dig_fase", key_slug),
    )
    st.markdown("---")


def render_aba_funil_empreendimentos(
    df_metas: pd.DataFrame,
    df_vendas: pd.DataFrame,
    col_contrato_gerado: Optional[str],
    filtros_glob: Optional["FiltrosGlobais"] = None,
) -> None:
    """Aba: todos os empreendimentos · uma carga SOQL · seções empilhadas."""
    st.subheader("Funil por Empreendimento — Direcional · RJ")
    st.caption(
        "Uma consulta Salesforce · vendas comerciais · "
        "meta por regressão histórica + teto de estoque (Disp/Mirror) · "
        "MTD + 7d no mês · rolling 30d/7d · pipeline aberto por fase · núcleo digital"
    )
    janelas = _janelas_funil_emp(date.today())
    st.markdown(
        f"""
        <style>
        .stTabs [data-baseweb="tab-panel"]:nth-of-type(2),
        .stTabs [data-baseweb="tab-panel"]:nth-of-type(2) [data-testid="stCaptionContainer"],
        .stTabs [data-baseweb="tab-panel"]:nth-of-type(2) [data-testid="stCaptionContainer"] p,
        .stTabs [data-baseweb="tab-panel"]:nth-of-type(2) [data-testid="stHeader"],
        .stTabs [data-baseweb="tab-panel"]:nth-of-type(2) [data-testid="stMarkdownContainer"] h4,
        .stTabs [data-baseweb="tab-panel"]:nth-of-type(2) [data-testid="stMarkdownContainer"] h5,
        .stTabs [data-baseweb="tab-panel"]:nth-of-type(2) [data-testid="stExpander"] label,
        .stTabs [data-baseweb="tab-panel"]:nth-of-type(2) [data-testid="stExpander"] p,
        .stTabs [data-baseweb="tab-panel"]:nth-of-type(2) [data-testid="stMetricLabel"],
        .stTabs [data-baseweb="tab-panel"]:nth-of-type(2) [data-testid="stMetricValue"],
        .aba-funil-kpi .vel-kpi .val, .aba-funil-kpi .vel-kpi .lbl {{
            color: {COR_TEXTO_PRETO} !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    try:
        pacote = carregar_funil_empreendimento_sf()
        df_ag = _coalesce_df(pacote.get("agendamentos"))
        df_pastas = _coalesce_df(pacote.get("pastas"))
        df_opps = _coalesce_df(pacote.get("oportunidades"))
        df_ven_funil = _coalesce_df(pacote.get("vendas"))
        df_estoque = _coalesce_df(pacote.get("estoque"))
        t_sf = float((pacote.get("timings") or {}).get("total_s", 0.0))
        st.caption(
            f"Janela SF desde {pacote.get('inicio_janela', '?')}"
            + (f" · carregado em {fmt_num(t_sf)}s" if t_sf else "")
            + f" · ag {len(df_ag):,} · pas {len(df_pastas):,} · "
            f"opp {len(df_opps):,} · ven {len(df_ven_funil):,} · est {len(df_estoque):,}".replace(",", ".")
            + " · dados reutilizados em todos os empreendimentos abaixo"
        )
    except Exception as exc:
        st.error(f"Não foi possível carregar dados do funil por empreendimento: {exc}")
        return

    empreendimentos = _empreendimentos_rj_direcional(
        df_metas, df_ven_funil, df_ag, df_pastas, df_opps,
    )
    if not empreendimentos:
        st.info("Nenhum empreendimento encontrado para Direcional · RJ.")
        return

    # Ordena por vendas MTD (maior primeiro)
    ini_mes = janelas["ini_mes"]
    fim = janelas["fim"]

    def _score_emp(e: str) -> float:
        t = totais_funil_empreendimento(df_ag, df_pastas, df_ven_funil, e, ini_mes, fim)
        return float(t.get("vendas", 0)) * 1000 + sum(float(t.get(k, 0)) for k in FUNIL_ETAPAS)

    empreendimentos = sorted(empreendimentos, key=_score_emp, reverse=True)
    estoque_map = resumo_estoque_empreendimentos(df_estoque)
    try:
        total_unidades_map = carregar_total_unidades_por_emp_sf()
    except Exception:
        total_unidades_map = {}
    try:
        df_cot_funil = carregar_cotacoes_painel_sf()
    except Exception:
        df_cot_funil = pd.DataFrame()
    col_c_hist = col_contrato_gerado or achar_coluna(df_vendas, ALIASES_CONTRATO_GERADO)

    # Mapa empreendimento → canal (derivado das vendas SF)
    mapa_emp_canal: Dict[str, str] = {}
    if not df_ven_funil.empty and "Empreendimento" in df_ven_funil.columns:
        col_ca = achar_coluna(df_ven_funil, ["Canal", "Imobiliária", "Imobiliaria"])
        for emp_n in empreendimentos:
            sub = df_ven_funil[df_ven_funil["Empreendimento"].map(_limpar_emp) == _limpar_emp(emp_n)]
            if sub.empty:
                continue
            if "Canal_Agrupado" in sub.columns:
                ca = sub["Canal_Agrupado"].mode()
                mapa_emp_canal[emp_n] = str(ca.iloc[0]) if not ca.empty else "RIO"
            elif col_ca and col_ca in sub.columns:
                ca = sub[col_ca].astype(str).str.upper().str.strip().mode()
                mapa_emp_canal[emp_n] = str(ca.iloc[0]) if not ca.empty else "RIO"

    canais_disp = sorted(set(mapa_emp_canal.values()) | {"RIO", "DIR", "PARC", "RJ"})

    if filtros_glob is not None:
        pool_emp = filtros_glob.emps_sel or empreendimentos
        canais_sel = [] if filtros_glob.canal_sel in ("Todos", "TODOS", "RIO", "") else [filtros_glob.canal_sel]
        data_ini = filtros_glob.data_ini
        data_fim = filtros_glob.data_fim
        st.caption(
            f"Filtros globais: {data_ini:%d/%m/%Y} → {data_fim:%d/%m/%Y} · "
            f"Canal {filtros_glob.canal_sel or 'Todos'}"
        )
        emp_sel = st.selectbox(
            "Empreendimento",
            pool_emp if pool_emp else empreendimentos,
            index=0,
            key="funil_emp_sel",
        )
    else:
        st.markdown("#### Filtros")
        fc1, fc2, fc3, fc4 = st.columns(4)
        with fc1:
            emp_sel = st.selectbox(
                "Empreendimento",
                empreendimentos,
                index=0,
                key="funil_emp_sel",
            )
        with fc2:
            canais_sel = st.multiselect(
                "Canal",
                canais_disp,
                default=[],
                key="funil_canal_sel",
            )
        with fc3:
            data_ini = st.date_input(
                "Data inicial análise",
                value=janelas["ini_mes"],
                key="funil_data_ini",
            )
        with fc4:
            data_fim = st.date_input(
                "Data final análise",
                value=janelas["fim"],
                key="funil_data_fim",
            )

    if canais_sel and mapa_emp_canal.get(emp_sel, "RIO") not in canais_sel:
        st.info("Nenhum empreendimento corresponde aos filtros selecionados.")
        return
    emps_render = [emp_sel]

    janelas = _janelas_funil_emp(date.today())
    janelas["ini_mes"] = data_ini
    janelas["fim"] = data_fim
    janelas["ini_7d_mes"] = max(data_ini, data_fim - timedelta(days=6))

    sinal_renda_map = calcular_sinal_sobre_renda_por_emp(
        df_cot_funil, df_pastas, data_ini, data_fim,
    )

    st.info(
        f"Empreendimento: **{emp_sel}** · meta = regressão ({PAINEL_MESES_VENDAS}m vendas) "
        f"limitada por estoque vendável (relatório {SF_REPORT_ESTOQUE_ID})."
    )

    emp = emps_render[0]
    key_slug = _plotly_key("emp", emp)
    _render_secao_funil_empreendimento(
        emp,
        df_metas,
        df_ag,
        df_pastas,
        df_opps,
        df_ven_funil,
        janelas,
        key_slug,
        df_vendas_hist=df_vendas,
        col_contrato=col_c_hist,
        estoque_emp=estoque_map.get(_limpar_emp(emp)) or estoque_map.get(emp),
        total_unidades_emp=total_unidades_map.get(_limpar_emp(emp)),
        sinal_sobre_renda=sinal_renda_map.get(_limpar_emp(emp)),
    )






# Referência ao próprio módulo (substitui import velocimetro nos blocos inline)
def _v_self():
    import sys
    return sys.modules[__name__]


def _v():
    return _v_self()

# =============================================================================
# PAINEL V2 (inline — ex-velocimetro_painel_v2.py)
# =============================================================================

# Painel v2 — metas Coordenadores + Canal, estoque analítico, velocímetros por coordenador.

import calendar
import html
import math
import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# Planilhas (somente metas)
# -----------------------------------------------------------------------------

FATORES_CANAL = {"RIO": 1.0, "DIR": 0.5, "PARC": 0.25, "RJ": 0.25, "TRI": 0.0}
CANAIS_META = ["RIO", "DIR", "PARC", "RJ", "TRI"]
TIPOS_INDICADOR = ("vendas", "agendamentos", "visitas")
TIPOS_META_COL = ("Desafio", "BP", "BP 70%")

COL_META_MAP = {
    ("vendas", "Desafio"): "Meta Vendas Desafio",
    ("vendas", "BP"): "Meta Vendas BP",
    ("vendas", "BP 70%"): "Meta Vendas BP 70%",
    ("agendamentos", "Desafio"): "Meta Agendamentos Desafio",
    ("agendamentos", "BP"): "Meta Agendamentos BP",
    ("agendamentos", "BP 70%"): "Meta Agendamentos BP 70%",
    ("visitas", "Desafio"): "Meta Visitas Desafio",
    ("visitas", "BP"): "Meta Visitas BP",
    ("visitas", "BP 70%"): "Meta Visitas BP 70%",
}

# Metas VGV na planilha de coordenadores — preferir colunas «Caixa Único» (não «Max»)
COL_META_VGV_MAP = {
    ("Desafio"): "Meta VGV Desafio (Caixa Único)",
    ("BP"): "Meta VGV BP (Caixa Único)",
    ("BP 70%"): "Meta VGV BP 70% (Caixa Único)",
}


def coluna_meta_vgv_coord(df_metas: pd.DataFrame, tipo_meta_col: str) -> str:
    """Resolve coluna VGV da meta (Caixa Unico preferido; ignora colunas Max)."""
    if df_metas is None or df_metas.empty:
        return ""
    preferida = COL_META_VGV_MAP.get(tipo_meta_col, "")
    if preferida and preferida in df_metas.columns:
        return preferida
    tipo_low = (tipo_meta_col or "").strip().lower()
    for col in df_metas.columns:
        cl = str(col).strip().lower()
        if "vgv" not in cl or tipo_low not in cl:
            continue
        if "max" in cl:
            continue
        if "caixa" in cl or "único" in cl or "unico" in cl:
            return str(col)
    for col in df_metas.columns:
        cl = str(col).strip().lower()
        if "vgv" in cl and tipo_low in cl and "max" not in cl:
            return str(col)
    return preferida if preferida in df_metas.columns else ""


def soma_meta_vgv_coord(
    df_metas: pd.DataFrame,
    mes: int,
    ano: int,
    tipo_meta_col: str,
    coordenadores: Optional[List[str]] = None,
    empreendimentos: Optional[List[str]] = None,
) -> float:
    if df_metas is None or df_metas.empty:
        return 0.0
    col = coluna_meta_vgv_coord(df_metas, tipo_meta_col)
    m = _filtrar_metas_mes_ano(df_metas, mes, ano)
    if coordenadores:
        m = m[m["Coordenador"].astype(str).str.strip().isin(coordenadores)]
    if empreendimentos:
        v = _v()
        emps_norm = {v._limpar_emp(e) for e in empreendimentos}
        m = m[m["Empreendimento"].map(lambda x: v._limpar_emp(x) in emps_norm)]
    total = 0.0
    if col and col in m.columns:
        total = float(pd.to_numeric(m[col], errors="coerce").fillna(0.0).sum())
    if total <= 0 and "Meta_VGV" in df_metas.columns:
        m2 = _filtrar_metas_mes_ano(df_metas, mes, ano)
        if coordenadores:
            m2 = m2[m2["Coordenador"].astype(str).str.strip().isin(coordenadores)]
        if empreendimentos:
            v = _v()
            emps_norm = {v._limpar_emp(e) for e in empreendimentos}
            m2 = m2[m2["Empreendimento"].map(lambda x: v._limpar_emp(x) in emps_norm)]
        vgv = float(pd.to_numeric(m2["Meta_VGV"], errors="coerce").fillna(0.0).sum())
        if vgv > 0:
            if tipo_meta_col == "BP":
                vgv *= 0.85
            elif tipo_meta_col == "BP 70%":
                vgv *= 0.7
            return vgv
    return total


def adaptar_metas_melt_para_coord(
    df_metas: pd.DataFrame,
    ano_meta: Optional[int] = None,
) -> pd.DataFrame:
    """Fallback: metas legado (melt) → schema da planilha Coordenadores Comerciais."""
    if df_metas is None or df_metas.empty:
        return pd.DataFrame()
    ano = int(ano_meta or date.today().year)
    out = df_metas.copy()
    if "Empreendimento" not in out.columns:
        for c in out.columns:
            if str(c).lower() in ("empreendimento", "obra"):
                out = out.rename(columns={c: "Empreendimento"})
                break
    if "Coordenador" not in out.columns:
        out["Coordenador"] = "Não Informado"
    if "Mes_Num" not in out.columns and "Mes" in out.columns:
        out["Mes_Num"] = out["Mes"].map(_parse_mes_num)
    elif "Mes_Num" in out.columns:
        out["Mes_Num"] = out["Mes_Num"].map(_parse_mes_num)
    out["Ano_Num"] = ano
    qtd = pd.to_numeric(
        out["Meta_Qtd"] if "Meta_Qtd" in out.columns else 0,
        errors="coerce",
    ).fillna(0.0)
    vgv = pd.to_numeric(
        out["Meta_VGV"] if "Meta_VGV" in out.columns else 0,
        errors="coerce",
    ).fillna(0.0)
    for (ind, tipo), col_name in COL_META_MAP.items():
        if col_name in out.columns:
            continue
        if ind == "vendas":
            if tipo == "Desafio":
                out[col_name] = qtd
            elif tipo == "BP":
                out[col_name] = qtd * 0.85
            elif tipo == "BP 70%":
                out[col_name] = qtd * 0.7
        else:
            out[col_name] = 0.0
    for tipo, col_name in COL_META_VGV_MAP.items():
        if col_name in out.columns:
            continue
        if tipo == "Desafio":
            out[col_name] = vgv
        elif tipo == "BP":
            out[col_name] = vgv * 0.85
        elif tipo == "BP 70%":
            out[col_name] = vgv * 0.7
    return out


def _metas_coord_tem_vgv_mes(df: pd.DataFrame, mes: int, ano: int) -> bool:
    """True se há meta VGV > 0 no mês/ano (colunas Caixa Único ou Meta_VGV legado)."""
    if df is None or df.empty:
        return False
    m = _filtrar_metas_mes_ano(df, mes, ano)
    if m.empty:
        return False
    for col in COL_META_VGV_MAP.values():
        if col in m.columns and float(pd.to_numeric(m[col], errors="coerce").fillna(0.0).sum()) > 0:
            return True
    if "Meta_VGV" in m.columns and float(pd.to_numeric(m["Meta_VGV"], errors="coerce").fillna(0.0).sum()) > 0:
        return True
    return False


def _metas_coord_tem_dados_mes(df: pd.DataFrame, mes: int, ano: int) -> bool:
    if df is None or df.empty:
        return False
    m = _filtrar_metas_mes_ano(df, mes, ano)
    if m.empty:
        return False
    for col in list(COL_META_VGV_MAP.values()) + [c for c in COL_META_MAP.values() if c.startswith("Meta Vendas")]:
        if col in m.columns and float(pd.to_numeric(m[col], errors="coerce").fillna(0.0).sum()) > 0:
            return True
    if "Meta_VGV" in m.columns and float(pd.to_numeric(m["Meta_VGV"], errors="coerce").fillna(0.0).sum()) > 0:
        return True
    if "Meta_Qtd" in m.columns and float(pd.to_numeric(m["Meta_Qtd"], errors="coerce").fillna(0.0).sum()) > 0:
        return True
    return False


def carregar_metas_coordenadores_com_fallback(
    cred_fp: str,
    df_metas_legacy: Optional[pd.DataFrame] = None,
    ano_meta: Optional[int] = None,
    mes_meta: Optional[int] = None,
) -> Tuple[pd.DataFrame, Optional[str]]:
    aviso: Optional[str] = None
    df_coord: Optional[pd.DataFrame] = None
    try:
        df_coord = carregar_metas_coordenadores(cred_fp)
        if df_coord is not None and not df_coord.empty:
            mes_ref = int(mes_meta or date.today().month)
            ano_ref = int(ano_meta or date.today().year)
            if _metas_coord_tem_vgv_mes(df_coord, mes_ref, ano_ref):
                return df_coord, None
            if df_metas_legacy is not None and not df_metas_legacy.empty:
                aviso = "Metas coordenadores sem VGV — usando planilha Metas legado."
                return adaptar_metas_melt_para_coord(df_metas_legacy, ano_meta), aviso
            if _metas_coord_tem_dados_mes(df_coord, mes_ref, ano_ref):
                return df_coord, "Metas coordenadores sem colunas VGV preenchidas."
            aviso = "Metas coordenadores sem valores para o mês/ano selecionado."
    except Exception as exc:
        aviso = str(exc)
    if df_metas_legacy is not None and not df_metas_legacy.empty:
        return adaptar_metas_melt_para_coord(df_metas_legacy, ano_meta), aviso
    if df_coord is not None and not df_coord.empty:
        return df_coord, aviso
    return pd.DataFrame(), aviso or "Planilha Metas Coordenadores vazia ou inacessível."

ALIASES_ESTOQUE_VFK = ["Valor Final com Kit", "ValorFinalComKit__c", "Valor Final Com Kit"]
ALIASES_ESTOQUE_AVAL = ["Valor de Avaliação Bancária", "Valor de Avaliação", "Valor_de_Avalia_o_Banc_ria__c"]
ALIASES_ESTOQUE_FOLGA = ["Valor Folga", "Valor_Folga__c", "Folga Comercial", "Folga_Comercial__c"]
ALIASES_ESTOQUE_BA = ["Bônus Adimplência", "Bonus Adimplencia", "B_nus_Adimpl_ncia__c"]
ALIASES_ESTOQUE_IDENT = ["Identificador", "Identificador__c"]
ALIASES_ESTOQUE_AREA = ["Area", "Área", "Area__c", "AreaComercial__c"]
ALIASES_ESTOQUE_HABITE = [
    "Habite-se", "Habite_se__c", "Previsão de expedição do habite-se",
    "DataExpedicaoHabitese__c", "Data Expedição Habite-se",
]
ALIASES_ESTOQUE_TIPOLOGIA = ["Tipologia__c", "Tipologia"]


def parse_identificador_unidade(ident: Any) -> Dict[str, str]:
    import re
    s = str(ident or "").strip().upper()
    bloco = ""
    m = re.search(r"BL\s*0*(\d+)", s)
    if m:
        bloco = m.group(1)
    elif "-" in s:
        bloco = s.split("-")[0].replace("BL", "").strip()
    andar = s[-4:-2] if len(s) >= 4 else ""
    return {"identificador": s, "bloco": bloco, "andar": andar}


@dataclass
class FiltrosPainelV2:
    data_ini: date
    data_fim: date
    mes_meta: int
    ano_meta: int
    tipo_indicador: str
    tipo_meta_col: str
    canal_meta: str
    coordenadores_sel: List[str]
    emps_sel: List[str]




def _parse_num_br(val: Any) -> float:
    v = _v()
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    return float(v.parse_valor_br(val))


def _sum_col_num(df: pd.DataFrame, col: str, default: float = 0.0) -> float:
    """Soma coluna numérica tolerando strings vindas do Google Sheets."""
    if df is None or df.empty or col not in df.columns:
        return default
    return float(pd.to_numeric(df[col], errors="coerce").fillna(0.0).sum())


def assegurar_metricas_vendas(df: pd.DataFrame) -> pd.DataFrame:
    """Garante tipos numéricos nas métricas de venda após leitura do cache Sheets."""
    if df is None or df.empty:
        return df
    out = df.copy()
    for col in ("_vgv", "_peso_coord", "_qtd_venda", "_vgv_venda"):
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce")
    if "_peso_coord" in out.columns:
        out["_peso_coord"] = out["_peso_coord"].fillna(1.0)
    else:
        out["_peso_coord"] = 1.0
    if "_vgv" in out.columns:
        out["_vgv"] = out["_vgv"].fillna(0.0)
    if "_qtd_venda" not in out.columns:
        vgv_base = out["_vgv"] if "_vgv" in out.columns else 0.0
        out["_qtd_venda"] = 1.0 * out["_peso_coord"]
        if "_vgv_venda" not in out.columns:
            out["_vgv_venda"] = vgv_base * out["_peso_coord"]
    else:
        out["_qtd_venda"] = out["_qtd_venda"].fillna(out["_peso_coord"])
    if "_vgv_venda" not in out.columns:
        vgv_base = out["_vgv"] if "_vgv" in out.columns else 0.0
        out["_vgv_venda"] = vgv_base * out["_peso_coord"]
    else:
        out["_vgv_venda"] = out["_vgv_venda"].fillna(0.0)
    for col in ("_mes", "_ano"):
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce")
    return out


@st.cache_data(ttl=300, show_spinner=False)
def carregar_metas_coordenadores(_cred_fp: str) -> pd.DataFrame:
    v = _v()
    df = v.ler_planilha_aba_df(SPREADSHEET_METAS_COORD_ID, WS_METAS_COORD, _cred_fp)
    df = v.normalizar_colunas(df)
    if df.empty:
        return df
    ren = {}
    for c in df.columns:
        cl = c.strip().lower()
        if cl in ("empreendimento", "obra"):
            ren[c] = "Empreendimento"
        elif cl == "coordenador":
            ren[c] = "Coordenador"
        elif cl in ("mês", "mes"):
            ren[c] = "Mes"
        elif cl == "ano":
            ren[c] = "Ano"
    df = df.rename(columns=ren)
    for col in df.columns:
        if col.startswith("Meta ") or col.startswith("Premiação") or col.startswith("Premiacao"):
            df[col] = df[col].map(_parse_num_br)
        if str(col).startswith("Meta VGV"):
            df[col] = df[col].map(_parse_num_br)
    if "Mes" in df.columns:
        df["Mes_Num"] = df["Mes"].map(_parse_mes_num)
    if "Ano" in df.columns:
        df["Ano_Num"] = pd.to_numeric(df["Ano"], errors="coerce").fillna(0).astype(int)
    return df


@st.cache_data(ttl=300, show_spinner=False)
def carregar_metas_canal(_cred_fp: str) -> pd.DataFrame:
    v = _v()
    df = v.ler_planilha_aba_df(
        SPREADSHEET_BASES_IVAN_ID, WS_CANAL, _cred_fp, aliases=WS_CANAL_ALIASES,
    )
    df = v.normalizar_colunas(df)
    if df.empty:
        return df
    ren = {}
    for c in df.columns:
        cl = c.strip().lower()
        if cl in ("mês", "mes"):
            ren[c] = "Mes"
        elif cl == "ano":
            ren[c] = "Ano"
        elif cl == "canal":
            ren[c] = "Canal"
        elif cl == "vgv":
            ren[c] = "VGV"
        elif cl == "vendas":
            ren[c] = "Vendas"
        elif "vgv real" in cl:
            ren[c] = "VGV_Real"
    df = df.rename(columns=ren)
    for col in ("VGV", "Vendas", "VGV_Real"):
        if col in df.columns:
            df[col] = df[col].map(_parse_num_br)
    if "Mes" in df.columns:
        df["Mes_Num"] = df["Mes"].map(_parse_mes_num)
    if "Ano" in df.columns:
        df["Ano_Num"] = pd.to_numeric(df["Ano"], errors="coerce").fillna(0).astype(int)
    if "Canal" in df.columns:
        df["Canal"] = df["Canal"].astype(str).str.strip().str.upper()
    return df


def _parse_mes_num(val: Any) -> int:
    """Converte mês numérico ou nome (ex.: 'Agosto', '8') para 1–12."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0
    s = str(val).strip().lower()
    if not s:
        return 0
    if s.isdigit():
        n = int(s)
        return n if 1 <= n <= 12 else 0
    meses = {
        "jan": 1, "janeiro": 1,
        "fev": 2, "fevereiro": 2,
        "mar": 3, "março": 3, "marco": 3,
        "abr": 4, "abril": 4,
        "mai": 5, "maio": 5,
        "jun": 6, "junho": 6,
        "jul": 7, "julho": 7,
        "ago": 8, "agosto": 8,
        "set": 9, "setembro": 9,
        "out": 10, "outubro": 10,
        "nov": 11, "novembro": 11,
        "dez": 12, "dezembro": 12,
    }
    if s in meses:
        return meses[s]
    for nome, num in meses.items():
        if s.startswith(nome):
            return num
    dig = re.search(r"(\d+)", s)
    if dig:
        n = int(dig.group(1))
        return n if 1 <= n <= 12 else 0
    return 0


def _filtrar_metas_mes_ano(df: pd.DataFrame, mes: int, ano: int) -> pd.DataFrame:
    """Filtra metas por mês/ano com fallback quando o recorte exato está vazio."""
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    m = df.copy()
    tem_mes = "Mes_Num" in m.columns
    tem_ano = "Ano_Num" in m.columns
    if tem_mes and tem_ano:
        exato = m[(m["Mes_Num"] == mes) & (m["Ano_Num"] == ano)]
        if not exato.empty:
            return exato
        por_mes = m[m["Mes_Num"] == mes]
        if not por_mes.empty:
            max_ano = int(pd.to_numeric(por_mes["Ano_Num"], errors="coerce").max())
            if max_ano > 0:
                return por_mes[por_mes["Ano_Num"] == max_ano]
            return por_mes
        por_ano = m[m["Ano_Num"] == ano]
        if not por_ano.empty:
            return por_ano
    elif tem_mes:
        por_mes = m[m["Mes_Num"] == mes]
        if not por_mes.empty:
            return por_mes
    return m


def coluna_meta_coord(tipo_indicador: str, tipo_meta_col: str) -> str:
    return COL_META_MAP.get((tipo_indicador, tipo_meta_col), "")


def _empreendimentos_para_analitico(
    filtros: "FiltrosPainelV2",
    mapa_coord: Dict[str, str],
    df_vendas: pd.DataFrame,
    df_metas_fallback: Optional[pd.DataFrame] = None,
) -> List[str]:
    """Lista empreendimentos para tabela analítica, com fallbacks se mapa de metas vazio."""
    v = _v()
    emps = list(filtros.emps_sel or [])
    if not emps:
        emps = sorted(set(mapa_coord.keys()))
    if not emps and df_metas_fallback is not None and not df_metas_fallback.empty:
        m = _filtrar_metas_mes_ano(df_metas_fallback, filtros.mes_meta, filtros.ano_meta)
        if "Empreendimento" in m.columns:
            emps = sorted(
                {v._limpar_emp(e) for e in m["Empreendimento"].dropna() if v._limpar_emp(e)}
            )
    if not emps and df_vendas is not None and not df_vendas.empty and "Empreendimento" in df_vendas.columns:
        emps = sorted(
            {v._limpar_emp(e) for e in df_vendas["Empreendimento"].dropna() if v._limpar_emp(e)}
        )
    if filtros.coordenadores_sel:
        coords = set(filtros.coordenadores_sel)
        if mapa_coord:
            emps = sorted(
                e for e in emps
                if mapa_coord.get(v._limpar_emp(e), "") in coords
            )
        elif df_metas_fallback is not None and not df_metas_fallback.empty:
            m = _filtrar_metas_mes_ano(df_metas_fallback, filtros.mes_meta, filtros.ano_meta)
            if "Coordenador" in m.columns and "Empreendimento" in m.columns:
                mask = m["Coordenador"].astype(str).str.strip().isin(coords)
                emps = sorted(
                    {v._limpar_emp(e) for e in m.loc[mask, "Empreendimento"].dropna() if v._limpar_emp(e)}
                )
    return emps


def mapa_emp_coordenador(df_metas: pd.DataFrame, mes: int, ano: int) -> Dict[str, str]:
    if df_metas is None or df_metas.empty:
        return {}
    m = _filtrar_metas_mes_ano(df_metas, mes, ano)
    out: Dict[str, str] = {}
    v = _v()
    for _, row in m.iterrows():
        emp = v._limpar_emp(row.get("Empreendimento"))
        coord = str(row.get("Coordenador") or "").strip()
        if emp and coord:
            out[emp] = coord
    return out


def meta_canal_vgv_vendas(
    df_canal: pd.DataFrame,
    mes: int,
    ano: int,
    canal: str,
) -> Tuple[float, float]:
    """Retorna (meta_vgv, meta_vendas) escalados pelo fator do canal sobre linha RIO."""
    if df_canal is None or df_canal.empty:
        return 0.0, 0.0
    base = _filtrar_metas_mes_ano(df_canal, mes, ano)
    if base.empty or "Canal" not in base.columns:
        return 0.0, 0.0
    rio = base[base["Canal"].astype(str).str.strip().str.upper() == "RIO"]
    if rio.empty:
        return 0.0, 0.0
    vgv_rio = float(rio["VGV"].sum()) if "VGV" in rio.columns else 0.0
    ven_rio = float(rio["Vendas"].sum()) if "Vendas" in rio.columns else 0.0
    fator = FATORES_CANAL.get((canal or "RIO").strip().upper(), 0.0)
    return vgv_rio * fator, ven_rio * fator


def soma_meta_coord(
    df_metas: pd.DataFrame,
    mes: int,
    ano: int,
    tipo_indicador: str,
    tipo_meta_col: str,
    coordenadores: Optional[List[str]] = None,
    empreendimentos: Optional[List[str]] = None,
) -> float:
    col = coluna_meta_coord(tipo_indicador, tipo_meta_col)
    if not col or df_metas is None or df_metas.empty or col not in df_metas.columns:
        return 0.0
    m = _filtrar_metas_mes_ano(df_metas, mes, ano)
    if coordenadores:
        m = m[m["Coordenador"].astype(str).str.strip().isin(coordenadores)]
    if empreendimentos:
        v = _v()
        emps_norm = {v._limpar_emp(e) for e in empreendimentos}
        m = m[m["Empreendimento"].map(lambda x: v._limpar_emp(x) in emps_norm)]
    return float(pd.to_numeric(m[col], errors="coerce").fillna(0.0).sum())


def escala_meta_regressao(
    proj: Optional[Dict[str, Any]],
    meta_mes: float,
    hoje: Optional[date] = None,
) -> Dict[str, float]:
    """Meta dia/semana/mês usando curva da regressão de vendas."""
    hoje = hoje or date.today()
    dias_mes = calendar.monthrange(hoje.year, hoje.month)[1]
    meta_mes = max(float(meta_mes), 0.0)
    linear_pace = meta_mes * hoje.day / dias_mes if dias_mes else 0.0
    if not proj or meta_mes <= 0:
        return {
            "meta_mes": meta_mes,
            "meta_dia": meta_mes / dias_mes if dias_mes else 0.0,
            "meta_semana": meta_mes * min(7, hoje.day) / dias_mes if dias_mes else 0.0,
            "meta_acum_hoje": linear_pace,
            "frac_hoje": hoje.day / dias_mes if dias_mes else 0.0,
        }
    pred_raw = proj.get("pred_reg_mes")
    if pred_raw is not None:
        pred = np.maximum(np.asarray(pred_raw, dtype=float), 0.0)
    else:
        ano, mes = hoje.year, hoje.month
        dias = [date(ano, mes, d) for d in range(1, dias_mes + 1)]
        v = _v()
        modelo = proj.get("modelo") or {}
        pred = np.maximum(
            np.asarray(v.prever_qtd_dias(modelo, dias, incluir_mes=proj.get("incluir_mes", True)), dtype=float),
            0.0,
        )
    total = float(pred.sum())
    if total <= 1e-9:
        frac = hoje.day / dias_mes if dias_mes else 0.0
    else:
        frac = float(pred[: hoje.day].sum()) / total
    meta_acum = meta_mes * frac
    pesos = pred / total if total > 1e-9 else np.ones(dias_mes) / max(dias_mes, 1)
    meta_dia = meta_mes * float(pesos[hoje.day - 1]) if hoje.day <= len(pesos) else 0.0
    ini_sem = max(1, hoje.day - hoje.weekday())
    meta_semana = meta_mes * float(pesos[ini_sem - 1 : hoje.day].sum())
    return {
        "meta_mes": meta_mes,
        "meta_dia": meta_dia,
        "meta_semana": meta_semana,
        "meta_acum_hoje": meta_acum,
        "frac_hoje": frac,
    }


def filtrar_vendas_painel_v2(
    df_vendas: pd.DataFrame,
    filtros: FiltrosPainelV2,
    col_contrato: str,
    col_canal: Optional[str] = None,
    aplicar_periodo: bool = True,
) -> pd.DataFrame:
    """Recorte de vendas alinhado aos filtros v2 (período, emp, canal)."""
    v = _v()
    base = df_vendas.copy()
    if filtros.emps_sel and "Empreendimento" in base.columns:
        emps = {v._limpar_emp(e) for e in filtros.emps_sel}
        base = base[base["Empreendimento"].map(v._limpar_emp).isin(emps)]
    if aplicar_periodo and col_contrato and col_contrato in base.columns:
        base = _filtrar_df_periodo(base, col_contrato, filtros.data_ini, filtros.data_fim)
    canal = (filtros.canal_meta or "RIO").strip().upper()
    if canal != "RIO" and not base.empty:
        mask = pd.Series(False, index=base.index)
        if canal == "DIR":
            mask |= base.get("Canal_Agrupado", pd.Series("", index=base.index)) == "DV RJ"
        elif canal == "PARC" and col_canal:
            mask |= base[col_canal].astype(str).str.upper().str.strip().apply(
                lambda x: x.split("-")[0].strip() == "RJG" or x == "RJG"
            )
        elif canal == "RJ" and col_canal:
            mask |= base[col_canal].astype(str).str.upper().str.strip().apply(
                lambda x: x.split("-")[0].strip() == "RJ" or x == "RJ"
            )
        elif canal == "TRI":
            mask = pd.Series(False, index=base.index)
        base = base[mask]
    return base


def empreendimentos_de_coord(
    mapa: Dict[str, str],
    coordenadores: List[str],
) -> List[str]:
    coords = {c.strip() for c in coordenadores}
    return sorted(emp for emp, c in mapa.items() if c in coords)


def realizado_vendas_periodo(
    df_vendas: pd.DataFrame,
    col_contrato: str,
    ini: date,
    fim: date,
    empreendimentos: Optional[List[str]] = None,
) -> Tuple[float, float]:
    v = _v()
    base = _filtrar_df_periodo(df_vendas, col_contrato, ini, fim)
    if empreendimentos and "Empreendimento" in base.columns:
        emps = {v._limpar_emp(e) for e in empreendimentos}
        base = base[base["Empreendimento"].map(v._limpar_emp).isin(emps)]
    qtd = (
        _sum_col_num(base, "_qtd_venda", float(len(base)))
        if "_qtd_venda" in base.columns
        else float(len(base))
    )
    vgv = _sum_col_num(base, "_vgv_venda", 0.0)
    return qtd, vgv


def realizado_funil_periodo(
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    df_vendas: pd.DataFrame,
    emp: str,
    ini: date,
    fim: date,
) -> Dict[str, float]:
    v = _v()
    return v.totais_funil_empreendimento(df_ag, df_pastas, df_vendas, emp, ini, fim)


def _preco_tabela_row(row: pd.Series, v) -> float:
    vfk = _parse_num_br(row.get(v.achar_coluna(pd.DataFrame([row]), ALIASES_ESTOQUE_VFK) or ""))
    if vfk <= 0:
        for a in ALIASES_ESTOQUE_VFK:
            if a in row.index:
                vfk = _parse_num_br(row.get(a))
                break
    ba = 0.0
    folga = 0.0
    for a in ALIASES_ESTOQUE_BA:
        if a in row.index:
            ba = _parse_num_br(row.get(a))
            break
    for a in ALIASES_ESTOQUE_FOLGA:
        if a in row.index:
            folga = _parse_num_br(row.get(a))
            break
    if vfk <= 0:
        cols = list(row.index)
        vfk = _parse_num_br(row.get(v.achar_coluna(pd.DataFrame(columns=cols), ALIASES_ESTOQUE_VFK) or ""))
    return max(vfk - ba - folga, 0.0)


def is_garden(identificador: Any) -> bool:
    s = str(identificador or "").strip().upper()
    if not s or len(s) < 4:
        return False
    andar = s[-4:-2] if len(s) >= 4 else ""
    return andar == "01"


def agregar_estoque(df_est: pd.DataFrame) -> Tuple[Dict[str, Any], pd.DataFrame]:
    """KPI global + dataframe enriquecido por unidade."""
    v = _v()
    if df_est is None or df_est.empty:
        return {"unidades": 0, "vgv": 0.0, "ticket": 0.0}, pd.DataFrame()
    df = df_est.copy()
    col_vfk = v.achar_coluna(df, ALIASES_ESTOQUE_VFK)
    col_av = v.achar_coluna(df, ALIASES_ESTOQUE_AVAL)
    col_folga = v.achar_coluna(df, ALIASES_ESTOQUE_FOLGA)
    col_ba = v.achar_coluna(df, ALIASES_ESTOQUE_BA)
    col_tip = v.achar_coluna(df, ALIASES_ESTOQUE_TIPOLOGIA)
    col_id = v.achar_coluna(df, ALIASES_ESTOQUE_IDENT)
    col_area = v.achar_coluna(df, ALIASES_ESTOQUE_AREA)
    col_hab = v.achar_coluna(df, ALIASES_ESTOQUE_HABITE)
    col_emp = v.achar_coluna(df, v.ALIASES_EMPREENDIMENTO) or "Empreendimento"

    def _col_val(row, col):
        return _parse_num_br(row.get(col)) if col and col in row.index else 0.0

    rows = []
    for _, r in df.iterrows():
        vfk = _col_val(r, col_vfk)
        ba = _col_val(r, col_ba)
        folga = _col_val(r, col_folga)
        aval = _col_val(r, col_av)
        preco_tab = max(vfk - ba - folga, 0.0)
        area = _col_val(r, col_area)
        ident = r.get(col_id) if col_id else ""
        parsed = parse_identificador_unidade(ident) if ident else {"identificador": "", "bloco": "", "andar": ""}
        tipologia = str(r.get(col_tip) or "").strip() if col_tip else ""
        hab = r.get(col_hab) if col_hab else None
        rows.append({
            "Empreendimento": v._limpar_emp(r.get(col_emp)),
            "Identificador": ident,
            "Bloco": parsed.get("bloco", ""),
            "Andar": parsed.get("andar", ""),
            "Tipologia": tipologia,
            "VFK": vfk,
            "BA": ba,
            "Folga": folga,
            "Avaliacao": aval,
            "PrecoTabela": preco_tab,
            "Area": area,
            "HabiteSe": hab,
            "Garden": is_garden(ident),
            "Desenquadrado": aval > preco_tab if preco_tab > 0 and aval > 0 else False,
            "AtoNecessario": max(0.0, preco_tab - 0.8 * aval) if aval > 0 else 0.0,
            "Investidor": str(r.get("Possui Investidor") or r.get("Possui_Investidor__c") or "").strip().upper()
            in ("TRUE", "1", "SIM", "YES"),
        })
    enr = pd.DataFrame(rows)
    vgv = float(enr["PrecoTabela"].sum()) if not enr.empty else 0.0
    n = len(enr)
    kpi = {"unidades": n, "vgv": vgv, "ticket": (vgv / n if n else 0.0)}
    return kpi, enr


def resumo_estoque_por_emp(enr: pd.DataFrame, hoje: Optional[date] = None) -> pd.DataFrame:
    v = _v()
    hoje = hoje or date.today()
    if enr is None or enr.empty:
        return pd.DataFrame()
    out_rows = []
    for emp, g in enr.groupby("Empreendimento"):
        n = len(g)
        gardens = int(g["Garden"].sum())
        pct_garden = gardens / n * 100.0 if n else 0.0
        hab_dates = []
        for h in g["HabiteSe"].dropna():
            dt = v.parse_data_serie(pd.Series([h])).iloc[0] if h is not None else pd.NaT
            if pd.notna(dt):
                hab_dates.append(dt.date() if hasattr(dt, "date") else dt)
        if hab_dates:
            hab_ref = min(hab_dates)
            if hab_ref <= hoje:
                hab_txt = "Pronto"
                meses_hab = 0.0
            else:
                hab_txt = f"{fmt_num((hab_ref - hoje).days / 30.4)} meses"
                meses_hab = (hab_ref - hoje).days / 30.4
        else:
            hab_txt = "—"
            meses_hab = 0.0
        vendas_mes_nec = (n / meses_hab) if meses_hab > 0.1 else 0.0
        desenq = float(g["Desenquadrado"].mean() * 100.0) if n else 0.0
        n_inv = int(g["Investidor"].sum()) if "Investidor" in g.columns else 0
        pct_inv = n_inv / n * 100.0 if n else 0.0
        out_rows.append({
            "Empreendimento": emp,
            "Unidades": n,
            "Diff_Avaliacao": float(g["Avaliacao"].sum() - g["PrecoTabela"].sum()),
            "m2_Total": float(g["Area"].sum()),
            "VGV_Tabela": float(g["PrecoTabela"].sum()),
            "m2_Medio": float(g["Area"].mean()) if n else 0.0,
            "Preco_m2_Medio": float(g["PrecoTabela"].sum() / g["Area"].sum()) if g["Area"].sum() > 0 else 0.0,
            "Pct_Garden": pct_garden,
            "HabiteSe": hab_txt,
            "Vendas_Mes_Necessarias": vendas_mes_nec,
            "Desenquadramento_Pct": desenq,
            "Ato_Necessario": float(g["AtoNecessario"].sum()),
            "Preco_Medio_Tabela": float(g["PrecoTabela"].mean()) if n else 0.0,
            "Pct_Tabela_Investidor": pct_inv,
            "Pct_Tabela_Direta": 100.0 - pct_inv,
        })
    return pd.DataFrame(out_rows)


def _metricas_vendas_avancadas(
    df_vendas: pd.DataFrame,
    emp: str,
    ini: date,
    fim: date,
    col_contrato: str,
) -> Dict[str, float]:
    v = _v()
    base = _filtrar_df_periodo(df_vendas, col_contrato, ini, fim)
    if base.empty or "Empreendimento" not in base.columns:
        return {"vendas_futuras": 0.0, "vendas_comunicadas": 0.0}
    base = base[base["Empreendimento"].map(v._limpar_emp) == v._limpar_emp(emp)]
    vf = 0.0
    vc = 0.0
    for col in ("Venda futura", "Venda_Futura__c", "VendaFutura__c"):
        if col in base.columns:
            vf = float(base[col].astype(str).str.upper().isin(("TRUE", "1", "SIM", "YES")).sum())
            break
    for col in ("Venda comunicada", "GeradoComunicadoVenda__c", "VendaComunicadaAutomaticamente__c"):
        if col in base.columns:
            vc = float(base[col].astype(str).str.upper().isin(("TRUE", "1", "SIM", "YES")).sum())
            break
    return {"vendas_futuras": vf, "vendas_comunicadas": vc}


def _metricas_renda_90d(
    df_pastas: pd.DataFrame,
    emp: str,
    fim: date,
) -> float:
    v = _v()
    if df_pastas is None or df_pastas.empty:
        return 0.0
    ini = fim - timedelta(days=90)
    col_e = v.achar_coluna(df_pastas, v.ALIASES_EMPREENDIMENTO)
    col_d = v.achar_coluna_primeiro_envio_analise(df_pastas) or v.achar_coluna(
        df_pastas, v.ALIASES_DATA_CRIACAO
    )
    if not col_e or not col_d:
        return 0.0
    sub = _filtrar_df_periodo(df_pastas, col_d, ini, fim)
    sub = sub[sub[col_e].map(v._limpar_emp) == v._limpar_emp(emp)]
    col_renda = None
    for c in ("Valor da Renda", "Renda", "Valor_da_Renda__c", "Renda__c"):
        if c in sub.columns:
            col_renda = c
            break
    if not col_renda or sub.empty:
        return 0.0
    vals = sub[col_renda].map(_parse_num_br)
    vals = vals[vals > 0]
    return float(vals.mean()) if not vals.empty else 0.0


def _metricas_cotacoes_emp(
    df_cot: pd.DataFrame,
    emp: str,
    ini: date,
    fim: date,
) -> Dict[str, float]:
    v = _v()
    if df_cot is None or df_cot.empty:
        return {"vcx": 0.0, "pct_pro_soluto": 0.0, "pct_fluxo_escalonado": 0.0}
    col_e = v.achar_coluna(df_cot, v.ALIASES_EMPREENDIMENTO)
    col_d = v.achar_coluna(df_cot, v.ALIASES_DATA_CRIACAO)
    if not col_e:
        return {"vcx": 0.0, "pct_pro_soluto": 0.0, "pct_fluxo_escalonado": 0.0}
    sub = df_cot[df_cot[col_e].map(v._limpar_emp) == v._limpar_emp(emp)]
    if col_d:
        sub = _filtrar_df_periodo(sub, col_d, ini, fim)
    vcx = 0.0
    for col in ("Volta ao caixa", "VoltaAoCaixa__c"):
        if col in sub.columns:
            vcx = float(sub[col].astype(str).str.upper().isin(("TRUE", "1", "SIM", "YES")).sum())
            break
    pct_ps = 0.0
    col_ps = None
    for col in ("Percentual Pro Soluto", "PercentualdoProSoluto__c"):
        if col in sub.columns:
            col_ps = col
            break
    if col_ps and not sub.empty:
        pct_ps = float(sub[col_ps].map(_parse_num_br).mean())
    tem_ps = 0.0
    for col in ("Tem Pro Soluto", "TemProSoluto__c"):
        if col in sub.columns and not sub.empty:
            tem_ps = float(sub[col].astype(str).str.upper().isin(("TRUE", "1", "SIM", "YES")).mean() * 100.0)
            break
    return {"vcx": vcx, "pct_pro_soluto": pct_ps, "pct_fluxo_escalonado": tem_ps}


def montar_tabela_analitica(
    empreendimentos: List[str],
    df_est_enr: pd.DataFrame,
    resumo_est: pd.DataFrame,
    df_vendas: pd.DataFrame,
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    df_metas_coord: pd.DataFrame,
    filtros: FiltrosPainelV2,
    col_contrato: str,
    proj: Optional[Dict[str, Any]],
    df_cotacoes: Optional[pd.DataFrame] = None,
    df_pastas_aprov: Optional[pd.DataFrame] = None,
    df_tabela_comp: Optional[pd.DataFrame] = None,
    estoque_map: Optional[Dict[str, Dict[str, int]]] = None,
    total_unidades_por_emp: Optional[Dict[str, int]] = None,
) -> pd.DataFrame:
    v = _v()
    rows = []
    mapa_coord = mapa_emp_coordenador(df_metas_coord, filtros.mes_meta, filtros.ano_meta)
    sinal_renda_map = calcular_sinal_sobre_renda_por_emp(
        df_cotacoes, df_pastas, filtros.data_ini, filtros.data_fim,
    )
    estoque_map = estoque_map or {}
    lib_map = metricas_liberacao_estoque_por_emp(estoque_map, total_unidades_por_emp or {})
    for emp in empreendimentos:
        emp_c = v._limpar_emp(emp)
        meta_m = soma_meta_coord(
            df_metas_coord, filtros.mes_meta, filtros.ano_meta,
            filtros.tipo_indicador, filtros.tipo_meta_col,
            empreendimentos=[emp_c],
        )
        esc = escala_meta_regressao(proj, meta_m)
        funil = realizado_funil_periodo(
            df_ag, df_pastas, df_vendas, emp_c,
            filtros.data_ini, filtros.data_fim,
        )
        ven_qtd, ven_vgv = realizado_vendas_periodo(
            df_vendas, col_contrato, filtros.data_ini, filtros.data_fim, [emp_c],
        )
        res_est = resumo_est.loc[resumo_est["Empreendimento"] == emp_c] if not resumo_est.empty else pd.DataFrame()
        base_est = res_est.iloc[0].to_dict() if not res_est.empty else {}
        conv_ag_vis = v.taxa_conversao(funil.get("agendamentos", 0), funil.get("visitas", 0))
        conv_vis_pas = v.taxa_conversao(funil.get("visitas", 0), funil.get("pastas", 0))
        conv_pas_ap = v.taxa_conversao(funil.get("pastas", 0), funil.get("pastas_aprovadas", 0))
        conv_ap_ven = v.taxa_conversao(funil.get("pastas_aprovadas", 0), funil.get("vendas", 0))
        preco_venda = 0.0
        ven_emp = _filtrar_df_periodo(df_vendas, col_contrato, filtros.data_ini, filtros.data_fim)
        if not ven_emp.empty and "Empreendimento" in ven_emp.columns:
            ve = ven_emp[ven_emp["Empreendimento"].map(v._limpar_emp) == emp_c]
            col_val = v.achar_coluna(ve, ["Valor Real de Venda", "Valor Real", "Valor"])
            if col_val and not ve.empty:
                preco_venda = float(ve[col_val].map(_parse_num_br).mean())
        meta_ano = soma_meta_coord(
            df_metas_coord, filtros.mes_meta, filtros.ano_meta,
            "vendas", filtros.tipo_meta_col, empreendimentos=[emp_c],
        )
        meta_ano *= 12 / max(filtros.mes_meta, 1)
        ven_ano_ini = date(filtros.ano_meta, 1, 1)
        ven_ano_fim = date(filtros.ano_meta, 12, 31)
        ven_ano, _ = realizado_vendas_periodo(df_vendas, col_contrato, ven_ano_ini, ven_ano_fim, [emp_c])
        ven_adv = _metricas_vendas_avancadas(df_vendas, emp_c, filtros.data_ini, filtros.data_fim, col_contrato)
        renda_90 = _metricas_renda_90d(df_pastas, emp_c, filtros.data_fim)
        cot = _metricas_cotacoes_emp(df_cotacoes, emp_c, filtros.data_ini, filtros.data_fim)
        res_pc = calcular_resumo_ineficiencia_emp(
            df_pastas_aprov if df_pastas_aprov is not None else pd.DataFrame(),
            df_est_enr,
            df_vendas,
            df_tabela_comp if df_tabela_comp is not None else pd.DataFrame(),
            emp_c,
        )
        lib = lib_map.get(emp_c, {})
        ratio_sr = sinal_renda_map.get(emp_c)
        row = {
            "Empreendimento": emp_c,
            "Coordenador": mapa_coord.get(emp_c, ""),
            "Unidades_Estoque": base_est.get("Unidades", 0),
            "Diff_Avaliacao": base_est.get("Diff_Avaliacao", 0),
            "VGV_Estoque": base_est.get("VGV_Tabela", 0),
            "Pct_Garden": base_est.get("Pct_Garden", 0),
            "HabiteSe": base_est.get("HabiteSe", "—"),
            "Vendas_Mes_Nec": base_est.get("Vendas_Mes_Necessarias", 0),
            "Desenquadramento_Pct": base_est.get("Desenquadramento_Pct", 0),
            "Ato_Necessario": base_est.get("Ato_Necessario", 0),
            "Pct_Tabela_Direta": base_est.get("Pct_Tabela_Direta", 0),
            "Pct_Tabela_Investidor": base_est.get("Pct_Tabela_Investidor", 0),
            "Vendas_Realizadas": ven_qtd,
            "Vendas_Futuras": ven_adv["vendas_futuras"],
            "Vendas_Comunicadas": ven_adv["vendas_comunicadas"],
            "VCX": cot["vcx"],
            "Pct_Pro_Soluto": cot["pct_pro_soluto"],
            "Pct_Fluxo_Escalonado": cot["pct_fluxo_escalonado"],
            "Renda_Media_90d": renda_90,
            "Pastas_Aprovadas": res_pc.pastas_aprovadas,
            "Pastas_PC_Suficiente": res_pc.pastas_pc_suficiente,
            "Ineficiencia_Qtd": res_pc.ineficiencia_qtd,
            "Ineficiencia_Pct": res_pc.ineficiencia_pct,
            "Meta_Dia": esc["meta_dia"],
            "Meta_Semana": esc["meta_semana"],
            "Meta_Mes": meta_m,
            "Pct_Meta_Dia": (ven_qtd / esc["meta_dia"] * 100) if esc["meta_dia"] > 0 else 0,
            "Pct_Meta_Semana": (ven_qtd / esc["meta_semana"] * 100) if esc["meta_semana"] > 0 else 0,
            "Pct_Meta_Mes": (ven_qtd / meta_m * 100) if meta_m > 0 else 0,
            "Agendamentos": funil.get("agendamentos", 0),
            "Visitas": funil.get("visitas", 0),
            "Pastas": funil.get("pastas", 0),
            "Pastas_Aprov": funil.get("pastas_aprovadas", 0),
            "Conv_Ag_Vis": conv_ag_vis,
            "Conv_Vis_Pas": conv_vis_pas,
            "Conv_Pas_Ap": conv_pas_ap,
            "Conv_Ap_Ven": conv_ap_ven,
            "Preco_Medio_Venda": preco_venda,
            "Preco_Medio_Tabela": base_est.get("Preco_Medio_Tabela", 0),
            "Meta_Ano": meta_ano,
            "Vendido_Ano": ven_ano,
            "Unidades_Disponiveis": int(lib.get("disponivel", 0)),
            "Unidades_Liberadas": int(lib.get("liberadas", 0)),
            "Unidades_Total_SF": int(lib.get("total", 0)),
            "Pct_Disp_Liberadas": round(float(lib.get("pct_disp_liberadas", 0.0)), 1),
            "Pct_Liberadas_Total": round(float(lib.get("pct_liberadas_total", 0.0)), 1),
            "Sinal_Sobre_Renda_Pct": round(ratio_sr * 100.0, 1) if ratio_sr is not None else None,
        }
        rows.append(row)
    return pd.DataFrame(rows)


def render_filtros_painel_v2(
    df_metas_coord: pd.DataFrame,
    cred_fp: str,
    filtros_externos: Optional["FiltrosGlobais"] = None,
) -> FiltrosPainelV2:
    if filtros_externos is not None:
        return filtros_glob_to_v2(filtros_externos)
    hoje = date.today()
    ini_mes = date(hoje.year, hoje.month, 1)
    coords = _opcoes_unicas(
        _serie_coluna(df_metas_coord, "Coordenador"),
    )
    emps = _opcoes_unicas(
        _serie_coluna(df_metas_coord, "Empreendimento"),
    )
    st.markdown("#### Filtros de análise")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        data_ini = st.date_input("Data inicial", value=ini_mes, key="v2_data_ini")
    with c2:
        data_fim = st.date_input("Data final", value=hoje, key="v2_data_fim")
    with c3:
        mes_meta = st.selectbox("Mês meta", list(range(1, 13)), index=hoje.month - 1, key="v2_mes_meta")
    with c4:
        ano_meta = st.number_input("Ano meta", min_value=2020, max_value=2035, value=hoje.year, key="v2_ano_meta")
    c5, c6, c7, c8 = st.columns(4)
    with c5:
        tipo_indicador = st.selectbox("Indicador", TIPOS_INDICADOR, key="v2_tipo_ind")
    with c6:
        tipo_meta_col = st.selectbox("Tipo meta", TIPOS_META_COL, key="v2_tipo_meta")
    with c7:
        canal_meta = st.selectbox("Canal (velocímetro principal)", CANAIS_META, key="v2_canal")
    with c8:
        coords_sel = st.multiselect("Coordenadores", coords, default=coords, key="v2_coords")
    emps_sel = st.multiselect("Empreendimentos (opcional)", emps, default=[], key="v2_emps")
    return FiltrosPainelV2(
        data_ini=data_ini,
        data_fim=data_fim,
        mes_meta=int(mes_meta),
        ano_meta=int(ano_meta),
        tipo_indicador=tipo_indicador,
        tipo_meta_col=tipo_meta_col,
        canal_meta=canal_meta,
        coordenadores_sel=coords_sel,
        emps_sel=emps_sel,
    )


def render_kpi_estoque(kpi: Dict[str, Any]) -> None:
    v = _v()
    st.markdown(
        f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Unidades em estoque</div>
            <div class="val">{int(kpi.get('unidades', 0))}</div></div>
            <div class="vel-kpi"><div class="lbl">VGV em estoque</div>
            <div class="val">{v.fmt_br_milhoes(float(kpi.get('vgv', 0)))}</div></div>
            <div class="vel-kpi"><div class="lbl">Ticket médio do estoque</div>
            <div class="val">{v.fmt_br_milhoes(float(kpi.get('ticket', 0)))}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_metas_vendas(
    meta_comercial: float,
    meta_bp: float,
    meta_bp70: float,
    qtd_vendas_mes: float,
    vgv_vendas_mes: float,
) -> None:
    """Cards de meta (linha 2) e vendas do mês (linha 3)."""
    v = _v()
    st.markdown(
        f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Meta comercial (Desafio VGV)</div>
            <div class="val">{v.fmt_br_milhoes(meta_comercial)}</div></div>
            <div class="vel-kpi"><div class="lbl">Meta BP (VGV)</div>
            <div class="val">{v.fmt_br_milhoes(meta_bp)}</div></div>
            <div class="vel-kpi"><div class="lbl">Meta BP 70% (VGV)</div>
            <div class="val">{v.fmt_br_milhoes(meta_bp70)}</div></div>
        </div>
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Quantidade de vendas do mês</div>
            <div class="val">{v.fmt_qtd(qtd_vendas_mes)}</div></div>
            <div class="vel-kpi"><div class="lbl">VGV de vendas do mês</div>
            <div class="val">{v.fmt_br_milhoes(vgv_vendas_mes)}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_resumo_painel(
    kpi_estoque: Dict[str, Any],
    df_metas_coord: pd.DataFrame,
    df_vendas: pd.DataFrame,
    col_contrato: Optional[str],
    mes_meta: int,
    ano_meta: int,
    emps_sel: Optional[List[str]] = None,
) -> None:
    """Três linhas de KPI no topo: estoque, metas e vendas do mês."""
    render_kpi_estoque(kpi_estoque)
    meta_com = soma_meta_vgv_coord(
        df_metas_coord, mes_meta, ano_meta, "Desafio", empreendimentos=emps_sel or None,
    )
    meta_bp = soma_meta_vgv_coord(
        df_metas_coord, mes_meta, ano_meta, "BP", empreendimentos=emps_sel or None,
    )
    meta_bp70 = soma_meta_vgv_coord(
        df_metas_coord, mes_meta, ano_meta, "BP 70%", empreendimentos=emps_sel or None,
    )
    hoje = date.today()
    ini_mes = date(ano_meta if ano_meta else hoje.year, mes_meta if mes_meta else hoje.month, 1)
    fim_mes = min(hoje, date(ini_mes.year, ini_mes.month, calendar.monthrange(ini_mes.year, ini_mes.month)[1]))
    qtd_mes, vgv_mes = realizado_vendas_periodo(
        df_vendas, col_contrato or "", ini_mes, fim_mes, emps_sel or None,
    )
    render_kpi_metas_vendas(meta_com, meta_bp, meta_bp70, qtd_mes, vgv_mes)


def render_velocimetro_principal(
    filtros: FiltrosPainelV2,
    df_canal: pd.DataFrame,
    df_metas_coord: pd.DataFrame,
    realizado_qtd: float,
    realizado_vgv: float,
    meta_qtd: float,
    meta_vgv: float,
    esc: Dict[str, float],
) -> None:
    v = _v()
    st.subheader("Velocímetro principal")
    titulo = f"VGV vendido / Meta Desafio · {filtros.canal_meta}"
    emps = filtros.emps_sel or None
    meta_vgv_desafio = soma_meta_vgv_coord(
        df_metas_coord, filtros.mes_meta, filtros.ano_meta, "Desafio",
        empreendimentos=emps,
    )
    if meta_vgv_desafio <= 0:
        meta_vgv_desafio = meta_vgv
    if filtros.tipo_indicador == "vendas":
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            v.criar_medidor(
                titulo, realizado_qtd, meta_qtd, realizado_vgv, meta_vgv_desafio, realizado_qtd,
                mostrar_vgv=True, metrica="vgv",
            )
    else:
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            v.criar_medidor(
                f"{filtros.tipo_indicador.title()} · {filtros.tipo_meta_col}",
                realizado_qtd, meta_qtd, 0.0, 0.0, realizado_qtd,
                mostrar_vgv=False,
            )
    st.caption(
        f"Meta acumulada (regressão): {v.fmt_qtd(esc['meta_acum_hoje'])} · "
        f"Meta dia: {v.fmt_qtd(esc['meta_dia'])} · "
        f"Meta semana: {v.fmt_qtd(esc['meta_semana'])} · "
        f"VGV meta Desafio: {v.fmt_br_milhoes(meta_vgv_desafio)}"
    )


def render_velocimetros_coordenador(
    filtros: FiltrosPainelV2,
    df_metas_coord: pd.DataFrame,
    df_vendas: pd.DataFrame,
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    col_contrato: str,
    mapa_coord: Dict[str, str],
) -> None:
    v = _v()
    st.subheader("Por coordenador")
    coords = filtros.coordenadores_sel or sorted(set(mapa_coord.values()))
    if not coords:
        st.info("Nenhum coordenador na planilha de metas.")
        return
    cols = st.columns(min(3, len(coords)) or 1)
    for i, coord in enumerate(coords):
        emps = empreendimentos_de_coord(mapa_coord, [coord])
        if filtros.emps_sel:
            emps = [e for e in emps if e in {v._limpar_emp(x) for x in filtros.emps_sel}]
        meta = soma_meta_coord(
            df_metas_coord, filtros.mes_meta, filtros.ano_meta,
            filtros.tipo_indicador, filtros.tipo_meta_col,
            coordenadores=[coord], empreendimentos=emps or None,
        )
        meta_vgv_coord = soma_meta_vgv_coord(
            df_metas_coord, filtros.mes_meta, filtros.ano_meta, "Desafio",
            coordenadores=[coord], empreendimentos=emps or None,
        )
        if filtros.tipo_indicador == "vendas":
            qtd, vgv = realizado_vendas_periodo(
                df_vendas, col_contrato, filtros.data_ini, filtros.data_fim, emps or None,
            )
            with cols[i % len(cols)]:
                v.criar_medidor(
                    coord, qtd, meta, vgv, meta_vgv_coord, qtd,
                    mostrar_vgv=True, metrica="vgv",
                )
        else:
            tot = 0.0
            for emp in (emps or []):
                f = realizado_funil_periodo(
                    df_ag, df_pastas, df_vendas, emp,
                    filtros.data_ini, filtros.data_fim,
                )
                tot += float(f.get(filtros.tipo_indicador, 0))
            with cols[i % len(cols)]:
                v.criar_medidor(coord, tot, meta, 0.0, 0.0, tot, mostrar_vgv=False)


OPCOES_ORDENACAO_ANALITICO = [
    "Alfabética (A → Z)",
    "Alfabética (Z → A)",
    "Valor crescente (1º campo Y)",
    "Valor decrescente (1º campo Y)",
]

CORES_GRAFICO_ANALITICO = (
    "#2563eb", "#cb0935", "#0f766e", "#d97706", "#7c3aed",
    "#0891b2", "#be123c", "#4d7c0f", "#9333ea", "#0369a1",
)


def _colunas_y_grafico_analitico(df: pd.DataFrame) -> List[str]:
    """Colunas numéricas disponíveis como métricas do eixo Y."""
    excluir = {"Empreendimento", "Coordenador", "HabiteSe"}
    out: List[str] = []
    for c in df.columns:
        if c in excluir:
            continue
        if pd.api.types.is_numeric_dtype(df[c]):
            out.append(c)
            continue
        if pd.to_numeric(df[c], errors="coerce").notna().any():
            out.append(c)
    return out


def _coluna_agregacao_media_analitico(col: str) -> bool:
    cl = str(col).lower()
    return any(
        tok in cl
        for tok in ("pct", "conv", "desenq", "preco", "renda", "diff", "ato", "sinal")
    )


def _agregar_analitico_por_eixo(df: pd.DataFrame, eixo_x: str) -> pd.DataFrame:
    """Uma linha por empreendimento ou agregação por coordenador."""
    if eixo_x == "Empreendimento":
        out = df.copy()
        out["Empreendimento"] = out["Empreendimento"].astype(str).str.strip()
        return out[out["Empreendimento"].ne("") & out["Empreendimento"].ne("—")]
    cols_num = _colunas_y_grafico_analitico(df)
    agg = {c: ("mean" if _coluna_agregacao_media_analitico(c) else "sum") for c in cols_num}
    base = df.copy()
    base["Coordenador"] = base["Coordenador"].fillna("").astype(str).str.strip()
    base.loc[base["Coordenador"].eq(""), "Coordenador"] = "—"
    return base.groupby("Coordenador", as_index=False).agg(agg)


def _ordenar_dados_grafico_analitico(
    df: pd.DataFrame,
    eixo_x: str,
    ordenacao: str,
    primeiro_y: Optional[str],
) -> pd.DataFrame:
    out = df.copy()
    if ordenacao == "Alfabética (A → Z)":
        return out.sort_values(eixo_x, ascending=True, kind="stable")
    if ordenacao == "Alfabética (Z → A)":
        return out.sort_values(eixo_x, ascending=False, kind="stable")
    if primeiro_y and primeiro_y in out.columns:
        asc = ordenacao.startswith("Valor crescente")
        return out.sort_values(primeiro_y, ascending=asc, kind="stable")
    return out


def _montar_fig_grafico_analitico(
    df: pd.DataFrame,
    eixo_x: str,
    campos_y: List[str],
    eixos_por_campo: Dict[str, str],
    ordenacao: str,
    tipo_grafico: str = "Barras",
) -> go.Figure:
    plot_df = _agregar_analitico_por_eixo(df, eixo_x)
    plot_df = _ordenar_dados_grafico_analitico(
        plot_df, eixo_x, ordenacao, campos_y[0] if campos_y else None,
    )
    categorias = plot_df[eixo_x].astype(str).tolist()
    rotulos = {c: _rotulo_coluna_tabela(c) for c in campos_y}

    usa_esquerdo = any(eixos_por_campo.get(c, "Esquerdo") == "Esquerdo" for c in campos_y)
    usa_direito = any(eixos_por_campo.get(c) == "Direito" for c in campos_y)

    fig = go.Figure()
    for i, campo in enumerate(campos_y):
        vals = pd.to_numeric(plot_df[campo], errors="coerce").fillna(0.0)
        nome = rotulos[campo]
        cor = CORES_GRAFICO_ANALITICO[i % len(CORES_GRAFICO_ANALITICO)]
        eixo_lado = eixos_por_campo.get(campo, "Esquerdo")
        yaxis = "y2" if eixo_lado == "Direito" else "y"
        if tipo_grafico == "Linhas":
            fig.add_trace(go.Scatter(
                x=categorias, y=vals, name=nome, mode="lines+markers",
                line=dict(color=cor, width=2), marker=dict(size=7, color=cor), yaxis=yaxis,
            ))
        else:
            fig.add_trace(go.Bar(
                x=categorias, y=vals, name=nome, marker_color=cor, yaxis=yaxis,
            ))

    layout: Dict[str, Any] = dict(
        title=dict(
            text=f"Analítico — {_rotulo_coluna_tabela(eixo_x) if eixo_x in ROTULOS_COLUNAS_TABELA else eixo_x}",
            font=dict(family="Inter", color=COR_TEXTO_PRETO),
        ),
        xaxis=dict(
            title=eixo_x,
            tickangle=-45 if len(categorias) > 6 else 0,
            tickfont=dict(family="Inter", color=COR_TEXTO_PRETO),
        ),
        barmode="group",
        height=max(440, 100 + len(categorias) * 22),
        margin=dict(l=60, r=80, t=80, b=120 if len(categorias) > 6 else 80),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        font=dict(family="Inter", color=COR_TEXTO_PRETO),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    if usa_esquerdo:
        layout["yaxis"] = dict(
            title="Esquerdo",
            showgrid=True,
            gridcolor="rgba(226,232,240,0.5)",
            tickfont=dict(family="Inter", color=COR_TEXTO_PRETO),
        )
    else:
        layout["yaxis"] = dict(visible=False)
    if usa_direito:
        layout["yaxis2"] = dict(
            title="Direito",
            overlaying="y",
            side="right",
            showgrid=False,
            tickfont=dict(family="Inter", color=COR_TEXTO_PRETO),
        )
    fig.update_layout(**layout)
    return fig


def _exportar_fig_plotly(fig: go.Figure) -> Tuple[Optional[bytes], str]:
    """Gera bytes PNG (se kaleido disponível) e HTML do gráfico."""
    html_str = fig.to_html(include_plotlyjs="cdn", full_html=True)
    png_bytes: Optional[bytes] = None
    try:
        import plotly.io as pio
        altura = int(fig.layout.height or 600)
        png_bytes = pio.to_image(fig, format="png", width=1400, height=altura, scale=2)
    except Exception:
        pass
    return png_bytes, html_str


@_st_dialog_decorator("Gráfico personalizado — Analítico por empreendimento", width="large")
def dialog_grafico_analitico(df: pd.DataFrame) -> None:
    if df is None or df.empty:
        st.warning("Sem dados para montar o gráfico.")
        return

    cols_y = _colunas_y_grafico_analitico(df)
    if not cols_y:
        st.warning("Não há colunas numéricas disponíveis para o eixo Y.")
        return

    rotulos = {c: _rotulo_coluna_tabela(c) for c in cols_y}

    c1, c2 = st.columns(2)
    with c1:
        eixo_x = st.selectbox("Eixo X", ["Empreendimento", "Coordenador"], key="dlg_analit_eixo_x")
    with c2:
        tipo_grafico = st.selectbox("Tipo de gráfico", ["Barras", "Linhas"], key="dlg_analit_tipo")

    campos_y = st.multiselect(
        "Campos do eixo Y",
        options=cols_y,
        format_func=lambda c: rotulos[c],
        key="dlg_analit_campos_y",
    )

    eixos_por_campo: Dict[str, str] = {}
    if campos_y:
        st.markdown("**Alinhamento do eixo Y por campo**")
        for campo in campos_y:
            eixos_por_campo[campo] = st.radio(
                rotulos[campo],
                ["Esquerdo", "Direito"],
                horizontal=True,
                key=f"dlg_analit_eixo_y_{campo}",
            )

    ordenacao = st.selectbox(
        "Ordenação",
        OPCOES_ORDENACAO_ANALITICO,
        key="dlg_analit_ordem",
    )

    if not campos_y:
        st.info("Selecione ao menos um campo para o eixo Y.")
        return

    if st.button("Gerar gráfico", type="primary", key="dlg_analit_gerar"):
        fig = _montar_fig_grafico_analitico(
            df, eixo_x, campos_y, eixos_por_campo, ordenacao, tipo_grafico,
        )
        st.session_state["analit_grafico_fig"] = fig

    fig = st.session_state.get("analit_grafico_fig")
    if fig is None:
        return

    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": True})
    png_bytes, html_str = _exportar_fig_plotly(fig)
    d1, d2 = st.columns(2)
    with d1:
        if png_bytes:
            st.download_button(
                "Baixar PNG",
                data=png_bytes,
                file_name="grafico_analitico.png",
                mime="image/png",
                key="dlg_analit_dl_png",
                use_container_width=True,
            )
        else:
            st.caption("PNG indisponível (instale `kaleido` para exportar imagem).")
    with d2:
        st.download_button(
            "Baixar HTML",
            data=html_str,
            file_name="grafico_analitico.html",
            mime="text/html",
            key="dlg_analit_dl_html",
            use_container_width=True,
        )


def render_tabela_analitica(df: pd.DataFrame) -> None:
    c_tit, c_btn = st.columns([5, 1])
    with c_tit:
        st.subheader("Analítico por empreendimento")
        st.caption(
            "Consolidado: ineficiência, hipereficiência, tempos de conversão, shares, "
            "sinal/renda, estoque liberado, PS/VGV, sinais/VGV, VSO, metas, radar e MTD."
        )
    with c_btn:
        if st.button(
            "Gráfico personalizado",
            key="btn_grafico_analitico",
            use_container_width=True,
            disabled=df is None or df.empty,
        ):
            st.session_state.pop("analit_grafico_fig", None)
            dialog_grafico_analitico(df)
    if df is None or df.empty:
        st.info("Sem dados para a tabela analítica.")
        return

    disp = _ordenar_empreendimento_primeiro(preparar_df_tabela_exibicao(df))
    col_desenq = ROTULOS_COLUNAS_TABELA["Desenquadramento_Pct"]
    styler = _styler_desenquadramento(disp, col_desenq)
    _exibir_dataframe_preparada(disp, styler=styler)


def render_perfil_vendas_mtd(vendas_f: pd.DataFrame) -> None:
    """KPIs de vendas facilitadas vs normais no período filtrado."""
    st.subheader("Perfil das Vendas")
    qtd_facilitada = (
        _sum_col_num(vendas_f[vendas_f["Tipo_Venda"] == "Facilitada"], "_qtd_venda", 0.0)
        if "Tipo_Venda" in vendas_f.columns and "_qtd_venda" in vendas_f.columns
        else 0.0
    )
    qtd_normal = (
        _sum_col_num(vendas_f[vendas_f["Tipo_Venda"] == "Normal"], "_qtd_venda", 0.0)
        if "Tipo_Venda" in vendas_f.columns and "_qtd_venda" in vendas_f.columns
        else 0.0
    )
    st.markdown(
        f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Vendas Facilitadas</div><div class="val">{fmt_qtd(qtd_facilitada)}</div></div>
            <div class="vel-kpi"><div class="lbl">Vendas Normais</div><div class="val">{fmt_qtd(qtd_normal)}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _mtd_qtd_mes_e_media_hist(
    dff: pd.DataFrame,
    col_d: str,
    emp_c: str,
    col_e: Optional[str],
    dia_atual: int,
) -> Tuple[float, float]:
    """Quantidade MTD parcial do mês corrente e média histórica (mesma janela dia 1→hoje)."""
    if dff is None or dff.empty or not col_d:
        return 0.0, 0.0
    sub = dff.copy()
    if col_e and col_e in sub.columns:
        sub = sub[sub[col_e].map(_limpar_emp) == emp_c]
    comp = _montar_df_comparativo_mtd_parcial(sub, col_d, dia_atual)
    if comp.empty:
        return 0.0, 0.0
    hoje = date.today()
    cur = comp[(comp["_ano_c"] == hoje.year) & (comp["_mes_c"] == hoje.month)]
    qtd_mes = float(cur["QTD"].iloc[0]) if not cur.empty else float(comp.iloc[-1]["QTD"])
    hist = comp[(comp["_ano_c"] != hoje.year) | (comp["_mes_c"] != hoje.month)]
    media = float(hist["QTD"].mean()) if not hist.empty else 0.0
    return qtd_mes, media


def filtros_v2_to_dashboard(fv2: "FiltrosPainelV2") -> "FiltrosDashboard":
    return FiltrosDashboard(
        data_ini=fv2.data_ini,
        data_fim=fv2.data_fim,
        mes_meta=fv2.mes_meta,
        ano_meta=fv2.ano_meta,
        tipo_meta_col=fv2.tipo_meta_col,
        emps_sel=fv2.emps_sel,
        coords_sel=fv2.coordenadores_sel,
        canal_sel=fv2.canal_meta or "Todos",
        imobs_sel=[],
    )


def _anexar_colunas_por_emp(
    base: pd.DataFrame,
    extra: pd.DataFrame,
    *,
    sobrescrever: bool = False,
) -> pd.DataFrame:
    """Left join de colunas extras por Empreendimento (normalizado)."""
    if base is None or base.empty:
        return base if base is not None else pd.DataFrame()
    if extra is None or extra.empty or "Empreendimento" not in extra.columns:
        return base
    out = base.copy()
    ext = extra.copy()
    ext["_k"] = ext["Empreendimento"].map(_limpar_emp)
    out["_k"] = out["Empreendimento"].map(_limpar_emp)
    cols = [
        c for c in ext.columns
        if c not in ("Empreendimento", "_k") and (sobrescrever or c not in out.columns)
    ]
    if not cols:
        return out.drop(columns=["_k"], errors="ignore")
    m = ext[["_k"] + cols].drop_duplicates(subset=["_k"], keep="last")
    merged = out.merge(m, on="_k", how="left")
    return merged.drop(columns=["_k"], errors="ignore")


def _share_estoque_por_emp(
    df_estoque: pd.DataFrame,
    empreendimentos: List[str],
    status_sel: Optional[List[str]] = None,
) -> pd.DataFrame:
    if df_estoque is None or df_estoque.empty or not empreendimentos:
        return pd.DataFrame()
    df = df_estoque.copy()
    col_st = "StatusUnidade__c" if "StatusUnidade__c" in df.columns else None
    if col_st and status_sel:
        df = df[df[col_st].astype(str).str.strip().isin(status_sel)]
    if "Empreendimento" not in df.columns:
        return pd.DataFrame()
    cont = df.groupby(df["Empreendimento"].map(_limpar_emp)).size()
    total = float(cont.sum())
    rows = []
    for emp in empreendimentos:
        emp_c = _limpar_emp(emp)
        u = int(cont.get(emp_c, 0))
        rows.append({
            "Empreendimento": emp_c,
            "Share_Estoque_Pct": round(u / total * 100.0, 1) if total > 0 else 0.0,
        })
    return pd.DataFrame(rows)


def _share_vendas_por_emp(
    df_vendas: pd.DataFrame,
    col_data: str,
    filtros_dc: "FiltrosDashboard",
    mapa_coord: Dict[str, str],
    empreendimentos: List[str],
) -> pd.DataFrame:
    if not col_data or df_vendas.empty or not empreendimentos:
        return pd.DataFrame()
    base = _aplicar_filtros_base(df_vendas, filtros_dc, mapa_coord, col_data, usar_periodo=True)
    if base.empty or "Empreendimento" not in base.columns:
        return pd.DataFrame()
    qcol = "_qtd_venda" if "_qtd_venda" in base.columns else None
    if not qcol:
        base = base.copy()
        base["_q"] = 1.0
        qcol = "_q"
    cont = base.groupby(base["Empreendimento"].map(_limpar_emp))[qcol].sum()
    total = float(cont.sum())
    rows = []
    for emp in empreendimentos:
        emp_c = _limpar_emp(emp)
        q = float(cont.get(emp_c, 0))
        rows.append({
            "Empreendimento": emp_c,
            "Share_Vendas_Pct": round(q / total * 100.0, 1) if total > 0 else 0.0,
        })
    return pd.DataFrame(rows)


def enriquecer_analitico_metricas_extras(
    tab: pd.DataFrame,
    filtros: "FiltrosPainelV2",
    mapa_coord: Dict[str, str],
    df_vendas: pd.DataFrame,
    df_estoque: pd.DataFrame,
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    df_pastas_aprov: pd.DataFrame,
    df_tabela_comp: pd.DataFrame,
    df_metas_coord: pd.DataFrame,
    df_cotacoes: Optional[pd.DataFrame],
    col_contrato: str,
    col_data: str,
    df_est_enr: pd.DataFrame,
    resumo_est: pd.DataFrame,
    status_estoque: Optional[List[str]] = None,
) -> pd.DataFrame:
    """Anexa tempos, hiper, shares, PS/VGV, VSO e radar à tabela analítica."""
    if tab is None or tab.empty:
        return tab if tab is not None else pd.DataFrame()

    emps = [_limpar_emp(e) for e in tab["Empreendimento"].tolist()]
    out = tab.copy()

    tab_tempos = montar_tabela_tempos_conversao(df_ag, df_pastas, df_vendas, emps)
    out = _anexar_colunas_por_emp(out, tab_tempos)

    pas_aprov = (
        _v().deduplicar_pastas_aprovadas_funil(df_pastas)
        if df_pastas is not None and not df_pastas.empty
        else pd.DataFrame()
    )
    tab_hiper = calcular_hipereficiencia_por_emp(
        df_vendas, pas_aprov, df_est_enr, df_tabela_comp, emps,
    )
    out = _anexar_colunas_por_emp(out, tab_hiper)

    if not resumo_est.empty and "Empreendimento" in resumo_est.columns:
        m2_rows = []
        for emp in emps:
            rs = resumo_est.loc[resumo_est["Empreendimento"] == emp]
            m2_rows.append({
                "Empreendimento": emp,
                "m2_Total": float(rs["m2_Total"].iloc[0]) if not rs.empty and "m2_Total" in rs.columns else 0.0,
            })
        out = _anexar_colunas_por_emp(out, pd.DataFrame(m2_rows))

    filtros_dc = filtros_v2_to_dashboard(filtros)
    filtros_dc.emps_sel = emps

    df_v_ps = enriquecer_vendas_vcx(df_vendas, df_cotacoes)
    col_d = col_data or col_contrato
    if col_d:
        tab_ps = montar_tabela_ps_sinais_vgv(df_v_ps, col_d, filtros_dc, mapa_coord)
        out = _anexar_colunas_por_emp(out, tab_ps)

        tab_share_v = _share_vendas_por_emp(df_v_ps, col_d, filtros_dc, mapa_coord, emps)
        out = _anexar_colunas_por_emp(out, tab_share_v)

        df_v_f = _aplicar_filtros_base(df_v_ps, filtros_dc, mapa_coord, col_d, usar_periodo=True)
        df_vso = calcular_vso_por_emp(
            df_v_ps, df_estoque, filtros_dc, mapa_coord, col_d, ref_fim=filtros.data_fim,
        )
        if not df_vso.empty:
            vso_rows = []
            for _, r in df_vso.iterrows():
                emp = _limpar_emp(r["Empreendimento"])
                if emp not in emps:
                    continue
                meta_v = soma_meta_coord(
                    df_metas_coord, filtros.mes_meta, filtros.ano_meta,
                    "vendas", filtros.tipo_meta_col, empreendimentos=[emp],
                )
                sub = df_v_f[df_v_f["Empreendimento"].map(_limpar_emp) == emp] if not df_v_f.empty else pd.DataFrame()
                real_qtd, _ = _qtd_vgv(sub)
                row: Dict[str, Any] = {"Empreendimento": emp}
                for dias in JANELAS_VSO:
                    row[f"VSO_{dias}d"] = round(float(r.get(f"VSO_{dias}d", 0)), 2)
                if "Meta_Vendas" not in out.columns:
                    row["Meta_Vendas"] = meta_v
                if "Pct_Meta" not in out.columns:
                    row["Pct_Meta"] = round(real_qtd / meta_v * 100.0, 1) if meta_v > 0 else 0.0
                vso_rows.append(row)
            out = _anexar_colunas_por_emp(out, pd.DataFrame(vso_rows))

    st_status = status_estoque or list(ESTOQUE_STATUS_TODOS)
    tab_share_e = _share_estoque_por_emp(df_estoque, emps, st_status)
    out = _anexar_colunas_por_emp(out, tab_share_e)

    mapa_coord = mapa_emp_coordenador(df_metas_coord, filtros.mes_meta, filtros.ano_meta)
    df_pas_aprov_radar = (
        pas_aprov if not pas_aprov.empty
        else (df_pastas_aprov if df_pastas_aprov is not None and not df_pastas_aprov.empty else pd.DataFrame())
    )
    if col_d:
        tab_radar = montar_tabela_radar_emp(
            df_vendas,
            df_estoque,
            df_metas_coord,
            df_pas_aprov_radar,
            df_tabela_comp if df_tabela_comp is not None else pd.DataFrame(),
            filtros_dc,
            mapa_coord,
            col_d,
        )
        out = _anexar_colunas_por_emp(out, tab_radar)

    return out


def enriquecer_analitico_com_mtd(
    df: pd.DataFrame,
    df_vendas: pd.DataFrame,
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    col_contrato: str,
) -> pd.DataFrame:
    """Acrescenta colunas MTD mês vs média histórica por indicador na tabela analítica."""
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    v = _v()
    dia_atual = datetime.now().day
    out = df.copy()

    df_ag_u = df_ag if df_ag is not None else pd.DataFrame()
    df_pas_u = df_pastas if df_pastas is not None else pd.DataFrame()
    df_pas_ap = v.deduplicar_pastas_aprovadas_funil(df_pas_u)

    try:
        pacote = carregar_funil_historico_painel_sf()
        df_ag_u = deduplicar_agendamentos_funil(_coalesce_df(pacote.get("agendamentos")))
        df_pas_raw = _coalesce_df(pacote.get("pastas"))
        df_pas_u = deduplicar_pastas_funil(df_pas_raw)
        df_pas_ap = deduplicar_pastas_aprovadas_funil(df_pas_raw)
    except Exception:
        pass

    col_ag = v.achar_coluna(df_ag_u, v.ALIASES_DATA_CRIACAO)
    col_vis = v.achar_coluna(df_ag_u, getattr(v, "ALIASES_DATA_VISITA", ["Data da visita"]))
    col_envio = v.achar_coluna_primeiro_envio_analise(df_pas_u)
    col_safi = v.achar_coluna_aprovacao_safi(df_pas_ap)
    col_emp_ag = v.achar_coluna(df_ag_u, v.ALIASES_EMPREENDIMENTO)
    col_emp_pas = v.achar_coluna(df_pas_u, v.ALIASES_EMPREENDIMENTO)
    col_emp_pas_ap = v.achar_coluna(df_pas_ap, v.ALIASES_EMPREENDIMENTO)

    fontes: List[Tuple[str, pd.DataFrame, Optional[str], Optional[str]]] = [
        ("Vendas", df_vendas, col_contrato or None, "Empreendimento"),
        ("Agendamentos", df_ag_u, col_ag, col_emp_ag),
        ("Visitas", df_ag_u, col_vis, col_emp_ag),
        ("Pastas", df_pas_u, col_envio, col_emp_pas),
        ("Pastas_Aprov", df_pas_ap, col_safi, col_emp_pas_ap),
    ]

    for sufixo, dff, col_d, col_e in fontes:
        col_mtd = f"MTD_{sufixo}"
        col_med = f"Media_Hist_{sufixo}"
        vals_mtd: List[float] = []
        vals_med: List[float] = []
        for emp in out["Empreendimento"]:
            emp_c = v._limpar_emp(emp)
            qtd_m, qtd_h = _mtd_qtd_mes_e_media_hist(dff, col_d or "", emp_c, col_e, dia_atual)
            vals_mtd.append(qtd_m)
            vals_med.append(qtd_h)
        out[col_mtd] = vals_mtd
        out[col_med] = vals_med

    if col_contrato:
        hoje = date.today()
        ini_mes = date(hoje.year, hoje.month, 1)
        precos_mtd: List[float] = []
        gaps_preco: List[float] = []
        for idx, emp in enumerate(out["Empreendimento"]):
            emp_c = v._limpar_emp(emp)
            ven_emp = _filtrar_df_periodo(df_vendas, col_contrato, ini_mes, hoje)
            if "Empreendimento" in ven_emp.columns:
                ven_emp = ven_emp[ven_emp["Empreendimento"].map(_limpar_emp) == emp_c]
            col_val = v.achar_coluna(ven_emp, ["Valor Real de Venda", "Valor Real", "Valor"])
            preco_v = (
                float(ven_emp[col_val].map(_parse_num_br).mean())
                if col_val and not ven_emp.empty else 0.0
            )
            preco_t = float(out.at[idx, "Preco_Medio_Tabela"]) if "Preco_Medio_Tabela" in out.columns else 0.0
            precos_mtd.append(preco_v)
            gaps_preco.append(
                ((preco_v - preco_t) / preco_t * 100.0) if preco_t > 0 else 0.0
            )
        out["Preco_Medio_Venda_MTD"] = precos_mtd
        out["Gap_Preco_Pct"] = gaps_preco

    return out


def render_painel_metas_v2(
    df_vendas: pd.DataFrame,
    df_vendas_painel: pd.DataFrame,
    col_contrato_gerado: Optional[str],
    cred_fp: str,
    df_estoque: Optional[pd.DataFrame] = None,
    df_ag: Optional[pd.DataFrame] = None,
    df_pastas: Optional[pd.DataFrame] = None,
    df_cotacoes: Optional[pd.DataFrame] = None,
    df_pastas_aprov: Optional[pd.DataFrame] = None,
    df_tabela_comp: Optional[pd.DataFrame] = None,
    proj: Optional[Dict[str, Any]] = None,
    filtros_glob: Optional["FiltrosGlobais"] = None,
    df_metas_fallback: Optional[pd.DataFrame] = None,
    col_canal: Optional[str] = None,
) -> FiltrosPainelV2:
    """Renderiza seção v2: estoque, velocímetros, tabela analítica."""
    v = _v()
    ano_fb = filtros_glob.ano_meta if filtros_glob else date.today().year
    mes_fb = filtros_glob.mes_meta if filtros_glob else date.today().month
    df_metas_coord, aviso_coord = carregar_metas_coordenadores_com_fallback(
        cred_fp, df_metas_fallback, ano_fb, mes_fb,
    )
    try:
        df_canal = carregar_metas_canal(cred_fp)
    except Exception as exc:
        st.warning(f"Metas canal (IVAN) indisponíveis: {exc}")
        df_canal = pd.DataFrame()
    if df_metas_coord.empty:
        st.error(f"Erro ao carregar planilhas de meta: {aviso_coord or 'sem dados'}")
        return FiltrosPainelV2(
            date.today().replace(day=1), date.today(),
            date.today().month, date.today().year,
            "vendas", "Desafio", "RIO", [], [],
        )
    if aviso_coord:
        st.warning(
            f"Metas coordenadores: usando planilha legado de metas ({aviso_coord})."
        )

    filtros = render_filtros_painel_v2(df_metas_coord, cred_fp, filtros_externos=filtros_glob)
    if filtros_glob is not None:
        st.caption(f"Filtros globais ativos · {filtros.data_ini:%d/%m/%Y} → {filtros.data_fim:%d/%m/%Y}")
    mapa_coord = mapa_emp_coordenador(df_metas_coord, filtros.mes_meta, filtros.ano_meta)

    kpi_est, enr = agregar_estoque(df_estoque if df_estoque is not None else pd.DataFrame())
    resumo_est = resumo_estoque_por_emp(enr)
    estoque_map_v2 = resumo_estoque_empreendimentos(
        df_estoque if df_estoque is not None else pd.DataFrame(),
    )
    try:
        total_unidades_v2 = carregar_total_unidades_por_emp_sf()
    except Exception:
        total_unidades_v2 = {}

    if filtros.tipo_indicador == "vendas":
        meta_vgv, meta_qtd_canal = meta_canal_vgv_vendas(
            df_canal, filtros.mes_meta, filtros.ano_meta, filtros.canal_meta,
        )
        meta_qtd = soma_meta_coord(
            df_metas_coord, filtros.mes_meta, filtros.ano_meta,
            "vendas", filtros.tipo_meta_col,
        )
        if meta_vgv <= 0:
            meta_vgv = soma_meta_vgv_coord(
                df_metas_coord, filtros.mes_meta, filtros.ano_meta, filtros.tipo_meta_col,
            )
        if meta_qtd <= 0 and 0 < meta_qtd_canal <= 5_000:
            meta_qtd = meta_qtd_canal
        real_qtd, real_vgv = realizado_vendas_periodo(
            df_vendas, col_contrato_gerado or "", filtros.data_ini, filtros.data_fim,
            filtros.emps_sel or None,
        )
    else:
        meta_vgv = 0.0
        meta_qtd = soma_meta_coord(
            df_metas_coord, filtros.mes_meta, filtros.ano_meta,
            filtros.tipo_indicador, filtros.tipo_meta_col,
            filtros.emps_sel or None,
        )
        real_qtd = 0.0
        real_vgv = 0.0
        emps_calc = filtros.emps_sel or list(mapa_coord.keys())
        df_ag_u = df_ag if df_ag is not None else pd.DataFrame()
        df_pas_u = df_pastas if df_pastas is not None else pd.DataFrame()
        for emp in emps_calc:
            f = realizado_funil_periodo(
                df_ag_u, df_pas_u, df_vendas, v._limpar_emp(emp),
                filtros.data_ini, filtros.data_fim,
            )
            real_qtd += float(f.get(filtros.tipo_indicador, 0))

    if col_canal and col_contrato_gerado:
        vendas_perfil = filtrar_vendas_painel_v2(
            df_vendas, filtros, col_contrato_gerado, col_canal,
        )
        render_perfil_vendas_mtd(vendas_perfil)
        st.markdown(
            "<hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>",
            unsafe_allow_html=True,
        )

    render_velocimetro_principal(
        filtros, df_canal, df_metas_coord,
        real_qtd, real_vgv, meta_qtd, meta_vgv,
        escala_meta_regressao(proj, meta_qtd),
    )
    render_velocimetros_coordenador(
        filtros, df_metas_coord, df_vendas,
        df_ag if df_ag is not None else pd.DataFrame(),
        df_pastas if df_pastas is not None else pd.DataFrame(),
        col_contrato_gerado or "", mapa_coord,
    )

    emps_tab = _empreendimentos_para_analitico(
        filtros, mapa_coord, df_vendas, df_metas_fallback,
    )
    tab = montar_tabela_analitica(
        emps_tab[:80],
        enr, resumo_est,
        df_vendas,
        df_ag if df_ag is not None else pd.DataFrame(),
        df_pastas if df_pastas is not None else pd.DataFrame(),
        df_metas_coord, filtros,
        col_contrato_gerado or "", proj,
        df_cotacoes=df_cotacoes,
        df_pastas_aprov=df_pastas_aprov,
        df_tabela_comp=df_tabela_comp,
        estoque_map=estoque_map_v2,
        total_unidades_por_emp=total_unidades_v2,
    )
    tab = enriquecer_analitico_com_mtd(
        tab,
        df_vendas,
        df_ag if df_ag is not None else pd.DataFrame(),
        df_pastas if df_pastas is not None else pd.DataFrame(),
        col_contrato_gerado or "",
    )
    status_est = (
        filtros_glob.status_estoque_sel
        if filtros_glob and filtros_glob.status_estoque_sel
        else list(ESTOQUE_STATUS_TODOS)
    )
    tab = enriquecer_analitico_metricas_extras(
        tab,
        filtros,
        mapa_coord,
        df_vendas,
        df_estoque if df_estoque is not None else pd.DataFrame(),
        df_ag if df_ag is not None else pd.DataFrame(),
        df_pastas if df_pastas is not None else pd.DataFrame(),
        df_pastas_aprov if df_pastas_aprov is not None else pd.DataFrame(),
        df_tabela_comp if df_tabela_comp is not None else pd.DataFrame(),
        df_metas_coord,
        df_cotacoes,
        col_contrato_gerado or "",
        col_contrato_gerado or "",
        enr,
        resumo_est,
        status_estoque=status_est,
    )
    render_tabela_analitica(tab)
    return filtros


# =============================================================================
# PODER DE COMPRA (inline — ex-velocimetro_poder_compra.py)
# =============================================================================

# Poder de compra — pastas aprovadas (Avaliacao_credito__c).
# pro_soluto_max = min(ProSoluto__c, renda×comprometimento×parcelas, preço×limite %)

import re
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Dict, List, Optional, Set, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# Defaults (diresimulator / condição comercial RJ)
DEFAULT_COMPROMETIMENTO_RENDA = 0.30
DEFAULT_MAX_PARCELAS = 84
DEFAULT_LIMITE_PS_PRECO = 0.25

ALIASES_RENDA = ["RendaApurada__c", "Renda Apurada", "Valor_da_Renda__c", "Valor da Renda", "Renda__c", "Renda"]
ALIASES_FGTS = ["FGTS_apurado__c", "FGTS apurado", "Valor_FGTS__c", "Valor FGTS"]
ALIASES_SUBSIDIO = ["Valor_de_Subsidio__c", "Valor do Subsidio", "Valor de Subsidio", "Subsídio"]
ALIASES_FINANCIAMENTO = ["Valor_Financiamento__c", "Valor do Financiamento"]
ALIASES_TIPO_AVAL = ["Tipo__c", "Tipo"]
ALIASES_TIPOLOGIA = ["Tipologia__c", "Tipologia"]
ALIASES_OPP_AVAL = ["Oportunidade__c", "Oportunidade", "ID da Oportunidade"]


@dataclass
class PastaPoderCompra:
    chave: str
    empreendimento: str
    tipo: str
    renda: float
    fgts: float
    subsidio: float
    financiamento: float
    pro_soluto_max: float
    pro_soluto_efetivo: float
    poder_compra: float
    oportunidade_id: str
    comprou: bool
    pc_suficiente: bool = False
    preco_referencia: float = 0.0


@dataclass
class ResumoIneficienciaEmp:
    empreendimento: str
    pastas_aprovadas: int = 0
    pastas_pc_suficiente: int = 0
    vendas: int = 0
    ineficiencia_qtd: int = 0
    ineficiencia_pct: float = 0.0




def _num(val: Any) -> float:
    v = _v()
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "nat"):
        return 0.0
    try:
        return float(v.parse_valor_br(val))
    except Exception:
        try:
            return float(s.replace(".", "").replace(",", "."))
        except Exception:
            return 0.0


def _col_val(row: pd.Series, aliases: List[str]) -> float:
    for a in aliases:
        if a in row.index:
            x = _num(row.get(a))
            if x > 0:
                return x
    v = _v()
    col = v.achar_coluna(pd.DataFrame([row]), aliases)
    return _num(row.get(col)) if col else 0.0


def renda_da_pasta(row: pd.Series) -> float:
    return _col_val(row, ALIASES_RENDA)


def calcular_sinal_sobre_renda_por_emp(
    df_cotacoes: Optional[pd.DataFrame],
    df_pastas: Optional[pd.DataFrame],
    data_ini: Optional[date] = None,
    data_fim: Optional[date] = None,
) -> Dict[str, float]:
    """
    Σ sinal / Σ renda por empreendimento.
    Pastas no período; sinal = maior Total Sinal Com por oportunidade (cotação).
    """
    if df_pastas is None or df_pastas.empty:
        return {}
    v = _v()
    pas = df_pastas.copy()
    col_e = v.achar_coluna(pas, v.ALIASES_EMPREENDIMENTO) or "Empreendimento"
    col_d = (
        v.achar_coluna_primeiro_envio_analise(pas)
        or v.achar_coluna(pas, v.ALIASES_DATA_CRIACAO)
    )
    if data_ini and data_fim and col_d:
        pas = _filtrar_df_periodo(pas, col_d, data_ini, data_fim)
    if pas.empty or col_e not in pas.columns:
        return {}

    col_opp = v.achar_coluna(pas, ALIASES_OPP_AVAL) or "Oportunidade"
    pas["_renda"] = pas.apply(renda_da_pasta, axis=1)

    sinal_por_opp: Dict[str, float] = {}
    if df_cotacoes is not None and not df_cotacoes.empty:
        col_opp_c = "ID da Oportunidade"
        col_sinal = next(
            (c for c in ("Total Sinal Com", "TotalSinalCom__c") if c in df_cotacoes.columns),
            None,
        )
        if col_opp_c in df_cotacoes.columns and col_sinal:
            cot = df_cotacoes.copy()
            cot["_sinal"] = cot[col_sinal].map(_parse_num_br)
            sinal_por_opp = cot.groupby(col_opp_c)["_sinal"].max().to_dict()

    if col_opp in pas.columns and sinal_por_opp:
        pas["_sinal"] = pas[col_opp].astype(str).map(
            lambda x: sinal_por_opp.get(x, 0.0)
        ).fillna(0.0)
    else:
        pas["_sinal"] = 0.0

    out: Dict[str, float] = {}
    pas["_emp_c"] = pas[col_e].map(_limpar_emp)
    for emp, g in pas.groupby("_emp_c"):
        sum_renda = float(g["_renda"].sum())
        sum_sinal = float(g["_sinal"].sum())
        if sum_renda > 0:
            out[str(emp)] = sum_sinal / sum_renda
    return out


def montar_tabela_sinal_renda_estoque_emp(
    empreendimentos: List[str],
    estoque_map: Dict[str, Dict[str, int]],
    total_unidades_por_emp: Dict[str, int],
    sinal_renda_por_emp: Dict[str, float],
) -> pd.DataFrame:
    """Tabela consolidada: sinal/renda e ratios de estoque por empreendimento."""
    lib_map = metricas_liberacao_estoque_por_emp(estoque_map, total_unidades_por_emp)
    rows = []
    for emp in empreendimentos:
        emp_c = _limpar_emp(emp)
        lib = lib_map.get(emp_c, {})
        ratio_sr = sinal_renda_por_emp.get(emp_c)
        rows.append({
            "Empreendimento": emp_c,
            "Unidades_Disponiveis": int(lib.get("disponivel", 0)),
            "Unidades_Liberadas": int(lib.get("liberadas", 0)),
            "Unidades_Total": int(lib.get("total", 0)),
            "Pct_Disp_Liberadas": round(float(lib.get("pct_disp_liberadas", 0.0)), 1),
            "Pct_Liberadas_Total": round(float(lib.get("pct_liberadas_total", 0.0)), 1),
            "Sinal_Sobre_Renda_Pct": round(ratio_sr * 100.0, 1) if ratio_sr is not None else None,
        })
    return pd.DataFrame(rows)


def parse_identificador_unidade(ident: Any) -> Dict[str, str]:
    """Extrai bloco, andar e identificador normalizado."""
    s = str(ident or "").strip().upper()
    bloco = ""
    m = re.search(r"BL\s*0*(\d+)", s)
    if m:
        bloco = m.group(1)
    elif "-" in s:
        bloco = s.split("-")[0].replace("BL", "").strip()
    andar = s[-4:-2] if len(s) >= 4 else ""
    return {"identificador": s, "bloco": bloco, "andar": andar}


def calcular_pro_soluto_maximo(
    renda: float,
    preco_unidade: float,
    tabela_row: Optional[Dict[str, Any]] = None,
    comprometimento: float = DEFAULT_COMPROMETIMENTO_RENDA,
    max_parcelas: int = DEFAULT_MAX_PARCELAS,
    limite_pct_preco: float = DEFAULT_LIMITE_PS_PRECO,
) -> float:
    """
    Pro soluto máximo (estilo diresimulator):
    - Tabela SF: ProSoluto__c e/ou renda × ComprometimentoDeRenda × numParcelas
    - Teto percentual sobre o preço da unidade (Limite Pro Soluto ~25%)
    """
    renda = max(float(renda or 0), 0.0)
    preco = max(float(preco_unidade or 0), 0.0)
    ps_tabela = 0.0
    ps_renda = 0.0
    if tabela_row:
        ps_tabela = _num(tabela_row.get("ProSoluto__c") or tabela_row.get("Pro Soluto") or tabela_row.get("ProSoluto"))
        comp = _num(tabela_row.get("ComprometimentoDeRenda__c") or tabela_row.get("Comprometimento de Renda Total"))
        if comp > 1:
            comp = comp / 100.0
        n_par = int(_num(tabela_row.get("numParcelas__c") or tabela_row.get("Nº Parcelas")) or max_parcelas)
        if renda > 0 and comp > 0:
            ps_renda = renda * comp * n_par
        elif renda > 0 and ps_tabela <= 0:
            ps_renda = renda * comprometimento * max_parcelas
    elif renda > 0:
        ps_renda = renda * comprometimento * max_parcelas

    ps_preco = preco * limite_pct_preco if preco > 0 else 0.0
    candidatos = [x for x in (ps_tabela, ps_renda, ps_preco) if x > 0]
    if not candidatos:
        return 0.0
    if ps_preco > 0:
        return min(min(candidatos), ps_preco) if len(candidatos) > 1 else min(candidatos[0], ps_preco)
    return min(candidatos)


def calcular_poder_compra(
    fgts: float,
    subsidio: float,
    financiamento: float,
    pro_soluto_max: float,
    pro_soluto_aprovado: Optional[float] = None,
) -> Tuple[float, float]:
    """Retorna (poder_compra, pro_soluto_efetivo)."""
    ps_max = max(float(pro_soluto_max or 0), 0.0)
    ps_ap = max(float(pro_soluto_aprovado or 0), 0.0) if pro_soluto_aprovado else ps_max
    ps_ef = min(ps_ap, ps_max) if ps_max > 0 else ps_ap
    pc = max(fgts, 0) + max(subsidio, 0) + max(financiamento, 0) + ps_ef
    return pc, ps_ef


def _normalizar_tipo(val: Any) -> str:
    return str(val or "").strip().upper()


def filtrar_estoque_tipo(df_est: pd.DataFrame, tipo_avaliacao: str) -> pd.DataFrame:
    """Filtra unidades vendáveis compatíveis com o tipo da avaliação."""
    v = _v()
    if df_est is None or df_est.empty:
        return pd.DataFrame()
    df = df_est.copy()
    col_st = v.achar_coluna(df, v.ALIASES_STATUS_UNIDADE if hasattr(v, "ALIASES_STATUS_UNIDADE") else ["StatusUnidade__c", "Status"])
    if col_st:
        vendavel = df[col_st].astype(str).str.strip().isin(v.ESTOQUE_STATUS_VENDAVEL)
        df = df.loc[vendavel]
    tipo = _normalizar_tipo(tipo_avaliacao)
    if not tipo:
        return df
    col_tip = v.achar_coluna(df, ALIASES_TIPOLOGIA)
    if col_tip:
        df = df[df[col_tip].map(_normalizar_tipo) == tipo]
    return df


def preco_minimo_empreendimento(
    df_est_enr: pd.DataFrame,
    emp: str,
    tipo: str = "",
) -> float:
    """Menor preço tabela (VFK−BA−folga) elegível para o empreendimento/tipo."""
    v = _v()
    emp_c = v._limpar_emp(emp)
    if df_est_enr is None or df_est_enr.empty:
        return 0.0
    sub = df_est_enr[df_est_enr["Empreendimento"].map(v._limpar_emp) == emp_c]
    if tipo:
        sub = sub[sub["Tipologia"].map(_normalizar_tipo) == _normalizar_tipo(tipo)] if "Tipologia" in sub.columns else sub
    if sub.empty or "PrecoTabela" not in sub.columns:
        return 0.0
    vals = sub["PrecoTabela"].astype(float)
    vals = vals[vals > 0]
    return float(vals.min()) if not vals.empty else 0.0


def mapa_tabela_por_oportunidade(df_tab: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    v = _v()
    out: Dict[str, Dict[str, Any]] = {}
    if df_tab is None or df_tab.empty:
        return out
    col_opp = v.achar_coluna(df_tab, ALIASES_OPP_AVAL) or "Oportunidade__c"
    for _, row in df_tab.iterrows():
        oid = str(row.get(col_opp) or "").strip()
        if oid and oid not in out:
            out[oid] = row.to_dict()
    return out


def ids_vendas_empreendimento(df_vendas: pd.DataFrame, emp: str) -> Set[str]:
    v = _v()
    emp_c = v._limpar_emp(emp)
    if df_vendas is None or df_vendas.empty:
        return set()
    col_emp = v.achar_coluna(df_vendas, v.ALIASES_EMPREENDIMENTO)
    col_id = v.achar_coluna(df_vendas, v.ALIASES_ID_OPORTUNIDADE)
    if not col_emp or not col_id:
        return set()
    sub = df_vendas[df_vendas[col_emp].map(v._limpar_emp) == emp_c]
    return {str(x).strip() for x in sub[col_id].dropna().unique() if str(x).strip()}


def analisar_pasta(
    row: pd.Series,
    emp: str,
    df_est_enr: pd.DataFrame,
    mapa_tab: Dict[str, Dict[str, Any]],
    vendas_ids: Set[str],
    preco_override: Optional[float] = None,
) -> PastaPoderCompra:
    v = _v()
    chave_col = v.achar_coluna(pd.DataFrame([row]), v.ALIASES_NOME_AVALIACAO_CREDITO) or "Name"
    chave = str(row.get(chave_col) or row.get("Name") or row.name)
    tipo = ""
    col_tipo = v.achar_coluna(pd.DataFrame([row]), ALIASES_TIPO_AVAL)
    if col_tipo:
        tipo = str(row.get(col_tipo) or "").strip()

    renda = renda_da_pasta(row)
    fgts = _col_val(row, ALIASES_FGTS)
    subsidio = _col_val(row, ALIASES_SUBSIDIO)
    financiamento = _col_val(row, ALIASES_FINANCIAMENTO)

    col_opp = v.achar_coluna(pd.DataFrame([row]), ALIASES_OPP_AVAL)
    opp_id = str(row.get(col_opp) or "").strip() if col_opp else ""

    preco_ref = preco_override if preco_override is not None else preco_minimo_empreendimento(df_est_enr, emp, tipo)
    tab_row = mapa_tab.get(opp_id)
    ps_max = calcular_pro_soluto_maximo(renda, preco_ref, tab_row)
    pc, ps_ef = calcular_poder_compra(fgts, subsidio, financiamento, ps_max)
    comprou = opp_id in vendas_ids if opp_id else False
    pc_ok = pc >= preco_ref if preco_ref > 0 else False

    return PastaPoderCompra(
        chave=chave,
        empreendimento=v._limpar_emp(emp),
        tipo=tipo,
        renda=renda,
        fgts=fgts,
        subsidio=subsidio,
        financiamento=financiamento,
        pro_soluto_max=ps_max,
        pro_soluto_efetivo=ps_ef,
        poder_compra=pc,
        oportunidade_id=opp_id,
        comprou=comprou,
        pc_suficiente=pc_ok,
        preco_referencia=preco_ref,
    )


def calcular_resumo_ineficiencia_emp(
    df_pastas_aprov: pd.DataFrame,
    df_est_enr: pd.DataFrame,
    df_vendas: pd.DataFrame,
    df_tabela: pd.DataFrame,
    emp: str,
    preco_override: Optional[float] = None,
) -> ResumoIneficienciaEmp:
    v = _v()
    emp_c = v._limpar_emp(emp)
    col_emp = v.achar_coluna(df_pastas_aprov, v.ALIASES_EMPREENDIMENTO) if not df_pastas_aprov.empty else None
    if not col_emp or df_pastas_aprov.empty:
        return ResumoIneficienciaEmp(empreendimento=emp_c)

    pastas = df_pastas_aprov[df_pastas_aprov[col_emp].map(v._limpar_emp) == emp_c]
    mapa_tab = mapa_tabela_por_oportunidade(df_tabela)
    vendas_ids = ids_vendas_empreendimento(df_vendas, emp_c)

    analises = [
        analisar_pasta(r, emp_c, df_est_enr, mapa_tab, vendas_ids, preco_override)
        for _, r in pastas.iterrows()
    ]
    n_aprov = len(analises)
    n_pc = sum(1 for a in analises if a.pc_suficiente)
    n_vendas = len(vendas_ids & {a.oportunidade_id for a in analises if a.oportunidade_id})
    inef_qtd = sum(1 for a in analises if a.pc_suficiente and not a.comprou)
    inef_pct = (inef_qtd / n_aprov * 100.0) if n_aprov > 0 else 0.0

    return ResumoIneficienciaEmp(
        empreendimento=emp_c,
        pastas_aprovadas=n_aprov,
        pastas_pc_suficiente=n_pc,
        vendas=n_vendas,
        ineficiencia_qtd=inef_qtd,
        ineficiencia_pct=inef_pct,
    )


def calcular_resumos_todos_empreendimentos(
    df_pastas_aprov: pd.DataFrame,
    df_est_enr: pd.DataFrame,
    df_vendas: pd.DataFrame,
    df_tabela: pd.DataFrame,
    empreendimentos: Optional[List[str]] = None,
) -> pd.DataFrame:
    v = _v()
    if df_pastas_aprov is None or df_pastas_aprov.empty:
        return pd.DataFrame()
    col_emp = v.achar_coluna(df_pastas_aprov, v.ALIASES_EMPREENDIMENTO)
    if not col_emp:
        return pd.DataFrame()
    emps = empreendimentos or sorted(df_pastas_aprov[col_emp].map(v._limpar_emp).dropna().unique())
    rows = []
    for emp in emps:
        r = calcular_resumo_ineficiencia_emp(df_pastas_aprov, df_est_enr, df_vendas, df_tabela, emp)
        rows.append({
            "Empreendimento": r.empreendimento,
            "Pastas_Aprovadas": r.pastas_aprovadas,
            "Pastas_PC_Suficiente": r.pastas_pc_suficiente,
            "Vendas": r.vendas,
            "Ineficiencia_Qtd": r.ineficiencia_qtd,
            "Ineficiencia_Pct": r.ineficiencia_pct,
        })
    return pd.DataFrame(rows)


def estatisticas_preco_estoque(
    df_est_enr: pd.DataFrame,
    emp: str,
    filtro_dim: str = "",
    filtro_val: str = "",
) -> Dict[str, float]:
    """Min, médio, mediano, máximo de PrecoTabela para recorte do estoque."""
    v = _v()
    emp_c = v._limpar_emp(emp)
    if df_est_enr is None or df_est_enr.empty:
        return {"min": 0, "medio": 0, "mediano": 0, "max": 0, "n": 0}
    sub = df_est_enr[df_est_enr["Empreendimento"].map(v._limpar_emp) == emp_c].copy()
    if filtro_dim and filtro_val:
        fv = str(filtro_val).strip().upper()
        if filtro_dim == "identificador":
            sub = sub[sub["Identificador"].astype(str).str.upper().str.strip() == fv]
        elif filtro_dim == "bloco":
            sub["_bloco"] = sub["Identificador"].map(lambda x: parse_identificador_unidade(x)["bloco"])
            sub = sub[sub["_bloco"] == fv.replace("BL", "").strip()]
        elif filtro_dim == "andar":
            sub["_andar"] = sub["Identificador"].map(lambda x: parse_identificador_unidade(x)["andar"])
            sub = sub[sub["_andar"] == fv]
        elif filtro_dim == "tipo":
            for c in ALIASES_TIPOLOGIA + ["Tipologia", "Tipo"]:
                if c in sub.columns:
                    sub = sub[sub[c].map(_normalizar_tipo) == fv]
                    break
    if sub.empty:
        return {"min": 0, "medio": 0, "mediano": 0, "max": 0, "n": 0}
    if "PrecoTabela" in sub.columns:
        vals = sub["PrecoTabela"].astype(float)
    else:
        col_preco = v.achar_coluna(sub, ALIASES_ESTOQUE_VFK) or "Valor Final com Kit"
        if col_preco not in sub.columns:
            return {"min": 0, "medio": 0, "mediano": 0, "max": 0, "n": 0}
        vals = sub[col_preco].map(_parse_num_br).astype(float)
    vals = vals[vals > 0]
    if vals.empty:
        return {"min": 0, "medio": 0, "mediano": 0, "max": 0, "n": 0}
    return {
        "min": float(vals.min()),
        "medio": float(vals.mean()),
        "mediano": float(vals.median()),
        "max": float(vals.max()),
        "n": int(len(vals)),
    }


def simular_novo_preco(
    df_pastas_aprov: pd.DataFrame,
    df_est_enr: pd.DataFrame,
    df_vendas: pd.DataFrame,
    df_tabela: pd.DataFrame,
    emp: str,
    novo_preco: float,
    preco_atual: float,
) -> Dict[str, Any]:
    """
    Clientes sem PC no preço atual que passariam a ter PC no novo preço (e ainda não compraram).
    """
    v = _v()
    emp_c = v._limpar_emp(emp)
    col_emp = v.achar_coluna(df_pastas_aprov, v.ALIASES_EMPREENDIMENTO)
    if not col_emp:
        return {"ganhos_pc": 0, "vendas_potenciais": 0, "detalhe": []}
    pastas = df_pastas_aprov[df_pastas_aprov[col_emp].map(v._limpar_emp) == emp_c]
    mapa_tab = mapa_tabela_por_oportunidade(df_tabela)
    vendas_ids = ids_vendas_empreendimento(df_vendas, emp_c)

    ganhos = 0
    detalhe = []
    for _, row in pastas.iterrows():
        a_atual = analisar_pasta(row, emp_c, df_est_enr, mapa_tab, vendas_ids, preco_atual)
        a_novo = analisar_pasta(row, emp_c, df_est_enr, mapa_tab, vendas_ids, novo_preco)
        if a_atual.comprou:
            continue
        if not a_atual.pc_suficiente and a_novo.pc_suficiente:
            ganhos += 1
            detalhe.append({
                "Pasta": a_atual.chave,
                "Poder Compra": a_novo.poder_compra,
                "Preço novo": novo_preco,
                "Gap": a_novo.poder_compra - novo_preco,
            })
    return {
        "ganhos_pc": ganhos,
        "vendas_potenciais": ganhos,
        "detalhe": detalhe,
    }


def render_aba_poder_compra(
    df_pastas_aprov: pd.DataFrame,
    df_est_enr: pd.DataFrame,
    df_vendas: pd.DataFrame,
    df_tabela: pd.DataFrame,
) -> None:
    v = _v()
    st.subheader("Poder de Compra & Simulação de Preço")
    st.caption(
        "Pastas aprovadas · poder de compra = FGTS + subsídio + financiamento + pro soluto "
        f"(máx. estilo diresimulator: renda×{DEFAULT_COMPROMETIMENTO_RENDA:.0%}×{DEFAULT_MAX_PARCELAS} parcelas "
        f"ou {DEFAULT_LIMITE_PS_PRECO:.0%} do preço, com tabela SF quando disponível)"
    )
    if df_pastas_aprov is None or df_pastas_aprov.empty:
        st.info("Sem pastas aprovadas carregadas.")
        return
    col_emp = v.achar_coluna(df_pastas_aprov, v.ALIASES_EMPREENDIMENTO)
    if not col_emp:
        st.warning("Coluna empreendimento não encontrada nas pastas.")
        return

    emps = sorted(df_pastas_aprov[col_emp].map(v._limpar_emp).dropna().unique())
    emp_sel = st.selectbox("Empreendimento", emps, key="pc_emp_sel")

    resumo = calcular_resumo_ineficiencia_emp(
        df_pastas_aprov, df_est_enr, df_vendas, df_tabela, emp_sel,
    )
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Pastas aprovadas", resumo.pastas_aprovadas)
    c2.metric("Com PC suficiente", resumo.pastas_pc_suficiente)
    c3.metric("Vendas", resumo.vendas)
    c4.metric("Ineficiência (qtd)", resumo.ineficiencia_qtd)
    c5.metric("Ineficiência (%)", fmt_pct_valor(resumo.ineficiencia_pct))

    st.markdown("#### Referência de preços no estoque")
    dim = st.selectbox(
        "Recorte",
        ["empreendimento", "tipo", "bloco", "andar", "identificador"],
        key="pc_dim",
    )
    filtro_val = ""
    emp_sub = df_est_enr[df_est_enr["Empreendimento"].map(v._limpar_emp) == v._limpar_emp(emp_sel)] if not df_est_enr.empty else pd.DataFrame()
    if dim == "tipo":
        opts = sorted({_normalizar_tipo(x) for x in emp_sub.get("Tipologia", emp_sub.get("Tipologia__c", pd.Series())).dropna() if str(x).strip()})
        filtro_val = st.selectbox("Tipo / Tipologia", opts or [""], key="pc_tipo")
    elif dim == "bloco":
        blocos = sorted({parse_identificador_unidade(x)["bloco"] for x in emp_sub.get("Identificador", pd.Series()) if parse_identificador_unidade(x)["bloco"]})
        filtro_val = st.selectbox("Bloco", blocos or [""], key="pc_bloco")
    elif dim == "andar":
        andares = sorted({parse_identificador_unidade(x)["andar"] for x in emp_sub.get("Identificador", pd.Series()) if parse_identificador_unidade(x)["andar"]})
        filtro_val = st.selectbox("Andar", andares or [""], key="pc_andar")
    elif dim == "identificador":
        idents = sorted(emp_sub["Identificador"].astype(str).str.upper().unique()) if "Identificador" in emp_sub.columns else []
        filtro_val = st.selectbox("Identificador", idents[:200] or [""], key="pc_ident")

    stats = estatisticas_preco_estoque(
        df_est_enr, emp_sel,
        filtro_dim="" if dim == "empreendimento" else dim,
        filtro_val=filtro_val,
    )
    st.markdown(
        f"Unidades: **{stats['n']}** · Mín: **{v.fmt_br_milhoes(stats['min'])}** · "
        f"Médio: **{v.fmt_br_milhoes(stats['medio'])}** · "
        f"Mediano: **{v.fmt_br_milhoes(stats['mediano'])}** · "
        f"Máx: **{v.fmt_br_milhoes(stats['max'])}**"
    )

    preco_atual = stats["min"] if stats["min"] > 0 else stats["mediano"]
    novo_preco = st.number_input(
        "Simular novo preço (R$)",
        min_value=0.0,
        value=float(preco_atual),
        step=1000.0,
        key="pc_novo_preco",
    )
    if st.button("Calcular impacto", key="pc_sim_btn"):
        sim = simular_novo_preco(
            df_pastas_aprov, df_est_enr, df_vendas, df_tabela,
            emp_sel, novo_preco, preco_atual,
        )
        st.success(
            f"Com preço **{v.fmt_br_milhoes(novo_preco)}** (referência atual **{v.fmt_br_milhoes(preco_atual)}**): "
            f"**{sim['ganhos_pc']}** clientes passariam a ter poder de compra suficiente "
            f"(potencial de **+{sim['vendas_potenciais']}** vendas entre quem ainda não comprou)."
        )
        if sim["detalhe"]:
            exibir_tabela(pd.DataFrame(sim["detalhe"]))


# =============================================================================
# DASHBOARD COMERCIAL (inline — ex-velocimetro_dashboard_comercial.py)
# =============================================================================

# Dashboard comercial — VSO, velocímetros VGV, rankings, share e evolução.

import calendar
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


CANAIS_VGV = [
    ("RIO", "RIO (100%)", None),
    ("DIR", "DIR (50%)", "DIR"),
    ("GC", "GC / RJG (25%)", "RJG"),
    ("PC", "PC / RJ (25%)", "RJ"),
]
FATORES_VGV = {"RIO": 1.0, "DIR": 0.5, "GC": 0.25, "PC": 0.25}
PREFIXOS_IMOB = {"DIR", "RJG", "RJ", "RIV"}
CANAIS_IMOB = {"RJ", "RJG"}
CANAIS_DV = {"DIR", "RIV"}
JANELAS_VSO = (30, 60, 90, 120)

RANKING_PERFIS = {
    "IMOBs": {"col": "Proprietário da oportunidade", "canais": CANAIS_IMOB},
    "Gerentes": {"col": "Proprietário da oportunidade", "canais": CANAIS_DV},
    "Imobiliárias": {"col": "Contato Corretor Proprietario", "canais": CANAIS_IMOB},
    "Corretores": {"col": "Contato Corretor Proprietario", "canais": CANAIS_DV},
    "Regionais": {"col": "Gerente regional", "canais": CANAIS_DV},
}


@dataclass
class FiltrosDashboard:
    data_ini: date
    data_fim: date
    mes_meta: int
    ano_meta: int
    tipo_meta_col: str
    emps_sel: List[str]
    coords_sel: List[str]
    canal_sel: str
    imobs_sel: List[str]


@dataclass
class FiltrosGlobais:
    data_ini: date
    data_fim: date
    mes_meta: int
    ano_meta: int
    tipo_meta_col: str
    tipo_indicador: str
    canal_sel: str
    canal_meta: str
    coords_sel: List[str]
    emps_sel: List[str]
    imobs_sel: List[str]
    status_estoque_sel: List[str]


def filtros_glob_to_dashboard(fg: FiltrosGlobais) -> FiltrosDashboard:
    return FiltrosDashboard(
        data_ini=fg.data_ini,
        data_fim=fg.data_fim,
        mes_meta=fg.mes_meta,
        ano_meta=fg.ano_meta,
        tipo_meta_col=fg.tipo_meta_col,
        emps_sel=fg.emps_sel,
        coords_sel=fg.coords_sel,
        canal_sel=fg.canal_sel,
        imobs_sel=fg.imobs_sel,
    )


def filtros_glob_to_v2(fg: FiltrosGlobais) -> "FiltrosPainelV2":
    return FiltrosPainelV2(
        data_ini=fg.data_ini,
        data_fim=fg.data_fim,
        mes_meta=fg.mes_meta,
        ano_meta=fg.ano_meta,
        tipo_indicador=fg.tipo_indicador,
        tipo_meta_col=fg.tipo_meta_col,
        canal_meta=fg.canal_meta,
        coordenadores_sel=fg.coords_sel,
        emps_sel=fg.emps_sel,
    )


def limpar_caches_velocimetro() -> None:
    """Limpa caches Streamlit (Sheets + SF) para forçar releitura."""
    st.cache_data.clear()
    try:
        st.cache_resource.clear()
    except Exception:
        pass
    vc.limpar_cache_local()


class CachePainelIndisponivel(RuntimeError):
    """Cache de exibição ausente ou expirado — painel não consulta SF ao vivo."""


def _painel_apenas_cache() -> bool:
    """Painel web: sempre Google Sheets (SF só no Actions / pré-compute offline)."""
    return True


def _forcar_sf_painel() -> bool:
    if _painel_apenas_cache():
        return False
    try:
        return bool(st.session_state.get("velocimetro_forcar_sf"))
    except Exception:
        return False


def _resolver_forcar_sf(forcar_sf: Optional[bool]) -> bool:
    if forcar_sf is not None:
        return forcar_sf and not _painel_apenas_cache()
    return _forcar_sf_painel()


def _bloquear_sf_live(dataset: str) -> None:
    raise CachePainelIndisponivel(
        f"Cache «{dataset}» indisponível ou expirado. "
        "Aguarde a sincronização agendada (07h, 10h, 13h, 16h e 19h BRT) "
        "ou use «Atualizar dados» para rodar o pré-compute."
    )


def _manifest_usavel(manifest: Optional[Dict[str, Any]]) -> bool:
    """No modo cache, aceita manifest antigo; fora dele respeita TTL do cache."""
    if not manifest or not manifest.get("atualizado_em"):
        return False
    if _painel_apenas_cache():
        return True
    return vc.manifest_valido(manifest)


def _dataset_cache_existe(info: Dict[str, Any], dataset: str) -> bool:
    df = vc.ler_dataset(dataset, info, prefer_local=False)
    return df is not None and not df.empty


# Fallback: abas «Cache · *» → planilha consolidada (sync SF→Sheets via Actions)
_SHEETS_FALLBACK: Dict[str, Tuple[str, str]] = {
    "vendas_raw": (SPREADSHEET_CONSOLIDADA_ID, WS_VENDAS),
    "estoque": (SPREADSHEET_CONSOLIDADA_ID, WS_ESTOQUE),
    "funil_ag": (SPREADSHEET_CONSOLIDADA_ID, ABA_AGENDAMENTOS_VISITAS),
    "funil_hist_ag": (SPREADSHEET_CONSOLIDADA_ID, ABA_AGENDAMENTOS_VISITAS),
    "funil_emp_ag": (SPREADSHEET_CONSOLIDADA_ID, ABA_AGENDAMENTOS_VISITAS),
    "funil_emp_ven": (SPREADSHEET_CONSOLIDADA_ID, WS_VENDAS),
    "funil_emp_est": (SPREADSHEET_CONSOLIDADA_ID, WS_ESTOQUE),
}


@st.cache_data(ttl=300, show_spinner=False)
def _ler_dado_painel(cache_key: str, _cred_fp: str) -> Tuple[pd.DataFrame, str]:
    """Carrega dataset: 1) aba Cache · *  2) aba consolidada  3) pastas (funil)."""
    info = _info_gsheets_atual()
    if info:
        df = vc.ler_dataset(cache_key, info, prefer_local=False)
        if df is not None and not df.empty:
            m = vc.ler_manifest(info)
            ts = m.get("atualizado_em", "")
            rotulo = vc.DATASETS.get(cache_key, cache_key)
            return df, f"Cache · {rotulo}" + (f" · {ts}" if ts else "")

    sid_ws = _SHEETS_FALLBACK.get(cache_key)
    if sid_ws:
        sid, ws = sid_ws
        try:
            df = normalizar_colunas(ler_planilha_aba_df(sid, ws, _cred_fp))
            if df is not None and not df.empty:
                if cache_key in ("vendas_raw", "funil_emp_ven"):
                    df = _recortar_vendas_painel(df)
                return df, f"Sheets · {ws}"
        except Exception:
            pass

    if cache_key in ("funil_pastas", "funil_hist_pastas", "funil_emp_pastas", "pc_pastas"):
        try:
            df, orig = carregar_df_pastas_funil(
                SPREADSHEET_FUNIL_ID or SPREADSHEET_CONSOLIDADA_ID,
                SPREADSHEET_CONSOLIDADA_ID,
                SPREADSHEET_PASTAS_ID or SPREADSHEET_CONSOLIDADA_ID,
                _cred_fp,
            )
            if df is not None and not df.empty:
                if cache_key == "pc_pastas":
                    df = deduplicar_pastas_aprovadas_funil(df)
                return df, orig or "Sheets · pastas"
        except Exception:
            pass

    return pd.DataFrame(), ""


def _total_unidades_do_manifest(manifest: Dict[str, Any]) -> Dict[str, int]:
    raw = manifest.get("total_unidades_emp_json")
    if not raw:
        return {}
    try:
        data = json.loads(str(raw))
        return {str(_limpar_emp(k)): int(v) for k, v in data.items()}
    except Exception:
        return {}


def _total_unidades_do_estoque_cache() -> Dict[str, int]:
    cf = _cred_fp_atual()
    df = _ler_cache_df("estoque", False, cf)
    if df is None or df.empty or "Empreendimento" not in df.columns:
        return {}
    cont = df.groupby(df["Empreendimento"].map(_limpar_emp)).size()
    return {str(k): int(v) for k, v in cont.items()}


def _info_gsheets_atual() -> Optional[Dict[str, Any]]:
    return montar_service_account_info(_secrets_connections_gsheets())


def _cred_fp_atual() -> str:
    info = _info_gsheets_atual()
    return _fingerprint_credenciais(info) if info else "0"


@st.cache_data(ttl=300, show_spinner=False)
def _ler_cache_df(dataset: str, forcar_sf: bool, _cred_fp: str) -> Optional[pd.DataFrame]:
    if forcar_sf:
        return None
    df, _ = _ler_dado_painel(dataset, _cred_fp)
    return df if df is not None and not df.empty else None


def _manifest_cache_valido(forcar_sf: bool) -> Optional[Dict[str, Any]]:
    if forcar_sf:
        return None
    info = _info_gsheets_atual()
    if not info:
        return None
    manifest = vc.ler_manifest(info)
    if _manifest_usavel(manifest):
        return manifest
    if _painel_apenas_cache() and _dataset_cache_existe(info, "vendas_painel"):
        return manifest or {"atualizado_em": "—"}
    return None


def render_filtros_globais(
    df_metas_coord: pd.DataFrame,
    df_vendas: pd.DataFrame,
    df_metas_fallback: Optional[pd.DataFrame] = None,
) -> FiltrosGlobais:
    """Filtros únicos aplicados em todas as abas."""
    hoje = date.today()
    ini_mes = date(hoje.year, hoje.month, 1)
    coords = _opcoes_unicas(
        _serie_coluna(df_metas_coord, "Coordenador"),
        _serie_coluna(df_metas_fallback, "Coordenador"),
    )
    emps = _opcoes_unicas(
        _serie_coluna(df_metas_coord, "Empreendimento"),
        _serie_coluna(df_metas_fallback, "Empreendimento"),
    )
    imobs: List[str] = []
    if "Imobiliária" in df_vendas.columns:
        imobs = sorted(str(i).strip() for i in df_vendas["Imobiliária"].dropna().unique() if str(i).strip())
    st.markdown("#### Filtros globais")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        data_ini = st.date_input("Data inicial", value=ini_mes, key="glob_data_ini")
    with c2:
        data_fim = st.date_input("Data final", value=hoje, key="glob_data_fim")
    with c3:
        mes_meta = st.selectbox("Mês meta", list(range(1, 13)), index=hoje.month - 1, key="glob_mes_meta")
    with c4:
        ano_meta = st.number_input("Ano meta", min_value=2020, max_value=2035, value=hoje.year, key="glob_ano_meta")
    with c5:
        tipo_meta_col = st.selectbox("Tipo meta", TIPOS_META_COL, key="glob_tipo_meta")
    c6, c7, c8, c9 = st.columns(4)
    with c6:
        canal_sel = st.selectbox("Canal", ["Todos", "RIO", "DIR", "GC", "PC"], key="glob_canal")
    with c7:
        canal_meta = st.selectbox("Canal meta (velocímetro)", CANAIS_META, key="glob_canal_meta")
    with c8:
        tipo_indicador = st.selectbox("Indicador funil/metas", TIPOS_INDICADOR, key="glob_tipo_ind")
    with c9:
        status_estoque_sel = st.multiselect(
            "Status estoque",
            list(ESTOQUE_STATUS_TODOS),
            default=list(ESTOQUE_STATUS_VENDAVEL),
            key="glob_status_est",
        )
    coords_sel = st.multiselect("Coordenador", coords, default=coords, key="glob_coords")
    c10, c11 = st.columns(2)
    with c10:
        emps_sel = st.multiselect("Empreendimento", emps, default=[], key="glob_emps")
    with c11:
        imobs_sel = st.multiselect("Imobiliária", imobs, default=[], key="glob_imobs")
    if st.button("Atualizar dados", type="primary", key="btn_atualizar_dados"):
        limpar_caches_velocimetro()
        st.success("Recarregando bases do Google Sheets…")
        st.rerun()
    return FiltrosGlobais(
        data_ini=data_ini,
        data_fim=data_fim,
        mes_meta=int(mes_meta),
        ano_meta=int(ano_meta),
        tipo_meta_col=tipo_meta_col,
        tipo_indicador=tipo_indicador,
        canal_sel=canal_sel,
        canal_meta=canal_meta,
        coords_sel=coords_sel,
        emps_sel=emps_sel,
        imobs_sel=imobs_sel,
        status_estoque_sel=status_estoque_sel,
    )




def _prefixo_imob(val: Any) -> str:
    return _v().canal_de_imobiliaria(val)




def enriquecer_vendas_vcx(
    df_vendas: pd.DataFrame,
    df_cotacoes: Optional[pd.DataFrame],
) -> pd.DataFrame:
    """Anexa VCX, PS e sinal (última cotação por oportunidade)."""
    out = df_vendas.copy()
    out["Volta ao caixa"] = 0.0
    out["Pct_PS"] = 0.0
    out["PS_VGV"] = 0.0
    out["Total_Sinal"] = 0.0
    out["Sinais_VGV"] = 0.0
    if df_cotacoes is None or df_cotacoes.empty:
        return out
    col_opp = "ID da Oportunidade"
    if col_opp not in df_cotacoes.columns or col_opp not in out.columns:
        return out
    cot = df_cotacoes.copy()
    col_vcx = next((c for c in ("Volta ao caixa", "VoltaAoCaixa__c") if c in cot.columns), None)
    col_ps = next((c for c in ("Percentual Pro Soluto", "PercentualdoProSoluto__c") if c in cot.columns), None)
    col_sinal = next((c for c in ("Total Sinal Com", "TotalSinalCom__c") if c in cot.columns), None)
    if col_vcx:
        cot["_vcx"] = cot[col_vcx].map(_parse_num_br)
    if col_ps:
        cot["_ps"] = pd.to_numeric(cot[col_ps], errors="coerce").fillna(0.0)
    if col_sinal:
        cot["_sinal"] = cot[col_sinal].map(_parse_num_br)
    agg_spec: Dict[str, str] = {}
    if col_vcx:
        agg_spec["_vcx"] = "max"
    if col_ps:
        agg_spec["_ps"] = "max"
    if col_sinal:
        agg_spec["_sinal"] = "max"
    if not agg_spec:
        return out
    g = cot.groupby(col_opp, as_index=False).agg(agg_spec)
    out = out.merge(g, on=col_opp, how="left")
    if "_vcx" in out.columns:
        out["Volta ao caixa"] = pd.to_numeric(out["_vcx"], errors="coerce").fillna(0.0)
        out = out.drop(columns=["_vcx"], errors="ignore")
    if "_ps" in out.columns:
        out["Pct_PS"] = out["_ps"].fillna(0.0)
        out["PS_VGV"] = out["Pct_PS"] / 100.0
        out = out.drop(columns=["_ps"], errors="ignore")
    if "_sinal" in out.columns:
        out["Total_Sinal"] = out["_sinal"].fillna(0.0)
        out = out.drop(columns=["_sinal"], errors="ignore")
    vgv_num = pd.to_numeric(
        out["_vgv"] if "_vgv" in out.columns else out.get("Valor Real de Venda", 0),
        errors="coerce",
    ).fillna(0.0)
    out["Sinais_VGV"] = np.where(vgv_num > 0, out["Total_Sinal"] / vgv_num, 0.0)
    return out


def _col_data_venda(df: pd.DataFrame) -> str:
    v = _v()
    return v.achar_coluna(df, ["Data da venda", "Data Venda", "Data de venda"]) or ""


def _aplicar_filtros_base(
    df: pd.DataFrame,
    filtros: FiltrosDashboard,
    mapa_coord: Dict[str, str],
    col_data: str,
    usar_periodo: bool = True,
) -> pd.DataFrame:
    v = _v()
    base = df.copy()
    if filtros.emps_sel and "Empreendimento" in base.columns:
        emps = {_limpar_emp(e) for e in filtros.emps_sel}
        base = base[base["Empreendimento"].map(_limpar_emp).isin(emps)]
    if filtros.coords_sel and mapa_coord and "Empreendimento" in base.columns:
        emps_coord = {
            e for e, c in mapa_coord.items() if c in set(filtros.coords_sel)
        }
        base = base[base["Empreendimento"].map(_limpar_emp).isin(emps_coord)]
    if filtros.imobs_sel and "Imobiliária" in base.columns:
        imobs = {str(i).strip() for i in filtros.imobs_sel}
        base = base[base["Imobiliária"].astype(str).str.strip().isin(imobs)]
    canal = (filtros.canal_sel or "Todos").strip().upper()
    if canal not in ("TODOS", "RIO", "") and "Imobiliária" in base.columns:
        cfg = next((c for c in CANAIS_VGV if c[0] == canal), None)
        if cfg and cfg[2]:
            base = base[base["Imobiliária"].map(_prefixo_imob) == cfg[2]]
        elif canal == "DIR":
            base = base[base["Imobiliária"].map(_prefixo_imob).isin(CANAIS_DV)]
        elif canal in ("GC", "PARC"):
            base = base[base["Imobiliária"].map(_prefixo_imob) == "RJG"]
        elif canal == "PC":
            base = base[base["Imobiliária"].map(_prefixo_imob) == "RJ"]
    if usar_periodo and col_data and col_data in base.columns:
        base = _filtrar_df_periodo(base, col_data, filtros.data_ini, filtros.data_fim)
    return base


def _filtrar_canal_velocimetro(df: pd.DataFrame, canal_key: str) -> pd.DataFrame:
    if df.empty or "Imobiliária" not in df.columns:
        return df
    cfg = next((c for c in CANAIS_VGV if c[0] == canal_key), None)
    if not cfg:
        return df
    if cfg[2] is None:
        return df
    return df[df["Imobiliária"].map(_prefixo_imob) == cfg[2]]


def _qtd_vgv(df: pd.DataFrame) -> Tuple[float, float]:
    qtd = (
        _sum_col_num(df, "_qtd_venda", float(len(df)))
        if "_qtd_venda" in df.columns
        else float(len(df))
    )
    if "_vgv_venda" in df.columns:
        vgv = _sum_col_num(df, "_vgv_venda", 0.0)
    elif "_vgv" in df.columns:
        vgv = _sum_col_num(df, "_vgv", 0.0)
    else:
        vgv = 0.0
    return qtd, vgv


def calcular_vso_por_emp(
    df_vendas: pd.DataFrame,
    df_estoque: pd.DataFrame,
    filtros: FiltrosDashboard,
    mapa_coord: Dict[str, str],
    col_data: str,
    ref_fim: Optional[date] = None,
) -> pd.DataFrame:
    """VSO = vendas / (vendas + unidades) por janela rolling."""
    v = _v()
    ref_fim = ref_fim or filtros.data_fim
    emps = sorted(set(mapa_coord.keys()))
    if filtros.emps_sel:
        emps = [_limpar_emp(e) for e in filtros.emps_sel]
    if filtros.coords_sel:
        emps = [e for e in emps if mapa_coord.get(e, "") in filtros.coords_sel]

    est = df_estoque.copy() if df_estoque is not None else pd.DataFrame()
    if not est.empty and "Empreendimento" in est.columns:
        est["_emp"] = est["Empreendimento"].map(_limpar_emp)
        if filtros.emps_sel:
            sel = {_limpar_emp(e) for e in filtros.emps_sel}
            est = est[est["_emp"].isin(sel)]
        unidades_por_emp = est.groupby("_emp").size().to_dict()
    else:
        unidades_por_emp = {}

    rows = []
    for emp in emps:
        row: Dict[str, Any] = {"Empreendimento": emp, "Coordenador": mapa_coord.get(emp, "")}
        unidades = float(unidades_por_emp.get(emp, 0))
        row["Unidades_Estoque"] = int(unidades)
        for dias in JANELAS_VSO:
            ini = ref_fim - timedelta(days=dias)
            sub = df_vendas.copy()
            if "Empreendimento" in sub.columns:
                sub = sub[sub["Empreendimento"].map(_limpar_emp) == emp]
            if filtros.imobs_sel and "Imobiliária" in sub.columns:
                imobs = {str(i).strip() for i in filtros.imobs_sel}
                sub = sub[sub["Imobiliária"].astype(str).str.strip().isin(imobs)]
            canal = (filtros.canal_sel or "Todos").strip().upper()
            if canal not in ("TODOS", "RIO", "") and "Imobiliária" in sub.columns:
                sub = _filtrar_canal_velocimetro(sub, canal if canal in FATORES_VGV else "RIO")
            if col_data and col_data in sub.columns:
                sub = _filtrar_df_periodo(sub, col_data, ini, ref_fim)
            vendas = (
                _sum_col_num(sub, "_qtd_venda", float(len(sub)))
                if "_qtd_venda" in sub.columns
                else float(len(sub))
            )
            denom = vendas + unidades
            row[f"VSO_{dias}d"] = (vendas / denom * 100.0) if denom > 0 else 0.0
            row[f"Vendas_{dias}d"] = vendas
        rows.append(row)
    return pd.DataFrame(rows)


def _meta_vgv_canal(
    df_canal: pd.DataFrame,
    mes: int,
    ano: int,
    canal_key: str,
) -> Tuple[float, float]:
    fator = FATORES_VGV.get(canal_key, 1.0)
    vgv, ven = meta_canal_vgv_vendas(df_canal, mes, ano, "RIO")
    return vgv * fator, ven * fator


def _meta_mensal_acumulada(
    df_canal: pd.DataFrame,
    ano: int,
    ate_mes: int,
    canal_key: str = "RIO",
) -> pd.DataFrame:
    """Meta VGV mês a mês (acumulada) até ate_mes."""
    rows = []
    acum = 0.0
    for m in range(1, ate_mes + 1):
        vgv, _ = _meta_vgv_canal(df_canal, m, ano, canal_key)
        acum += vgv
        rows.append({"Mes": m, "Meta_VGV_Acum": acum, "Meta_VGV_Mes": vgv})
    return pd.DataFrame(rows)


def _realizado_mensal_acumulado(
    df: pd.DataFrame,
    col_data: str,
    ano: int,
    ate_mes: int,
) -> pd.DataFrame:
    v = _v()
    if df.empty or not col_data:
        return pd.DataFrame()
    base = df.copy()
    base["_dt"] = v.parse_data_serie(base[col_data])
    base = base.dropna(subset=["_dt"])
    base = base[base["_dt"].dt.year == ano]
    base = base[base["_dt"].dt.month <= ate_mes]
    col_vgv = "_vgv_venda" if "_vgv_venda" in base.columns else "_vgv"
    rows = []
    acum = 0.0
    for m in range(1, ate_mes + 1):
        sub = base[base["_dt"].dt.month == m]
        mes_vgv = float(sub[col_vgv].sum()) if col_vgv in sub.columns else 0.0
        acum += mes_vgv
        rows.append({"Mes": m, "Real_VGV_Acum": acum, "Real_VGV_Mes": mes_vgv})
    return pd.DataFrame(rows)


def calcular_gap_meta(
    df_vendas: pd.DataFrame,
    df_canal: pd.DataFrame,
    filtros: FiltrosDashboard,
    col_data: str,
) -> Dict[str, float]:
    hoje = date.today()
    if hoje.month == 1:
        ate_mes = 12
        ano_ref = hoje.year - 1
    else:
        ate_mes = hoje.month - 1
        ano_ref = hoje.year
    meta_df = _meta_mensal_acumulada(df_canal, ano_ref, ate_mes, "RIO")
    meta_total = float(meta_df["Meta_VGV_Acum"].iloc[-1]) if not meta_df.empty else 0.0
    base = df_vendas.copy()
    if col_data:
        v = _v()
        base["_dt"] = v.parse_data_serie(base[col_data])
        base = base.dropna(subset=["_dt"])
        base = base[(base["_dt"].dt.year == ano_ref) & (base["_dt"].dt.month <= ate_mes)]
    col_vgv = "_vgv_venda" if "_vgv_venda" in base.columns else "_vgv"
    real_total = float(base[col_vgv].sum()) if col_vgv in base.columns else 0.0
    return {
        "meta_acum": meta_total,
        "real_acum": real_total,
        "gap": meta_total - real_total,
        "ate_mes": ate_mes,
        "ano_ref": ano_ref,
    }


def share_por_canal(
    df: pd.DataFrame,
) -> pd.DataFrame:
    if df.empty or "Imobiliária" not in df.columns:
        return pd.DataFrame()
    base = df.copy()
    base["_pfx"] = base["Imobiliária"].map(_prefixo_imob)
    base["_canal_grp"] = base["_pfx"].apply(
        lambda p: "DIR" if p in CANAIS_DV else ("GC" if p == "RJG" else ("PC" if p == "RJ" else "Outros"))
    )
    col_vgv = "_vgv_venda" if "_vgv_venda" in base.columns else "_vgv"
    agg = base.groupby("_canal_grp")[col_vgv].sum().reset_index()
    agg.columns = ["Canal", "VGV"]
    total = float(agg["VGV"].sum())
    agg["Share_Real"] = agg["VGV"] / total * 100.0 if total > 0 else 0.0
    meta_share = {"DIR": 50.0, "GC": 25.0, "PC": 25.0}
    agg["Share_Meta"] = agg["Canal"].map(meta_share).fillna(0.0)
    return agg


def share_por_diretor(df: pd.DataFrame) -> pd.DataFrame:
    col = "Diretor de vendas"
    if col not in df.columns:
        for c in ("Diretor_de_vendas__c", "Diretor"):
            if c in df.columns:
                col = c
                break
        else:
            return pd.DataFrame()
    base = df.copy()
    base[col] = base[col].fillna("(sem diretor)").astype(str).str.strip()
    col_vgv = "_vgv_venda" if "_vgv_venda" in base.columns else "_vgv"
    agg = base.groupby(col)[col_vgv].sum().reset_index()
    agg.columns = ["Diretor", "VGV"]
    total = float(agg["VGV"].sum())
    agg["Share"] = agg["VGV"] / total * 100.0 if total > 0 else 0.0
    return agg.sort_values("VGV", ascending=False)


def montar_ranking(
    df: pd.DataFrame,
    perfil: str,
    col_data: str,
    data_ini: date,
    data_fim: date,
    metrica: str = "VGV",
) -> pd.DataFrame:
    cfg = RANKING_PERFIS.get(perfil)
    if not cfg or df.empty:
        return pd.DataFrame()
    col_nome = cfg["col"]
    if col_nome not in df.columns or "Imobiliária" not in df.columns:
        return pd.DataFrame()
    base = df.copy()
    base["_pfx"] = base["Imobiliária"].map(_prefixo_imob)
    base = base[base["_pfx"].isin(cfg["canais"])]
    if col_data and col_data in base.columns:
        base = _filtrar_df_periodo(base, col_data, data_ini, data_fim)
    base[col_nome] = base[col_nome].fillna("(vazio)").astype(str).str.strip()
    base = base[base[col_nome] != "(vazio)"]
    if base.empty:
        return pd.DataFrame()
    if metrica == "Quantidade":
        agg = base.groupby(col_nome).agg(
            Qtd=("_qtd_venda", "sum") if "_qtd_venda" in base.columns else (col_nome, "count"),
            VGV=("_vgv_venda", "sum") if "_vgv_venda" in base.columns else ("_vgv", "sum"),
        ).reset_index()
        agg = agg.rename(columns={col_nome: "Nome"})
        agg = agg.sort_values("Qtd", ascending=False)
    else:
        col_vgv = "_vgv_venda" if "_vgv_venda" in base.columns else "_vgv"
        agg = base.groupby(col_nome).agg(
            VGV=(col_vgv, "sum"),
            Qtd=("_qtd_venda", "sum") if "_qtd_venda" in base.columns else (col_nome, "count"),
        ).reset_index()
        agg = agg.rename(columns={col_nome: "Nome"})
        agg = agg.sort_values("VGV", ascending=False)
    agg["Ranking"] = range(1, len(agg) + 1)
    return agg


def render_filtros_dashboard(
    df_metas_coord: pd.DataFrame,
    df_vendas: pd.DataFrame,
    filtros_externos: Optional["FiltrosGlobais"] = None,
) -> FiltrosDashboard:
    if filtros_externos is not None:
        return filtros_glob_to_dashboard(filtros_externos)
    hoje = date.today()
    ini_mes = date(hoje.year, hoje.month, 1)
    coords = _opcoes_unicas(
        _serie_coluna(df_metas_coord, "Coordenador"),
    )
    emps = _opcoes_unicas(
        _serie_coluna(df_metas_coord, "Empreendimento"),
    )
    imobs: List[str] = []
    if "Imobiliária" in df_vendas.columns:
        imobs = sorted(
            str(i).strip()
            for i in df_vendas["Imobiliária"].dropna().unique()
            if str(i).strip()
        )
    st.markdown("#### Filtros do dashboard comercial")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        data_ini = st.date_input("Data inicial", value=ini_mes, key="dc_data_ini")
    with c2:
        data_fim = st.date_input("Data final", value=hoje, key="dc_data_fim")
    with c3:
        mes_meta = st.selectbox("Mês meta", list(range(1, 13)), index=hoje.month - 1, key="dc_mes_meta")
    with c4:
        ano_meta = st.number_input("Ano meta", min_value=2020, max_value=2035, value=hoje.year, key="dc_ano_meta")
    with c5:
        tipo_meta_col = st.selectbox("Tipo meta", TIPOS_META_COL, key="dc_tipo_meta")
    c6, c7, c8 = st.columns(3)
    with c6:
        coords_sel = st.multiselect("Coordenador", coords, default=coords, key="dc_coords")
    with c7:
        canal_sel = st.selectbox(
            "Canal",
            ["Todos", "RIO", "DIR", "GC", "PC"],
            key="dc_canal",
        )
    with c8:
        emps_sel = st.multiselect("Empreendimento", emps, default=[], key="dc_emps")
    imobs_sel = st.multiselect("Imobiliária", imobs, default=[], key="dc_imobs")
    return FiltrosDashboard(
        data_ini=data_ini,
        data_fim=data_fim,
        mes_meta=int(mes_meta),
        ano_meta=int(ano_meta),
        tipo_meta_col=tipo_meta_col,
        emps_sel=emps_sel,
        coords_sel=coords_sel,
        canal_sel=canal_sel,
        imobs_sel=imobs_sel,
    )


def render_tabela_vso_meta(
    df_vso: pd.DataFrame,
    df_metas_coord: pd.DataFrame,
    df_vendas_f: pd.DataFrame,
    filtros: FiltrosDashboard,
    col_data: str,
) -> None:
    v = _v()
    st.subheader("VSO, meta e % meta por empreendimento")
    if df_vso.empty:
        st.info("Sem empreendimentos para exibir.")
        return
    rows = []
    for _, r in df_vso.iterrows():
        emp = r["Empreendimento"]
        meta = soma_meta_coord(
            df_metas_coord, filtros.mes_meta, filtros.ano_meta,
            "vendas", filtros.tipo_meta_col, empreendimentos=[emp],
        )
        sub = df_vendas_f[df_vendas_f["Empreendimento"].map(_limpar_emp) == emp] if not df_vendas_f.empty else pd.DataFrame()
        real_qtd, real_vgv = _qtd_vgv(sub)
        pct_meta = (real_qtd / meta * 100.0) if meta > 0 else 0.0
        row = {
            "Empreendimento": emp,
            "Coordenador": r.get("Coordenador", ""),
            "Meta_Vendas": meta,
            "Realizado_Vendas": real_qtd,
            "Pct_Meta": round(pct_meta, 1),
            "Unidades_Estoque": r.get("Unidades_Estoque", 0),
        }
        for dias in JANELAS_VSO:
            row[f"VSO_{dias}d"] = round(float(r.get(f"VSO_{dias}d", 0)), 2)
        rows.append(row)
    df_out = pd.DataFrame(rows)
    exibir_tabela(df_out)


def render_velocimetros_vgv(
    df_vendas: pd.DataFrame,
    df_canal: pd.DataFrame,
    filtros: FiltrosDashboard,
    col_data: str,
) -> None:
    v = _v()
    st.subheader("Velocímetros VGV por canal")
    base = df_vendas.copy()
    if col_data and col_data in base.columns:
        base = _filtrar_df_periodo(base, col_data, filtros.data_ini, filtros.data_fim)

    c_top, _, c_top2 = st.columns([1, 2, 1])
    vgv_rio, ven_rio = _meta_vgv_canal(df_canal, filtros.mes_meta, filtros.ano_meta, "RIO")
    qtd_rio, real_vgv_rio = _qtd_vgv(base)
    with c_top:
        v.criar_medidor(
            "RIO · 100%", qtd_rio, ven_rio, real_vgv_rio, vgv_rio, qtd_rio, mostrar_vgv=True,
        )

    cols = st.columns(3)
    for i, (key, titulo, _) in enumerate(CANAIS_VGV[1:]):
        sub = _filtrar_canal_velocimetro(base, key)
        qtd, real_vgv = _qtd_vgv(sub)
        meta_vgv, meta_qtd = _meta_vgv_canal(df_canal, filtros.mes_meta, filtros.ano_meta, key)
        with cols[i]:
            v.criar_medidor(
                titulo, qtd, meta_qtd, real_vgv, meta_vgv, qtd, mostrar_vgv=True,
            )


def render_velocimetros_coordenador_vgv(
    df_vendas: pd.DataFrame,
    df_metas_coord: pd.DataFrame,
    df_canal: pd.DataFrame,
    filtros: FiltrosDashboard,
    mapa_coord: Dict[str, str],
    col_data: str,
) -> None:
    v = _v()
    st.subheader("Velocímetro VGV por coordenador")
    coords = filtros.coords_sel or sorted(set(mapa_coord.values()))
    if not coords:
        st.info("Nenhum coordenador disponível.")
        return
    base = df_vendas.copy()
    if col_data and col_data in base.columns:
        base = _filtrar_df_periodo(base, col_data, filtros.data_ini, filtros.data_fim)
    cols = st.columns(min(3, len(coords)) or 1)
    for i, coord in enumerate(coords):
        emps = empreendimentos_de_coord(mapa_coord, [coord])
        if filtros.emps_sel:
            emps = [e for e in emps if e in {_limpar_emp(x) for x in filtros.emps_sel}]
        sub = base[base["Empreendimento"].map(_limpar_emp).isin(emps)] if emps else base.iloc[0:0]
        qtd, real_vgv = _qtd_vgv(sub)
        meta_qtd = soma_meta_coord(
            df_metas_coord, filtros.mes_meta, filtros.ano_meta,
            "vendas", filtros.tipo_meta_col,
        )
        meta_vgv_coord = soma_meta_vgv_coord(
            df_metas_coord, filtros.mes_meta, filtros.ano_meta, "Desafio",
            coordenadores=[coord], empreendimentos=emps or None,
        )
        with cols[i % len(cols)]:
            v.criar_medidor(
                coord, qtd, meta_qtd, real_vgv, meta_vgv_coord, qtd,
                mostrar_vgv=True, metrica="vgv",
            )


def render_grafico_share_canal(df: pd.DataFrame) -> None:
    st.subheader("Share por canal — realizado x meta")
    agg = share_por_canal(df)
    if agg.empty:
        st.info("Sem dados para share por canal.")
        return
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Share realizado (%)", x=agg["Canal"], y=agg["Share_Real"], marker_color="#2563eb"))
    fig.add_trace(go.Bar(name="Share meta (%)", x=agg["Canal"], y=agg["Share_Meta"], marker_color="#94a3b8"))
    fig.update_layout(barmode="group", height=380, margin=dict(l=20, r=20, t=40, b=20))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def render_grafico_evolucao_vgv(
    df_vendas: pd.DataFrame,
    df_canal: pd.DataFrame,
    filtros: FiltrosDashboard,
    col_data: str,
) -> None:
    st.subheader("Evolução VGV realizado x meta (Data da venda)")
    if not col_data:
        st.warning("Coluna Data da venda não encontrada.")
        return
    hoje = date.today()
    ate_mes = hoje.month
    real_df = _realizado_mensal_acumulado(df_vendas, col_data, filtros.ano_meta, ate_mes)
    meta_df = _meta_mensal_acumulada(df_canal, filtros.ano_meta, ate_mes, "RIO")
    if real_df.empty and meta_df.empty:
        st.info("Sem dados de evolução.")
        return
    meses = list(range(1, ate_mes + 1))
    nomes = [calendar.month_abbr[m] for m in meses]
    fig = go.Figure()
    if not meta_df.empty:
        fig.add_trace(go.Scatter(
            x=nomes, y=meta_df["Meta_VGV_Acum"], mode="lines+markers",
            name="Meta VGV acum.", line=dict(color="#64748b", dash="dash"),
        ))
    if not real_df.empty:
        fig.add_trace(go.Scatter(
            x=nomes, y=real_df["Real_VGV_Acum"], mode="lines+markers",
            name="Realizado VGV acum.", line=dict(color="#2563eb"),
        ))
    fig.update_layout(height=400, margin=dict(l=20, r=20, t=40, b=20), yaxis_title="VGV (R$)")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def render_box_gap(
    df_vendas: pd.DataFrame,
    df_canal: pd.DataFrame,
    filtros: FiltrosDashboard,
    col_data: str,
) -> None:
    v = _v()
    gap = calcular_gap_meta(df_vendas, df_canal, filtros, col_data)
    st.markdown("##### Gap disponível para meta")
    cor = "#16a34a" if gap["gap"] >= 0 else "#cb0935"
    st.markdown(
        f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Meta acum. até {gap['ate_mes']:02d}/{gap['ano_ref']}</div>
            <div class="val">{v.fmt_br_milhoes(gap['meta_acum'])}</div></div>
            <div class="vel-kpi"><div class="lbl">Realizado acum.</div>
            <div class="val">{v.fmt_br_milhoes(gap['real_acum'])}</div></div>
            <div class="vel-kpi"><div class="lbl">Gap disponível para meta</div>
            <div class="val" style="color:{cor}">{v.fmt_br_milhoes(gap['gap'])}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_rankings(
    df_vendas: pd.DataFrame,
    col_data: str,
    filtros: FiltrosDashboard,
) -> None:
    st.subheader("Rankings por data da venda")
    c1, c2 = st.columns(2)
    with c1:
        perfil = st.selectbox("Perfil", list(RANKING_PERFIS.keys()), key="dc_rank_perfil")
    with c2:
        metrica = st.selectbox("Indicador", ["VGV", "Quantidade"], key="dc_rank_metrica")
    rank = montar_ranking(
        df_vendas, perfil, col_data, filtros.data_ini, filtros.data_fim, metrica,
    )
    if rank.empty:
        st.info("Sem dados para o ranking selecionado.")
        return
    exibir_tabela(rank.head(50))


def render_share_diretor(df: pd.DataFrame) -> None:
    st.subheader("Share por diretor")
    agg = share_por_diretor(df)
    if agg.empty:
        st.info("Coluna Diretor de vendas não disponível na base.")
        return
    fig = go.Figure(go.Pie(labels=agg["Diretor"], values=agg["VGV"], hole=0.4))
    fig.update_layout(height=400, margin=dict(l=20, r=20, t=40, b=20))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    exibir_tabela(agg)


def render_tabela_detalhada(
    df: pd.DataFrame,
    col_data: str,
    filtros: FiltrosDashboard,
) -> None:
    v = _v()
    st.subheader("Detalhamento de vendas")
    c1, c2 = st.columns(2)
    with c1:
        dt_ini = st.date_input("Data venda — de", value=filtros.data_ini, key="dc_det_ini")
    with c2:
        dt_fim = st.date_input("Data venda — até", value=filtros.data_fim, key="dc_det_fim")
    base = df.copy()
    if col_data and col_data in base.columns:
        base = _filtrar_df_periodo(base, col_data, dt_ini, dt_fim)
    col_val = v.achar_coluna(base, ["Valor Real de Venda", "Valor Real", "Valor"])
    cols_out = []
    for c in ("Canal", "Empreendimento", "Imobiliária", "Ranking", col_val, "Volta ao caixa"):
        if c and c in base.columns:
            cols_out.append(c)
    if not cols_out:
        st.info("Colunas insuficientes para a tabela detalhada.")
        return
    out = base[cols_out].copy()
    if col_val and col_val in out.columns:
        out = out.rename(columns={col_val: "Valor Venda"})
    if "Canal" not in out.columns and "Imobiliária" in out.columns:
        out["Canal"] = out["Imobiliária"].map(_prefixo_imob)
    exibir_tabela(out)


# Metas de conversão de referência (aba Radar do Polaroid)
META_CONV_POLAROID = {
    "Visita → Pasta": 0.40,
    "Pasta → Pasta aprovada": 0.50,
    "Pasta aprovada → Venda": 0.25,
    "Visita → Venda": 0.10,
}


def _contar_funil_mtd(
    df: pd.DataFrame,
    col_data: str,
    mes: int,
    ano: int,
) -> float:
    v = _v()
    if df is None or df.empty or not col_data or col_data not in df.columns:
        return 0.0
    sub = df.copy()
    sub["_dt"] = v.parse_data_serie(sub[col_data])
    sub = sub.dropna(subset=["_dt"])
    sub = sub[(sub["_dt"].dt.month == mes) & (sub["_dt"].dt.year == ano)]
    return float(len(sub))


def render_radar_polaroid(
    df_vendas: pd.DataFrame,
    df_estoque: pd.DataFrame,
    df_canal: pd.DataFrame,
    df_metas_coord: pd.DataFrame,
    filtros: FiltrosDashboard,
    mapa_coord: Dict[str, str],
    col_data: str,
    df_ag: Optional[pd.DataFrame] = None,
    df_pastas: Optional[pd.DataFrame] = None,
) -> None:
    """Painel estilo aba Radar do Polaroid — metas, realizado, estoque, funil e pipeline."""
    v = _v()
    hoje = date.today()
    mes, ano = filtros.mes_meta, filtros.ano_meta
    ini_mes = date(ano, mes, 1)
    if mes == 12:
        fim_mes = date(ano, 12, 31)
    else:
        fim_mes = date(ano, mes + 1, 1) - timedelta(days=1)

    meta_vgv_desafio, _ = _meta_vgv_canal(df_canal, mes, ano, "RIO")
    meta_vgv_bp, _ = meta_canal_vgv_vendas(df_canal, mes, ano, "RIO")
    meta_qtd = soma_meta_coord(
        df_metas_coord, mes, ano, "vendas", "Desafio",
        coordenadores=filtros.coords_sel or None,
        empreendimentos=filtros.emps_sel or None,
    )
    meta_qtd_bp = soma_meta_coord(
        df_metas_coord, mes, ano, "vendas", "BP",
        coordenadores=filtros.coords_sel or None,
        empreendimentos=filtros.emps_sel or None,
    )
    esc = escala_meta_regressao(None, meta_vgv_desafio, hoje if mes == hoje.month and ano == hoje.year else fim_mes)

    base_mes = df_vendas.copy()
    if col_data and col_data in base_mes.columns:
        base_mes = _filtrar_df_periodo(base_mes, col_data, ini_mes, fim_mes)
    if "_vgv_venda" in base_mes.columns:
        real_vgv = _sum_col_num(base_mes, "_vgv_venda", 0.0)
    elif "_vgv" in base_mes.columns:
        real_vgv = _sum_col_num(base_mes, "_vgv", 0.0)
    else:
        real_vgv = 0.0
    real_qtd = (
        _sum_col_num(base_mes, "_qtd_venda", float(len(base_mes)))
        if "_qtd_venda" in base_mes.columns
        else float(len(base_mes))
    )

    pct_desafio = (real_vgv / meta_vgv_desafio * 100.0) if meta_vgv_desafio > 0 else 0.0
    pct_bp = (real_vgv / meta_vgv_bp * 100.0) if meta_vgv_bp > 0 else 0.0

    kpi_est, _ = agregar_estoque(df_estoque if df_estoque is not None else pd.DataFrame())
    vso = real_qtd / (real_qtd + float(kpi_est.get("unidades", 0))) if (real_qtd + kpi_est.get("unidades", 0)) > 0 else 0.0

    n_comunicadas = 0
    for col in ("Venda comunicada", "GeradoComunicadoVenda__c"):
        if col in base_mes.columns:
            n_comunicadas = int(base_mes[col].astype(str).str.upper().isin(("TRUE", "1", "SIM", "YES")).sum())
            break

    col_vis = v.achar_coluna(_coalesce_df(df_ag), v.ALIASES_DATA_VISITA) or "Data da visita"
    col_pas = v.achar_coluna_primeiro_envio_analise(_coalesce_df(df_pastas)) or v.achar_coluna(
        _coalesce_df(df_pastas), v.ALIASES_DATA_CRIACAO
    )
    col_apr = v.achar_coluna_aprovacao_safi(_coalesce_df(df_pastas))
    n_visitas = _contar_funil_mtd(df_ag, col_vis, mes, ano) if col_vis else 0.0
    n_pastas = _contar_funil_mtd(df_pastas, col_pas, mes, ano) if col_pas else 0.0
    n_aprov = _contar_funil_mtd(df_pastas, col_apr, mes, ano) if col_apr else 0.0

    st.subheader("Radar — visão Polaroid")
    st.caption(f"Referência: Polaroid RJ · {mes:02d}/{ano} · metas comercial (desafio) vs BP")
    st.markdown(
        f"""
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Meta Comercial (VGV)</div>
            <div class="val">{v.fmt_br_milhoes(meta_vgv_desafio)}</div></div>
            <div class="vel-kpi"><div class="lbl">Meta BP (VGV)</div>
            <div class="val">{v.fmt_br_milhoes(meta_vgv_bp)}</div></div>
            <div class="vel-kpi"><div class="lbl">Meta acum. dia</div>
            <div class="val">{v.fmt_br_milhoes(esc['meta_acum_hoje'])}</div></div>
            <div class="vel-kpi"><div class="lbl">Realizado mês (VGV)</div>
            <div class="val">{v.fmt_br_milhoes(real_vgv)}</div></div>
        </div>
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">% Meta Comercial</div>
            <div class="val">{fmt_pct_valor(pct_desafio)}</div></div>
            <div class="vel-kpi"><div class="lbl">% Meta BP</div>
            <div class="val">{fmt_pct_valor(pct_bp)}</div></div>
            <div class="vel-kpi"><div class="lbl">VSO mês</div>
            <div class="val">{fmt_pct_valor(vso * 100)}</div></div>
            <div class="vel-kpi"><div class="lbl">Comunicadas</div>
            <div class="val">{n_comunicadas}</div></div>
        </div>
        <div class="vel-kpi-row">
            <div class="vel-kpi"><div class="lbl">Estoque (unid.)</div>
            <div class="val">{int(kpi_est.get('unidades', 0))}</div></div>
            <div class="vel-kpi"><div class="lbl">Estoque VGV</div>
            <div class="val">{v.fmt_br_milhoes(float(kpi_est.get('vgv', 0)))}</div></div>
            <div class="vel-kpi"><div class="lbl">Ticket estoque</div>
            <div class="val">{v.fmt_br_milhoes(float(kpi_est.get('ticket', 0)))}</div></div>
            <div class="vel-kpi"><div class="lbl">Vendas (qtd mês)</div>
            <div class="val">{v.fmt_qtd(real_qtd)}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("##### Funil do mês")
        st.markdown(
            f"- Visitas: **{int(n_visitas)}**\n"
            f"- Pastas: **{int(n_pastas)}**\n"
            f"- Pastas aprovadas: **{int(n_aprov)}**\n"
            f"- Vendas: **{int(real_qtd)}**"
        )
    with c2:
        st.markdown("##### Metas de conversão (Polaroid)")
        for nome, ref in META_CONV_POLAROID.items():
            st.markdown(f"- {nome}: **{ref:.0%}**")


def montar_tabela_radar_emp(
    df_vendas: pd.DataFrame,
    df_estoque: pd.DataFrame,
    df_metas_coord: pd.DataFrame,
    df_pastas_aprov: pd.DataFrame,
    df_tabela_comp: pd.DataFrame,
    filtros: FiltrosDashboard,
    mapa_coord: Dict[str, str],
    col_data: str,
) -> pd.DataFrame:
    """Tabela por empreendimento inspirada na aba Radar (Polaroid)."""
    v = _v()
    hoje = date.today()
    mes, ano = filtros.mes_meta, filtros.ano_meta
    ini_mes = date(ano, mes, 1)
    fim_mes = date(ano, mes, calendar.monthrange(ano, mes)[1])
    dias_mes = calendar.monthrange(ano, mes)[1]
    dia_ref = hoje.day if mes == hoje.month and ano == hoje.year else dias_mes

    _, enr = agregar_estoque(df_estoque if df_estoque is not None else pd.DataFrame())
    resumo_est = resumo_estoque_por_emp(enr)
    emps = sorted(set(mapa_coord.keys()))
    if filtros.emps_sel:
        emps = [_limpar_emp(e) for e in filtros.emps_sel]
    if filtros.coords_sel:
        emps = [e for e in emps if mapa_coord.get(e, "") in filtros.coords_sel]

    rows = []
    for emp in emps:
        meta_m = soma_meta_coord(
            df_metas_coord, mes, ano, "vendas", filtros.tipo_meta_col, empreendimentos=[emp],
        )
        meta_dia = meta_m * dia_ref / dias_mes if dias_mes else 0.0
        ven_qtd, ven_vgv = realizado_vendas_periodo(
            df_vendas, col_data or "", ini_mes, fim_mes, [emp],
        )
        pct_dia = (ven_qtd / meta_dia * 100.0) if meta_dia > 0 else 0.0
        pct_mes = (ven_qtd / meta_m * 100.0) if meta_m > 0 else 0.0
        rs = resumo_est.loc[resumo_est["Empreendimento"] == emp] if not resumo_est.empty else pd.DataFrame()
        est_u = int(rs["Unidades"].iloc[0]) if not rs.empty else 0
        diff_av = float(rs["Diff_Avaliacao"].iloc[0]) if not rs.empty and "Diff_Avaliacao" in rs.columns else 0.0
        m2 = float(rs["m2_Total"].iloc[0]) if not rs.empty else 0.0
        res_pc = calcular_resumo_ineficiencia_emp(
            df_pastas_aprov, enr, df_vendas, df_tabela_comp, emp,
        )
        ven_fut = 0.0
        if col_data:
            sub = _filtrar_df_periodo(df_vendas, col_data, ini_mes, fim_mes)
            if not sub.empty and "Empreendimento" in sub.columns:
                ve = sub[sub["Empreendimento"].map(_limpar_emp) == emp]
                for col in ("Venda futura", "Venda_Futura__c"):
                    if col in ve.columns:
                        ven_fut = float(ve[col].astype(str).str.upper().isin(("TRUE", "1", "SIM", "YES")).sum())
                        break
        rows.append({
            "Empreendimento": emp,
            "Coordenador": mapa_coord.get(emp, ""),
            "Estoque_Un": est_u,
            "Diff_Avaliacao": round(diff_av, 0),
            "m2_Total": round(m2, 1),
            "Realizado_Vendas": ven_qtd,
            "Meta_Vendas_Dia": round(meta_dia, 1),
            "Meta_Vendas_Mes": meta_m,
            "Pct_Meta_Dia": round(pct_dia, 1),
            "Pct_Meta_Mes": round(pct_mes, 1),
            "Ineficiencia": res_pc.ineficiencia_qtd,
            "Vendas_Futuras": ven_fut,
            "Pct_Venda_Futura": round(ven_fut / ven_qtd * 100.0, 1) if ven_qtd > 0 else 0.0,
        })
    return pd.DataFrame(rows)


def render_tabela_radar_emp(df: pd.DataFrame) -> None:
    st.subheader("Radar por empreendimento")
    st.caption("Colunas alinhadas ao Polaroid: estoque, diff. avaliação, metas dia/mês, ineficiência, venda futura.")
    if df.empty:
        st.info("Sem empreendimentos para a tabela Radar.")
        return
    exibir_tabela(df)


CANAIS_STACK = ["RIO", "DIR", "GC", "PC"]


def _canal_stack(val: Any) -> str:
    p = _prefixo_imob(val)
    if p in CANAIS_DV:
        return "DIR"
    if p == "RJG":
        return "GC"
    if p == "RJ":
        return "PC"
    return "RIO"


def _prep_vendas_canal(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Canal_Stack" not in out.columns:
        col = "Imobiliária" if "Imobiliária" in out.columns else "Canal"
        if col in out.columns:
            out["Canal_Stack"] = out[col].map(_canal_stack)
        else:
            out["Canal_Stack"] = "RIO"
    return out


def _aplicar_filtros_funil_df(
    df: pd.DataFrame,
    filtros: FiltrosDashboard,
    mapa_coord: Dict[str, str],
    col_data: str,
) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    v = _v()
    out = df.copy()
    if filtros.emps_sel:
        col_emp = v.achar_coluna(out, v.ALIASES_EMPREENDIMENTO) or "Empreendimento"
        if col_emp in out.columns:
            sel = {_limpar_emp(e) for e in filtros.emps_sel}
            out = out[out[col_emp].map(_limpar_emp).isin(sel)]
    if filtros.coords_sel and mapa_coord:
        col_emp = v.achar_coluna(out, v.ALIASES_EMPREENDIMENTO) or "Empreendimento"
        if col_emp in out.columns:
            emps_c = {e for e, c in mapa_coord.items() if c in set(filtros.coords_sel)}
            out = out[out[col_emp].map(_limpar_emp).isin(emps_c)]
    if col_data and col_data in out.columns:
        out = _filtrar_df_periodo(out, col_data, filtros.data_ini, filtros.data_fim)
    return out


def calcular_pastas_sem_visita(
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    empreendimentos: Optional[List[str]] = None,
) -> Tuple[int, int, float]:
    """Pastas sem visita linkada (WhatId/AccountId) e percentual."""
    v = _v()
    if df_pastas is None or df_pastas.empty:
        return 0, 0, 0.0
    pas = v.deduplicar_pastas_funil(df_pastas)
    linked: set = set()
    if df_ag is not None and not df_ag.empty:
        ag = v.deduplicar_agendamentos_funil(df_ag)
        col_vis = v.achar_coluna(ag, v.ALIASES_DATA_VISITA) or "Data da visita"
        col_emp = v.achar_coluna(ag, v.ALIASES_EMPREENDIMENTO) or "Empreendimento"
        ag = ag.copy()
        ag["_vis"] = _parse_dt_series(ag, col_vis)
        ag = ag.dropna(subset=["_vis"])
        ag["_emp"] = ag[col_emp].map(_limpar_emp) if col_emp in ag.columns else ""
        for _, row in ag.iterrows():
            emp = row["_emp"]
            opp = _opp_id(row.get("WhatId"))
            conta = str(row.get("AccountId") or "").strip()
            if opp:
                linked.add(("opp", emp, opp))
            if conta and conta.lower() not in ("nan", "none", ""):
                linked.add(("conta", emp, conta))
    col_emp_p = v.achar_coluna(pas, v.ALIASES_EMPREENDIMENTO) or "Empreendimento"
    col_opp = v.achar_coluna(pas, ALIASES_OPP_AVAL) or "Oportunidade"
    col_conta = v.achar_coluna(pas, ["Conta__c", "Conta"]) or "Conta"
    if empreendimentos:
        sel = {_limpar_emp(e) for e in empreendimentos}
        pas = pas[pas[col_emp_p].map(_limpar_emp).isin(sel)]
    total = len(pas)
    sem = 0
    for _, row in pas.iterrows():
        emp = _limpar_emp(row.get(col_emp_p))
        opp = str(row.get(col_opp) or "").strip()
        conta = str(row.get(col_conta) or "").strip()
        ok = (opp and ("opp", emp, opp) in linked) or (
            conta and conta.lower() not in ("nan", "none", "")
            and ("conta", emp, conta) in linked
        )
        if not ok:
            sem += 1
    pct = (sem / total * 100.0) if total > 0 else 0.0
    return sem, total, pct


def _idx_pareto_corte(cum: np.ndarray, alvo: float) -> int:
    for i, v in enumerate(cum):
        if v >= alvo:
            return i
    return max(len(cum) - 1, 0)


def _pareto_linha_corte(fig: go.Figure, x: float, cor: str, rotulo: str) -> None:
    """Linha vertical de corte ABC — evita add_vline (incompatível com subplots/categóricos)."""
    fig.add_shape(
        type="line",
        x0=x,
        x1=x,
        y0=0,
        y1=1,
        xref="x",
        yref="paper",
        line=dict(dash="dash", color=cor, width=1.5),
        layer="below",
    )
    fig.add_annotation(
        x=x,
        y=1.02,
        xref="x",
        yref="paper",
        text=rotulo,
        showarrow=False,
        font=dict(color=cor, size=11),
        xanchor="center",
    )


def _plot_pareto_abc(
    df: pd.DataFrame,
    dim_col: str,
    titulo: str,
    key_prefix: str,
    top_n: int = 25,
) -> None:
    """Gráfico ABC: barras empilhadas por canal + linha % acumulada + cortes 75%/95%."""
    if df.empty or dim_col not in df.columns:
        st.info(f"Sem dados para {titulo}.")
        return
    prep = _prep_vendas_canal(df)
    qcol = "_qtd_venda" if "_qtd_venda" in prep.columns else None
    if not qcol:
        prep["_q"] = 1.0
        qcol = "_q"
    agg = prep.groupby([dim_col, "Canal_Stack"], as_index=False)[qcol].sum()
    pivot = agg.pivot(index=dim_col, columns="Canal_Stack", values=qcol).fillna(0)
    for c in CANAIS_STACK:
        if c not in pivot.columns:
            pivot[c] = 0.0
    pivot = pivot[CANAIS_STACK]
    pivot["_tot"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("_tot", ascending=False).head(top_n)
    if pivot.empty or pivot["_tot"].sum() <= 0:
        st.info(f"Sem volume para {titulo}.")
        return
    labels = list(pivot.index.astype(str))
    x_idx = list(range(len(labels)))
    totals = pivot["_tot"].values.astype(float)
    cum_pct = np.cumsum(totals) / totals.sum() * 100.0
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    cores = {"RIO": "#2563eb", "DIR": "#16a34a", "GC": "#f59e0b", "PC": "#8b5cf6"}
    for canal in CANAIS_STACK:
        fig.add_trace(
            go.Bar(x=x_idx, y=pivot[canal].values, name=canal, marker_color=cores.get(canal, "#64748b")),
            secondary_y=False,
        )
    fig.add_trace(
        go.Scatter(
            x=x_idx, y=cum_pct, name="% acumulada", mode="lines+markers",
            line=dict(color="#dc2626", width=2), marker=dict(size=6),
        ),
        secondary_y=True,
    )
    for alvo, cor in ((75.0, "#f59e0b"), (95.0, "#ef4444")):
        idx = _idx_pareto_corte(cum_pct, alvo)
        if idx < len(labels):
            _pareto_linha_corte(fig, float(idx), cor, f"{alvo:.0f}%")
    fig.update_layout(
        title=titulo, barmode="stack", height=460,
        margin=dict(l=20, r=20, t=50, b=80),
        legend=dict(orientation="h", y=-0.25),
    )
    fig.update_xaxes(tickmode="array", tickvals=x_idx, ticktext=labels)
    fig.update_yaxes(title_text="Quantidade", secondary_y=False)
    fig.update_yaxes(title_text="% acumulada", range=[0, 105], secondary_y=True)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False}, key=_plotly_key(key_prefix))


def _periodos_ps_sinais(hoje: Optional[date] = None) -> Dict[str, Tuple[date, date]]:
    hoje = hoje or date.today()
    ini_mtd = date(hoje.year, hoje.month, 1)
    if hoje.month == 1:
        ini_m1 = date(hoje.year - 1, 12, 1)
        fim_m1 = date(hoje.year - 1, 12, 31)
    else:
        ini_m1 = date(hoje.year, hoje.month - 1, 1)
        fim_m1 = date(hoje.year, hoje.month - 1, calendar.monthrange(hoje.year, hoje.month - 1)[1])
    ini_12m = hoje - timedelta(days=365)
    return {
        "MTD": (ini_mtd, hoje),
        "Mes_Anterior": (ini_m1, fim_m1),
        "Ultimos_12m": (ini_12m, hoje),
    }


def _metricas_ps_sinais_periodo(sub: pd.DataFrame) -> Tuple[float, float]:
    if sub.empty:
        return 0.0, 0.0
    if "_vgv_venda" in sub.columns:
        vgv = _sum_col_num(sub, "_vgv_venda", 0.0)
    elif "_vgv" in sub.columns:
        vgv = _sum_col_num(sub, "_vgv", 0.0)
    else:
        vgv = 0.0
    if vgv <= 0:
        return 0.0, 0.0
    if "PS_VGV" in sub.columns and "_vgv_venda" in sub.columns:
        ps = pd.to_numeric(sub["PS_VGV"], errors="coerce").fillna(0.0)
        vgv_col = pd.to_numeric(sub["_vgv_venda"], errors="coerce").fillna(0.0)
        ps_num = float((ps * vgv_col).sum())
    else:
        ps_num = 0.0
    sinal = float(sub["Total_Sinal"].sum()) if "Total_Sinal" in sub.columns else 0.0
    return ps_num / vgv, sinal / vgv


def montar_tabela_ps_sinais_vgv(
    df_vendas: pd.DataFrame,
    col_data: str,
    filtros: FiltrosDashboard,
    mapa_coord: Dict[str, str],
    hoje: Optional[date] = None,
) -> pd.DataFrame:
    """PS/VGV e Sinais/VGV por empreendimento — MTD, mês anterior, 12m (vs média 12m)."""
    if not col_data or df_vendas.empty or "Empreendimento" not in df_vendas.columns:
        return pd.DataFrame()
    base = _aplicar_filtros_base(df_vendas, filtros, mapa_coord, col_data, usar_periodo=False)
    base = _prep_vendas_canal(base)
    hoje = hoje or date.today()
    periodos = _periodos_ps_sinais(hoje)
    emps = sorted(base["Empreendimento"].map(_limpar_emp).dropna().unique())
    if filtros.emps_sel:
        emps = [_limpar_emp(e) for e in filtros.emps_sel if _limpar_emp(e) in emps]
    rows = []
    for emp in emps:
        sub_emp = base[base["Empreendimento"].map(_limpar_emp) == emp]
        row: Dict[str, Any] = {"Empreendimento": emp}
        vals_ps: List[float] = []
        vals_si: List[float] = []
        for chave, (ini, fim) in periodos.items():
            sub = _filtrar_df_periodo(sub_emp, col_data, ini, fim)
            ps, si = _metricas_ps_sinais_periodo(sub)
            row[f"PS_VGV_{chave}"] = round(ps * 100, 2)
            row[f"Sinais_VGV_{chave}"] = round(si * 100, 2)
            if chave == "Ultimos_12m":
                vals_ps.append(ps)
                vals_si.append(si)
        media_ps = float(np.mean(vals_ps)) if vals_ps else 0.0
        media_si = float(np.mean(vals_si)) if vals_si else 0.0
        row["PS_VGV_Media_12m"] = round(media_ps * 100, 2)
        row["Sinais_VGV_Media_12m"] = round(media_si * 100, 2)
        mtd_ps = row.get("PS_VGV_MTD", 0.0)
        mtd_si = row.get("Sinais_VGV_MTD", 0.0)
        row["PS_VGV_MTD_vs_12m"] = round(mtd_ps - row["PS_VGV_Media_12m"], 2)
        row["Sinais_VGV_MTD_vs_12m"] = round(mtd_si - row["Sinais_VGV_Media_12m"], 2)
        rows.append(row)
    return pd.DataFrame(rows)


def render_grafico_vendas_mes_canal(
    df_vendas: pd.DataFrame,
    col_data: str,
    key_prefix: str = "vendas_mes",
) -> None:
    st.subheader("Quantidade de vendas por mês (empilhado por canal)")
    if not col_data or df_vendas.empty:
        st.info("Sem dados de vendas por mês.")
        return
    v = _v()
    prep = _prep_vendas_canal(df_vendas)
    prep["_dt"] = v.parse_data_serie(prep[col_data])
    prep = prep.dropna(subset=["_dt"])
    if prep.empty:
        st.info("Sem datas válidas.")
        return
    prep["_mes"] = prep["_dt"].dt.to_period("M").astype(str)
    qcol = "_qtd_venda" if "_qtd_venda" in prep.columns else None
    if not qcol:
        prep["_q"] = 1.0
        qcol = "_q"
    agg = prep.groupby(["_mes", "Canal_Stack"], as_index=False)[qcol].sum()
    pivot = agg.pivot(index="_mes", columns="Canal_Stack", values=qcol).fillna(0).sort_index()
    for c in CANAIS_STACK:
        if c not in pivot.columns:
            pivot[c] = 0.0
    fig = go.Figure()
    cores = {"RIO": "#2563eb", "DIR": "#16a34a", "GC": "#f59e0b", "PC": "#8b5cf6"}
    for canal in CANAIS_STACK:
        fig.add_trace(go.Bar(x=pivot.index, y=pivot[canal], name=canal, marker_color=cores.get(canal)))
    fig.update_layout(barmode="stack", height=420, margin=dict(l=20, r=20, t=40, b=20))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False}, key=_plotly_key(key_prefix, "mes"))


def render_grafico_vendas_emp_canal(
    df_vendas: pd.DataFrame,
    key_prefix: str = "vendas_emp",
    top_n: int = 20,
) -> None:
    st.subheader("Quantidade de vendas por empreendimento (empilhado por canal)")
    if df_vendas.empty or "Empreendimento" not in df_vendas.columns:
        st.info("Sem dados por empreendimento.")
        return
    prep = _prep_vendas_canal(df_vendas)
    qcol = "_qtd_venda" if "_qtd_venda" in prep.columns else None
    if not qcol:
        prep["_q"] = 1.0
        qcol = "_q"
    agg = prep.groupby(["Empreendimento", "Canal_Stack"], as_index=False)[qcol].sum()
    tot = agg.groupby("Empreendimento")[qcol].sum().sort_values(ascending=False).head(top_n)
    agg = agg[agg["Empreendimento"].isin(tot.index)]
    pivot = agg.pivot(index="Empreendimento", columns="Canal_Stack", values=qcol).fillna(0)
    for c in CANAIS_STACK:
        if c not in pivot.columns:
            pivot[c] = 0.0
    pivot = pivot.loc[tot.index]
    fig = go.Figure()
    cores = {"RIO": "#2563eb", "DIR": "#16a34a", "GC": "#f59e0b", "PC": "#8b5cf6"}
    for canal in CANAIS_STACK:
        fig.add_trace(go.Bar(x=pivot.index, y=pivot[canal], name=canal, marker_color=cores.get(canal)))
    fig.update_layout(barmode="stack", height=460, margin=dict(l=20, r=20, t=40, b=100))
    fig.update_xaxes(tickangle=-35)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False}, key=_plotly_key(key_prefix, "emp"))


def render_grafico_share_estoque_produto(
    df_estoque: pd.DataFrame,
    status_sel: List[str],
    filtros: FiltrosDashboard,
    mapa_coord: Dict[str, str],
) -> None:
    st.subheader("Representação do produto no estoque")
    st.caption("Share = unidades do empreendimento (status filtrado) ÷ total filtrado.")
    if df_estoque is None or df_estoque.empty:
        st.info("Estoque indisponível.")
        return
    df = df_estoque.copy()
    col_st = "StatusUnidade__c" if "StatusUnidade__c" in df.columns else None
    if col_st and status_sel:
        df = df[df[col_st].astype(str).str.strip().isin(status_sel)]
    if filtros.emps_sel and "Empreendimento" in df.columns:
        sel = {_limpar_emp(e) for e in filtros.emps_sel}
        df = df[df["Empreendimento"].map(_limpar_emp).isin(sel)]
    if filtros.coords_sel and mapa_coord and "Empreendimento" in df.columns:
        emps_c = {e for e, c in mapa_coord.items() if c in set(filtros.coords_sel)}
        df = df[df["Empreendimento"].map(_limpar_emp).isin(emps_c)]
    if df.empty:
        st.info("Nenhuma unidade após filtros de status/empreendimento.")
        return
    cont = df.groupby(df["Empreendimento"].map(_limpar_emp)).size().sort_values(ascending=False)
    total = float(cont.sum())
    pct = (cont / total * 100.0).round(1)
    fig = go.Figure()
    fig.add_trace(go.Bar(x=cont.index.astype(str), y=cont.values, name="Unidades", marker_color="#2563eb"))
    fig.add_trace(go.Scatter(
        x=pct.index.astype(str), y=pct.values, name="Share (%)", mode="lines+markers",
        yaxis="y2", line=dict(color="#dc2626"),
    ))
    fig.update_layout(
        height=440, margin=dict(l=20, r=20, t=40, b=100),
        yaxis=dict(title="Unidades"),
        yaxis2=dict(title="Share (%)", overlaying="y", side="right", range=[0, max(pct.max() * 1.1, 5)]),
    )
    fig.update_xaxes(tickangle=-35)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False}, key=_plotly_key("share_estoque"))


def render_secao_analises_avancadas(
    df_vendas: pd.DataFrame,
    df_estoque: pd.DataFrame,
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    filtros: FiltrosDashboard,
    mapa_coord: Dict[str, str],
    col_data: str,
    status_estoque: List[str],
    df_cotacoes: Optional[pd.DataFrame] = None,
) -> None:
    """KPIs e gráficos: pastas sem visita, Pareto, vendas/mês, PS/VGV, estoque."""
    st.markdown("---")
    st.subheader("Análises avançadas")
    st.caption("Indicadores comparados com média dos últimos 12 meses quando aplicável.")

    df_ag_f = _aplicar_filtros_funil_df(df_ag, filtros, mapa_coord, "Data de criação")
    df_pas_f = _aplicar_filtros_funil_df(
        df_pastas, filtros, mapa_coord,
        _v().achar_coluna(df_pastas, COLUNAS_PASTAS_ALIASES) or "Data Primeiro Envio Análise",
    )
    emps = filtros.emps_sel or None
    sem, tot, pct = calcular_pastas_sem_visita(df_ag_f, df_pas_f, emps)
    st.markdown("##### Pastas sem visita linkada")
    c1, c2, c3 = st.columns(3)
    c1.metric("Pastas sem visita", sem)
    c2.metric("Total pastas", tot)
    c3.metric("% sem visita", fmt_pct_valor(pct))

    df_ctx = _aplicar_filtros_base(df_vendas, filtros, mapa_coord, col_data, usar_periodo=True)
    df_ctx = _prep_vendas_canal(df_ctx)

    st.markdown("##### Curvas de Pareto (ABC)")
    p1, p2 = st.columns(2)
    with p1:
        _plot_pareto_abc(df_ctx, "Empreendimento", "Pareto por empreendimento", "pareto_emp")
    with p2:
        col_reg = "Gerente regional" if "Gerente regional" in df_ctx.columns else "Região"
        if col_reg in df_ctx.columns:
            _plot_pareto_abc(df_ctx, col_reg, "Pareto por regional", "pareto_reg")
        else:
            st.info("Coluna regional não disponível.")
    if "Imobiliária" in df_ctx.columns:
        _plot_pareto_abc(df_ctx, "Imobiliária", "Pareto por imobiliária", "pareto_imob", top_n=30)

    g1, g2 = st.columns(2)
    with g1:
        render_grafico_vendas_mes_canal(df_ctx, col_data)
    with g2:
        render_grafico_vendas_emp_canal(df_ctx)

    render_grafico_share_estoque_produto(df_estoque, status_estoque, filtros, mapa_coord)


def render_dashboard_comercial(
    df_vendas: pd.DataFrame,
    df_estoque: pd.DataFrame,
    df_cotacoes: Optional[pd.DataFrame],
    cred_fp: str,
    col_data_venda: Optional[str] = None,
    filtros_glob: Optional["FiltrosGlobais"] = None,
    df_metas_fallback: Optional[pd.DataFrame] = None,
) -> None:
    """Ponto de entrada do dashboard comercial."""
    v = _v()
    ano_fb = filtros_glob.ano_meta if filtros_glob else date.today().year
    mes_fb = filtros_glob.mes_meta if filtros_glob else date.today().month
    df_metas_coord, aviso_coord = carregar_metas_coordenadores_com_fallback(
        cred_fp, df_metas_fallback, ano_fb, mes_fb,
    )
    try:
        df_canal = carregar_metas_canal(cred_fp)
    except Exception as exc:
        st.warning(f"Metas canal indisponíveis (aba Canal): {exc}")
        df_canal = pd.DataFrame()
    if df_metas_coord.empty:
        st.error(f"Erro ao carregar metas: {aviso_coord or 'sem dados'}")
        return
    if aviso_coord:
        st.warning(f"Metas coordenadores: usando planilha legado ({aviso_coord}).")

    df_vendas = enriquecer_vendas_vcx(df_vendas, df_cotacoes)
    col_data = col_data_venda or _col_data_venda(df_vendas)

    filtros = render_filtros_dashboard(df_metas_coord, df_vendas, filtros_externos=filtros_glob)
    if filtros_glob is not None:
        st.caption(
            f"Filtros globais: {filtros.data_ini:%d/%m/%Y} → {filtros.data_fim:%d/%m/%Y} · "
            f"Canal {filtros.canal_sel} · Meta {filtros.tipo_meta_col}"
        )
    mapa_coord = mapa_emp_coordenador(df_metas_coord, filtros.mes_meta, filtros.ano_meta)

    df_f = _aplicar_filtros_base(df_vendas, filtros, mapa_coord, col_data, usar_periodo=True)
    df_ctx = _aplicar_filtros_base(df_vendas, filtros, mapa_coord, col_data, usar_periodo=False)

    df_ag_radar = pd.DataFrame()
    df_pastas_radar = pd.DataFrame()
    df_pastas_aprov_radar = pd.DataFrame()
    df_tabela_radar = pd.DataFrame()
    try:
        pacote_funil_r = carregar_funil_painel_sf()
        df_ag_radar = _coalesce_dict_df(pacote_funil_r, "agendamentos")
        df_pastas_radar = _coalesce_dict_df(pacote_funil_r, "pastas")
        if not df_pastas_radar.empty:
            df_pastas_aprov_radar = deduplicar_pastas_aprovadas_funil(df_pastas_radar)
        pacote_pc_r = carregar_pacote_poder_compra_sf()
        df_tabela_radar = _coalesce_dict_df(pacote_pc_r, "tabela_comprometimento")
    except Exception:
        pass

    render_radar_polaroid(
        df_ctx, df_estoque, df_canal, df_metas_coord, filtros, mapa_coord, col_data,
        df_ag_radar, df_pastas_radar,
    )

    render_velocimetros_vgv(df_f, df_canal, filtros, col_data)
    render_velocimetros_coordenador_vgv(
        df_f, df_metas_coord, df_canal, filtros, mapa_coord, col_data,
    )

    c1, c2 = st.columns(2)
    with c1:
        render_grafico_share_canal(df_f)
    with c2:
        render_grafico_evolucao_vgv(df_ctx, df_canal, filtros, col_data)

    render_box_gap(df_ctx, df_canal, filtros, col_data)

    render_rankings(df_f, col_data, filtros)
    render_share_diretor(df_f)
    render_tabela_detalhada(df_f, col_data, filtros)

    status_est = (
        filtros_glob.status_estoque_sel
        if filtros_glob and filtros_glob.status_estoque_sel
        else list(ESTOQUE_STATUS_TODOS)
    )
    render_secao_analises_avancadas(
        df_vendas, df_estoque, df_ag_radar, df_pastas_radar,
        filtros, mapa_coord, col_data, status_est,
        df_cotacoes=df_cotacoes,
    )


# =============================================================================
# FUNIL TEMPOS (inline — ex-velocimetro_funil_tempos.py)
# =============================================================================

# Tempos médios de conversão do funil e hipereficiência.

from dataclasses import dataclass
from datetime import date
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st



@dataclass
class TemposEmpreendimento:
    empreendimento: str
    media_agend_visita: Optional[float] = None
    n_agend_visita: int = 0
    media_visita_pasta: Optional[float] = None
    n_visita_pasta: int = 0
    media_pasta_aprov: Optional[float] = None
    n_pasta_aprov: int = 0
    media_aprov_venda: Optional[float] = None
    n_aprov_venda: int = 0


@dataclass
class HipereficienciaEmp:
    empreendimento: str
    vendas: int = 0
    hipereficiencia_qtd: int = 0
    hipereficiencia_pct: float = 0.0






def _parse_dt_series(df: pd.DataFrame, col: str) -> pd.Series:
    v = _v()
    if col not in df.columns:
        return pd.Series(dtype="datetime64[ns]")
    return v.parse_data_serie(df[col])


def _dias_entre(dt_ini: Any, dt_fim: Any) -> Optional[float]:
    if dt_ini is None or dt_fim is None or pd.isna(dt_ini) or pd.isna(dt_fim):
        return None
    try:
        delta = (pd.Timestamp(dt_fim) - pd.Timestamp(dt_ini)).total_seconds() / 86400.0
    except Exception:
        return None
    return float(delta) if delta >= 0 else None


def _media(lst: List[float]) -> Optional[float]:
    if not lst:
        return None
    return float(np.mean(lst))


def _opp_id(val: Any) -> str:
    s = str(val or "").strip()
    return s if s.startswith("006") else ""


def _indexar_pastas(
    df_pastas: pd.DataFrame,
) -> Tuple[Dict[Tuple[str, str], List[Tuple[Any, Any]]], Dict[Tuple[str, str], List[Tuple[Any, Any, str]]]]:
    """Índices (emp, opp) e (emp, conta) → [(dt_criacao, dt_aprov, opp)]."""
    por_opp: Dict[Tuple[str, str], List[Tuple[Any, Any]]] = {}
    por_conta: Dict[Tuple[str, str], List[Tuple[Any, Any, str]]] = {}
    if df_pastas is None or df_pastas.empty:
        return por_opp, por_conta
    v = _v()
    col_emp = v.achar_coluna(df_pastas, v.ALIASES_EMPREENDIMENTO) or "Empreendimento"
    col_cri = v.achar_coluna(df_pastas, v.ALIASES_DATA_CRIACAO) or "Data de criação"
    col_apr = v.achar_coluna_aprovacao_safi(df_pastas)
    col_opp = v.achar_coluna(df_pastas, ALIASES_OPP_AVAL) or "Oportunidade"
    col_conta = v.achar_coluna(df_pastas, ["Conta__c", "Conta"]) or "Conta"
    df = df_pastas.copy()
    df["_cri"] = _parse_dt_series(df, col_cri)
    df["_apr"] = _parse_dt_series(df, col_apr) if col_apr else pd.NaT
    for _, row in df.iterrows():
        emp = _limpar_emp(row.get(col_emp))
        if not emp:
            continue
        opp = str(row.get(col_opp) or "").strip()
        conta = str(row.get(col_conta) or "").strip()
        cri, apr = row.get("_cri"), row.get("_apr")
        if opp:
            por_opp.setdefault((emp, opp), []).append((cri, apr))
        if conta and conta.lower() not in ("nan", "none", ""):
            por_conta.setdefault((emp, conta), []).append((cri, apr, opp))
    return por_opp, por_conta


def _menor_pasta_apos(
    candidatos: List[Tuple[Any, Any]],
    ref: Any,
) -> Optional[Any]:
    """Data de criação da pasta mais próxima após ref."""
    best = None
    best_d = None
    for cri, _ in candidatos:
        d = _dias_entre(ref, cri)
        if d is None:
            continue
        if best_d is None or d < best_d:
            best_d = d
            best = cri
    return best


def calcular_tempos_agend_visita(
    df_ag: pd.DataFrame,
    empreendimentos: Optional[List[str]] = None,
) -> Dict[str, Tuple[Optional[float], int]]:
    """Média dias: data de criação do agendamento → data da visita."""
    v = _v()
    out: Dict[str, Tuple[Optional[float], int]] = {}
    if df_ag is None or df_ag.empty:
        return out
    df = v.deduplicar_agendamentos_funil(df_ag)
    col_emp = v.achar_coluna(df, v.ALIASES_EMPREENDIMENTO) or "Empreendimento"
    col_cri = v.achar_coluna(df, v.ALIASES_DATA_CRIACAO) or "Data de criação"
    col_vis = v.achar_coluna(df, v.ALIASES_DATA_VISITA) or "Data da visita"
    if col_cri not in df.columns or col_vis not in df.columns:
        return out
    df = df.copy()
    df["_cri"] = _parse_dt_series(df, col_cri)
    df["_vis"] = _parse_dt_series(df, col_vis)
    if col_emp in df.columns:
        df["_emp"] = df[col_emp].map(_limpar_emp)
    else:
        df["_emp"] = ""
    if empreendimentos:
        sel = {_limpar_emp(e) for e in empreendimentos}
        df = df[df["_emp"].isin(sel)]
    por_emp: Dict[str, List[float]] = {}
    for _, row in df.iterrows():
        d = _dias_entre(row["_cri"], row["_vis"])
        if d is None:
            continue
        emp = row["_emp"] or "(sem empreendimento)"
        por_emp.setdefault(emp, []).append(d)
    for emp, vals in por_emp.items():
        out[emp] = (_media(vals), len(vals))
    return out


def calcular_tempos_visita_pasta(
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    empreendimentos: Optional[List[str]] = None,
) -> Dict[str, Tuple[Optional[float], int]]:
    """Média dias: data da visita → data de criação da pasta."""
    v = _v()
    out: Dict[str, Tuple[Optional[float], int]] = {}
    if df_ag is None or df_ag.empty or df_pastas is None or df_pastas.empty:
        return out
    ag = v.deduplicar_agendamentos_funil(df_ag)
    pas = v.deduplicar_pastas_funil(df_pastas)
    por_opp, por_conta = _indexar_pastas(pas)
    col_emp = v.achar_coluna(ag, v.ALIASES_EMPREENDIMENTO) or "Empreendimento"
    col_vis = v.achar_coluna(ag, v.ALIASES_DATA_VISITA) or "Data da visita"
    ag = ag.copy()
    ag["_vis"] = _parse_dt_series(ag, col_vis)
    if col_emp in ag.columns:
        ag["_emp"] = ag[col_emp].map(_limpar_emp)
    else:
        ag["_emp"] = ""
    if empreendimentos:
        sel = {_limpar_emp(e) for e in empreendimentos}
        ag = ag[ag["_emp"].isin(sel)]
    por_emp: Dict[str, List[float]] = {}
    for _, row in ag.iterrows():
        vis = row.get("_vis")
        if pd.isna(vis):
            continue
        emp = row["_emp"] or "(sem empreendimento)"
        opp = _opp_id(row.get("WhatId"))
        conta = str(row.get("AccountId") or "").strip()
        dt_pasta = None
        if opp:
            dt_pasta = _menor_pasta_apos(por_opp.get((emp, opp), []), vis)
        if dt_pasta is None and conta:
            dt_pasta = _menor_pasta_apos(
                [(c, a) for c, a, _ in por_conta.get((emp, conta), [])],
                vis,
            )
        d = _dias_entre(vis, dt_pasta)
        if d is None:
            continue
        por_emp.setdefault(emp, []).append(d)
    for emp, vals in por_emp.items():
        out[emp] = (_media(vals), len(vals))
    return out


def calcular_tempos_pasta_aprovada(
    df_pastas: pd.DataFrame,
    empreendimentos: Optional[List[str]] = None,
) -> Dict[str, Tuple[Optional[float], int]]:
    """Média dias: data de criação da pasta → data de aprovação SAFI."""
    v = _v()
    out: Dict[str, Tuple[Optional[float], int]] = {}
    if df_pastas is None or df_pastas.empty:
        return out
    df = v.deduplicar_pastas_funil(df_pastas)
    col_emp = v.achar_coluna(df, v.ALIASES_EMPREENDIMENTO) or "Empreendimento"
    col_cri = v.achar_coluna(df, v.ALIASES_DATA_CRIACAO) or "Data de criação"
    col_apr = v.achar_coluna_aprovacao_safi(df)
    if not col_apr or col_cri not in df.columns:
        return out
    df = df.copy()
    df["_cri"] = _parse_dt_series(df, col_cri)
    df["_apr"] = _parse_dt_series(df, col_apr)
    df = df.dropna(subset=["_cri", "_apr"])
    if col_emp in df.columns:
        df["_emp"] = df[col_emp].map(_limpar_emp)
    else:
        df["_emp"] = ""
    if empreendimentos:
        sel = {_limpar_emp(e) for e in empreendimentos}
        df = df[df["_emp"].isin(sel)]
    por_emp: Dict[str, List[float]] = {}
    for _, row in df.iterrows():
        d = _dias_entre(row["_cri"], row["_apr"])
        if d is None:
            continue
        emp = row["_emp"] or "(sem empreendimento)"
        por_emp.setdefault(emp, []).append(d)
    for emp, vals in por_emp.items():
        out[emp] = (_media(vals), len(vals))
    return out


def calcular_tempos_aprov_venda(
    df_pastas: pd.DataFrame,
    df_vendas: pd.DataFrame,
    empreendimentos: Optional[List[str]] = None,
) -> Dict[str, Tuple[Optional[float], int]]:
    """Média dias: data de aprovação SAFI → contrato gerado em (vendas)."""
    v = _v()
    out: Dict[str, Tuple[Optional[float], int]] = {}
    if df_pastas is None or df_pastas.empty or df_vendas is None or df_vendas.empty:
        return out
    pas_ap = v.deduplicar_pastas_aprovadas_funil(df_pastas)
    ven = v.deduplicar_vendas_funil(v.filtrar_vendas_comerciais(df_vendas))
    col_apr = v.achar_coluna_aprovacao_safi(pas_ap)
    col_opp_p = v.achar_coluna(pas_ap, ALIASES_OPP_AVAL) or "Oportunidade"
    col_opp_v = v.achar_coluna(ven, v.ALIASES_ID_OPORTUNIDADE) or "ID da Oportunidade"
    col_con = v.achar_coluna(ven, v.ALIASES_CONTRATO_GERADO) or "Contrato gerado em"
    col_emp_v = v.achar_coluna(ven, v.ALIASES_EMPREENDIMENTO) or "Empreendimento"
    if not col_apr or col_con not in ven.columns:
        return out
    pas_ap = pas_ap.copy()
    pas_ap["_apr"] = _parse_dt_series(pas_ap, col_apr)
    apr_por_opp: Dict[str, Any] = {}
    for _, row in pas_ap.iterrows():
        oid = str(row.get(col_opp_p) or "").strip()
        if oid and oid not in apr_por_opp:
            apr_por_opp[oid] = row["_apr"]
    ven = ven.copy()
    ven["_contrato"] = _parse_dt_series(ven, col_con)
    if col_emp_v in ven.columns:
        ven["_emp"] = ven[col_emp_v].map(_limpar_emp)
    else:
        ven["_emp"] = ""
    if empreendimentos:
        sel = {_limpar_emp(e) for e in empreendimentos}
        ven = ven[ven["_emp"].isin(sel)]
    por_emp: Dict[str, List[float]] = {}
    for _, row in ven.iterrows():
        oid = str(row.get(col_opp_v) or "").strip()
        if not oid or oid not in apr_por_opp:
            continue
        d = _dias_entre(apr_por_opp[oid], row["_contrato"])
        if d is None:
            continue
        emp = row["_emp"] or "(sem empreendimento)"
        por_emp.setdefault(emp, []).append(d)
    for emp, vals in por_emp.items():
        out[emp] = (_media(vals), len(vals))
    return out


def montar_tabela_tempos_conversao(
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    df_vendas: pd.DataFrame,
    empreendimentos: Optional[List[str]] = None,
) -> pd.DataFrame:
    t1 = calcular_tempos_agend_visita(df_ag, empreendimentos)
    t2 = calcular_tempos_visita_pasta(df_ag, df_pastas, empreendimentos)
    t3 = calcular_tempos_pasta_aprovada(df_pastas, empreendimentos)
    t4 = calcular_tempos_aprov_venda(df_pastas, df_vendas, empreendimentos)
    emps = set(t1) | set(t2) | set(t3) | set(t4)
    if empreendimentos:
        emps |= {_limpar_emp(e) for e in empreendimentos}
    rows = []
    for emp in sorted(emps):
        m1, n1 = t1.get(emp, (None, 0))
        m2, n2 = t2.get(emp, (None, 0))
        m3, n3 = t3.get(emp, (None, 0))
        m4, n4 = t4.get(emp, (None, 0))
        rows.append({
            "Empreendimento": emp,
            "Dias_Agend_Visita": round(m1, 1) if m1 is not None else None,
            "N_Agend_Visita": n1,
            "Dias_Visita_Pasta": round(m2, 1) if m2 is not None else None,
            "N_Visita_Pasta": n2,
            "Dias_Pasta_Aprov": round(m3, 1) if m3 is not None else None,
            "N_Pasta_Aprov": n3,
            "Dias_Aprov_Venda": round(m4, 1) if m4 is not None else None,
            "N_Aprov_Venda": n4,
        })
    return pd.DataFrame(rows)


def calcular_hipereficiencia_por_emp(
    df_vendas: pd.DataFrame,
    df_pastas_aprov: pd.DataFrame,
    df_est_enr: pd.DataFrame,
    df_tabela: pd.DataFrame,
    empreendimentos: Optional[List[str]] = None,
) -> pd.DataFrame:
    """
    Hipereficiência: clientes sem poder de compra suficiente que compraram unidade.
    Usa pasta aprovada vinculada à oportunidade e preço real da venda como referência.
    """
    v = _v()
    if df_vendas is None or df_vendas.empty:
        return pd.DataFrame()
    ven = v.deduplicar_vendas_funil(v.filtrar_vendas_comerciais(df_vendas))
    col_emp = v.achar_coluna(ven, v.ALIASES_EMPREENDIMENTO) or "Empreendimento"
    col_id = v.achar_coluna(ven, v.ALIASES_ID_OPORTUNIDADE) or "ID da Oportunidade"
    col_val = v.achar_coluna(ven, ["Valor Real de Venda", "Valor Real", "Valor"])
    if not col_id:
        return pd.DataFrame()

    pas_ap = pd.DataFrame()
    if df_pastas_aprov is not None and not df_pastas_aprov.empty:
        pas_ap = v.deduplicar_pastas_aprovadas_funil(df_pastas_aprov)
    elif df_pastas_aprov is not None:
        pas_ap = pd.DataFrame()

    mapa_tab = mapa_tabela_por_oportunidade(df_tabela)
    pas_por_opp: Dict[str, pd.Series] = {}
    if not pas_ap.empty:
        col_opp_p = v.achar_coluna(pas_ap, ALIASES_OPP_AVAL) or "Oportunidade"
        col_apr = v.achar_coluna_aprovacao_safi(pas_ap)
        if col_opp_p and col_apr:
            tmp = pas_ap.copy()
            tmp["_apr"] = _parse_dt_series(tmp, col_apr)
            tmp = tmp.sort_values("_apr", ascending=False, na_position="last")
            for _, row in tmp.iterrows():
                oid = str(row.get(col_opp_p) or "").strip()
                if oid and oid not in pas_por_opp:
                    pas_por_opp[oid] = row

    if empreendimentos:
        sel = {_limpar_emp(e) for e in empreendimentos}
        ven = ven[ven[col_emp].map(_limpar_emp).isin(sel)]

    por_emp: Dict[str, Dict[str, int]] = {}
    for _, row in ven.iterrows():
        emp = _limpar_emp(row.get(col_emp))
        oid = str(row.get(col_id) or "").strip()
        preco = _num(row.get(col_val)) if col_val else 0.0
        if preco <= 0:
            preco = float(row.get("_vgv") or row.get("_vgv_venda") or 0.0)
        if not emp or not oid:
            continue
        por_emp.setdefault(emp, {"vendas": 0, "hiper": 0})
        por_emp[emp]["vendas"] += 1
        pasta_row = pas_por_opp.get(oid)
        if pasta_row is None:
            continue
        analise = analisar_pasta(
            pasta_row, emp, df_est_enr, mapa_tab, {oid}, preco_override=preco,
        )
        if not analise.pc_suficiente:
            por_emp[emp]["hiper"] += 1

    rows = []
    for emp, d in sorted(por_emp.items()):
        vendas = d["vendas"]
        hiper = d["hiper"]
        rows.append({
            "Empreendimento": emp,
            "Vendas": vendas,
            "Hipereficiencia_Qtd": hiper,
            "Hipereficiencia_Pct": round(hiper / vendas * 100.0, 1) if vendas > 0 else 0.0,
        })
    return pd.DataFrame(rows)


def render_secao_funil_tempos(
    df_ag: pd.DataFrame,
    df_pastas: pd.DataFrame,
    df_vendas: pd.DataFrame,
    df_estoque: pd.DataFrame,
    df_tabela_comp: pd.DataFrame,
    empreendimentos: Optional[List[str]] = None,
) -> None:

    st.subheader("Tempos médios de conversão por empreendimento")
    st.caption(
        "Agendamento→Visita (criação×visita) · Visita→Pasta (visita×criação pasta) · "
        "Pasta→Aprovada (criação×aprovação) · Aprovada→Venda (aprovação×contrato gerado)"
    )
    tab_tempos = montar_tabela_tempos_conversao(
        df_ag, df_pastas, df_vendas, empreendimentos,
    )
    if tab_tempos.empty:
        st.info("Sem dados de funil para calcular tempos de conversão.")
    else:
        exibir_tabela(tab_tempos)

    st.subheader("Hipereficiência")
    st.caption(
        "Clientes **sem poder de compra** suficiente (avaliação aprovada) que **compraram** unidade."
    )
    _, enr = agregar_estoque(df_estoque if df_estoque is not None else pd.DataFrame())
    pas_aprov = pd.DataFrame()
    if df_pastas is not None and not df_pastas.empty:
        pas_aprov = _v().deduplicar_pastas_aprovadas_funil(df_pastas)
    tab_hiper = calcular_hipereficiencia_por_emp(
        df_vendas, pas_aprov, enr, df_tabela_comp, empreendimentos,
    )
    if tab_hiper.empty:
        st.info("Sem vendas para calcular hipereficiência.")
    else:
        exibir_tabela(tab_hiper)
def _corpo_painel_metas(
    df_vendas: pd.DataFrame,
    df_metas: pd.DataFrame,
    df_vendas_painel: pd.DataFrame,
    df_vendas_raw: pd.DataFrame,
    origem_vendas_painel: str,
    col_contrato_gerado: Optional[str],
    col_canal: Optional[str],
    cred_fp: str,
    sid: str,
    filtros_glob: Optional["FiltrosGlobais"] = None,
) -> None:

    df_estoque = pd.DataFrame()
    df_ag_v2 = pd.DataFrame()
    df_pastas_v2 = pd.DataFrame()
    df_cotacoes = pd.DataFrame()
    try:
        df_estoque = carregar_estoque_painel_sf()
    except Exception as exc:
        st.warning(f"Estoque indisponível no Google Sheets: {exc}")
    try:
        pacote_funil_v2 = carregar_funil_painel_sf()
        df_ag_v2 = deduplicar_agendamentos_funil(_coalesce_df(pacote_funil_v2.get("agendamentos")))
        df_pastas_v2 = deduplicar_pastas_funil(_coalesce_df(pacote_funil_v2.get("pastas")))
    except Exception as exc:
        st.warning(f"Funil indisponível no Google Sheets: {exc}")
    try:
        df_cotacoes = carregar_cotacoes_painel_sf()
    except Exception:
        pass

    df_pastas_aprov = pd.DataFrame()
    df_tabela_comp = pd.DataFrame()
    try:
        pacote_pc = carregar_pacote_poder_compra_sf()
        df_pastas_aprov = _coalesce_dict_df(pacote_pc, "pastas_aprovadas")
        df_tabela_comp = _coalesce_dict_df(pacote_pc, "tabela_comprometimento")
    except Exception as exc:
        st.warning(f"Pacote poder de compra: {exc}")

    filtros_v2 = render_painel_metas_v2(
        df_vendas,
        df_vendas_painel,
        col_contrato_gerado,
        cred_fp,
        df_estoque=df_estoque,
        df_ag=df_ag_v2,
        df_pastas=df_pastas_v2,
        df_cotacoes=df_cotacoes,
        df_pastas_aprov=df_pastas_aprov,
        df_tabela_comp=df_tabela_comp,
        proj=None,
        filtros_glob=filtros_glob,
        df_metas_fallback=df_metas,
        col_canal=col_canal,
    )

    vendas_f = filtrar_vendas_painel_v2(
        df_vendas, filtros_v2, col_contrato_gerado or "", col_canal,
    )

    df_metas_coord, _ = carregar_metas_coordenadores_com_fallback(
        cred_fp, df_metas, filtros_v2.ano_meta, filtros_v2.mes_meta,
    )
    try:
        df_canal = carregar_metas_canal(cred_fp)
    except Exception:
        df_canal = pd.DataFrame()

    if filtros_v2.tipo_indicador == "vendas":
        total_meta_vgv, total_meta_qtd_canal = meta_canal_vgv_vendas(
            df_canal, filtros_v2.mes_meta, filtros_v2.ano_meta, filtros_v2.canal_meta,
        )
        total_meta_qtd = soma_meta_coord(
            df_metas_coord, filtros_v2.mes_meta, filtros_v2.ano_meta,
            "vendas", filtros_v2.tipo_meta_col, filtros_v2.emps_sel or None,
        )
        if total_meta_vgv <= 0:
            total_meta_vgv = soma_meta_vgv_coord(
                df_metas_coord, filtros_v2.mes_meta, filtros_v2.ano_meta,
                filtros_v2.tipo_meta_col, empreendimentos=filtros_v2.emps_sel or None,
            )
        if total_meta_qtd <= 0 and 0 < total_meta_qtd_canal <= 5_000:
            total_meta_qtd = total_meta_qtd_canal
    else:
        total_meta_vgv = 0.0
        total_meta_qtd = soma_meta_coord(
            df_metas_coord, filtros_v2.mes_meta, filtros_v2.ano_meta,
            filtros_v2.tipo_indicador, filtros_v2.tipo_meta_col,
            filtros_v2.emps_sel or None,
        )

    fator_meta = FATORES_CANAL.get((filtros_v2.canal_meta or "RIO").strip().upper(), 0.0)
    total_realizado_qtd = (
        _sum_col_num(vendas_f, "_qtd_venda", float(len(vendas_f)))
        if "_qtd_venda" in vendas_f.columns
        else float(len(vendas_f))
    )
    total_vgv_realizado = _sum_col_num(vendas_f, "_vgv_venda", 0.0)

    # -------------------------------------------------------------------------
    # FUNIL IDEAL E ENGENHARIA REVERSA (dois funis lado a lado)
    # -------------------------------------------------------------------------
    st.markdown("<br><hr style='border:none;border-top:1px solid #e2e8f0;margin:1rem 0;'/>", unsafe_allow_html=True)
    st.subheader("Engenharia Reversa")

    v_meta = math.floor(total_meta_qtd)
    pa_ideal = math.ceil(v_meta / 0.64) if v_meta > 0 else 0
    p_ideal = math.ceil(pa_ideal / 0.64) if pa_ideal > 0 else 0
    vi_ideal = math.ceil(p_ideal / 0.25) if p_ideal > 0 else 0
    a_ideal = math.ceil(vi_ideal / 0.50) if vi_ideal > 0 else 0

    meta_global_referencia = (total_meta_qtd / fator_meta) if fator_meta > 0 else 0
    meta_dvrj_ref = meta_global_referencia * 0.5
    vd_ideal = math.ceil(meta_dvrj_ref * 0.40)

    od_ideal = math.ceil(vd_ideal / 0.044) if vd_ideal > 0 else 0
    ld_ideal = math.ceil(od_ideal / 0.50) if od_ideal > 0 else 0

    corretores_pessimista = math.ceil(v_meta / 0.15) if v_meta > 0 else 0
    corretores_moderado = math.ceil(v_meta / 0.20) if v_meta > 0 else 0
    corretores_otimista = math.ceil(v_meta / 0.25) if v_meta > 0 else 0

    col_funil_ideal, col_funil_mkt = st.columns(2)
    with col_funil_ideal:
        st.markdown("##### Funil ideal (conversões comerciais)")
        fig_ideal = _criar_fig_funil(
            ['Agendamentos', 'Visitas', 'Pastas', 'Past. Aprov.', 'Vendas (Meta)'],
            [a_ideal, vi_ideal, p_ideal, pa_ideal, v_meta],
            cores=["#022654", "#04428f", "#1e60b3", "#cb0935", "#9e0828"],
            altura=350,
        )
        st.plotly_chart(fig_ideal, use_container_width=True, config={"displayModeBar": False})
    with col_funil_mkt:
        st.markdown("##### Funil de marketing digital")
        fig_mkt = _criar_fig_funil(
            ['Leads Digitais', 'Oport. Digitais', 'Vendas Dig. (40% DV RJ)'],
            [ld_ideal, od_ideal, vd_ideal],
            cores=["#022654", "#1e60b3", "#cb0935"],
            altura=350,
        )
        st.plotly_chart(fig_mkt, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<br/>", unsafe_allow_html=True)
    st.subheader("Cenários: Corretores Ativos (Necessários para bater a meta global)")
    st.markdown(
        f"""
        <div class="vel-kpi-row" style="justify-content: center; margin-top: 1rem;">
                <div class="vel-kpi" style="flex: 0 1 300px;"><div class="lbl">Pessimista (15% convert.)</div><div class="val">{corretores_pessimista}</div></div>
                <div class="vel-kpi" style="flex: 0 1 300px;"><div class="lbl">Moderado (20% convert.)</div><div class="val">{corretores_moderado}</div></div>
                <div class="vel-kpi" style="flex: 0 1 300px;"><div class="lbl">Otimista (25% convert.)</div><div class="val">{corretores_otimista}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # -------------------------------------------------------------------------
    # Projeção completa do funil (ElasticNet + lags) vs médias sazonais
    # -------------------------------------------------------------------------
    meta_qtd_proj = float(total_meta_qtd)

    if col_contrato_gerado:
        base_proj = filtrar_vendas_painel_v2(
            df_vendas_painel, filtros_v2, col_contrato_gerado, col_canal,
            aplicar_periodo=False,
        )

        try:
            hoje_ef = date.today()
            ini_ef, fim_ef = janela_treino_meses_exatos(hoje_ef)
            serie_ef = serie_diaria_contratos(base_proj, col_contrato_gerado)
            if not serie_ef.empty:
                treino_ef = calendario_diario(ini_ef, fim_ef, serie_ef)
                efeitos = estimar_efeitos_sazonais(treino_ef)
                if efeitos:
                    render_efeitos_sazonais(efeitos)

            df_ag_vis_cache: Optional[pd.DataFrame] = None
            df_pastas_cache: Optional[pd.DataFrame] = None
            origem_ag_cache = ""
            origem_pastas_cache = ""
            try:
                pacote_funil = carregar_funil_painel_sf()
                df_ag_vis_cache = pacote_funil.get("agendamentos")
                df_pastas_cache = pacote_funil.get("pastas")
                origem_ag_cache = pacote_funil.get("origem_ag", "")
                origem_pastas_cache = pacote_funil.get("origem_pastas", "")
                tf = float((pacote_funil.get("timings") or {}).get("total_s", 0.0))
            except Exception as e_funil_load:
                st.warning(f"Não foi possível carregar funil do Google Sheets: {e_funil_load}")

            try:
                if df_ag_vis_cache is not None and not df_ag_vis_cache.empty:
                    df_ag_vis = df_ag_vis_cache.copy()
                    origem_ag = origem_ag_cache
                else:
                    raise RuntimeError("Sem agendamentos no pacote funil.")
                n_ag_bruto = len(df_ag_vis)
                df_ag_vis = deduplicar_agendamentos_funil(df_ag_vis)
                st.caption(
                    f"Agendamentos/visitas: {origem_ag} · "
                    f"{n_ag_bruto:,} → {len(df_ag_vis):,} linhas "
                    f"(dedup Código do agendamento)"
                )
            except Exception as e_sf:
                st.warning(
                    f"Agendamentos indisponíveis ({e_sf}). Tentando aba «{ABA_AGENDAMENTOS_VISITAS}»…"
                )
                df_ag_vis = normalizar_colunas(
                    ler_planilha_aba_df(
                        SPREADSHEET_FUNIL_ID, ABA_AGENDAMENTOS_VISITAS, cred_fp
                    )
                )
                df_ag_vis = deduplicar_agendamentos_funil(df_ag_vis)

            try:
                if df_pastas_cache is not None and not df_pastas_cache.empty:
                    df_pastas_funil = df_pastas_cache.copy()
                    origem_pastas = origem_pastas_cache
                else:
                    raise RuntimeError("Sem pastas no pacote funil.")
                n_pas_bruto = len(df_pastas_funil)
                col_envio = achar_coluna_primeiro_envio_analise(df_pastas_funil)
                col_safi = achar_coluna_aprovacao_safi(df_pastas_funil)
                n_com_envio = int(parse_data_serie(df_pastas_funil[col_envio]).notna().sum()) if col_envio else 0
                n_com_safi = int(parse_data_serie(df_pastas_funil[col_safi]).notna().sum()) if col_safi else 0
                n_aprov_filt = len(deduplicar_pastas_aprovadas_funil(df_pastas_funil))
                df_pastas_funil = deduplicar_pastas_funil(df_pastas_funil)
                st.caption(
                    f"Pastas: {origem_pastas} · "
                    f"{n_pas_bruto:,} → {len(df_pastas_funil):,} linhas "
                    f"(dedup Nome da Avaliação) · "
                    f"1º envio: '{col_envio or '?'}' ({n_com_envio:,}) · "
                    f"Aprov. SAFI: '{col_safi or '?'}' ({n_com_safi:,}) · "
                    f"aprovadas (dedup): {n_aprov_filt:,}"
                )
            except Exception as e_sf_p:
                st.warning(f"Pastas indisponíveis ({e_sf_p}). Tentando planilha Sheets…")
                sid_pastas = (SPREADSHEET_PASTAS_ID or "").strip()
                df_pastas_funil, origem_pastas = carregar_df_pastas_funil(
                    SPREADSHEET_FUNIL_ID, sid, sid_pastas, cred_fp
                )
                if df_pastas_funil.empty:
                    st.warning("Pastas não carregadas — funil sem essa etapa.")
                else:
                    df_pastas_funil = deduplicar_pastas_funil(df_pastas_funil)
                    st.caption(f"Pastas (Sheets): {origem_pastas}")

            df_vendas_funil = pd.DataFrame()
            serie_vendas_funil = None
            try:
                base_ven = (
                    df_vendas_raw.copy()
                    if df_vendas_raw is not None and not df_vendas_raw.empty
                    else df_vendas_painel.copy()
                )
                if base_ven.empty:
                    raise RuntimeError("Sem vendas no cache.")
                df_vendas_funil = normalizar_colunas(base_ven)
                n_ven_bruto = len(df_vendas_funil)
                from_raw = df_vendas_raw is not None and not df_vendas_raw.empty
                if from_raw:
                    df_vendas_funil = filtrar_vendas_comerciais(df_vendas_funil)
                n_ven_comercial = len(df_vendas_funil)
                df_vendas_funil = deduplicar_vendas_funil(df_vendas_funil)
                st.caption(
                    f"Vendas (painel): {origem_vendas_painel} · "
                    f"{n_ven_bruto:,} → {n_ven_comercial:,} comerciais → "
                    f"{len(df_vendas_funil):,} linhas (dedup ID da Oportunidade)"
                )
            except Exception as e_sf_v:
                st.warning(
                    f"Vendas do painel indisponíveis ({e_sf_v}). "
                    "Usando base filtrada do painel (Contrato gerado em)."
                )
                serie_vendas_funil = serie_diaria_contratos(base_proj, col_contrato_gerado)

            mapas_funil = montar_mapa_funil_diario(
                df_ag_vis,
                df_pastas_funil if df_pastas_funil is not None else pd.DataFrame(),
                serie_vendas=serie_vendas_funil,
                df_vendas=df_vendas_funil if not df_vendas_funil.empty else None,
            )
            proj_funil = projetar_funil_mes_atual(
                mapas_funil, incluir_mes=True, meta_qtd_mes=meta_qtd_proj,
            )
            if proj_funil:
                render_projecao_funil(proj_funil)
            else:
                st.info("Dados insuficientes para a projeção do funil comercial.")
        except Exception as exc:
            st.warning(f"Não foi possível calcular a projeção do funil: {exc}")
    else:
        st.warning("Coluna 'Contrato gerado em' não encontrada — projeção do funil indisponível.")

    # -------------------------------------------------------------------------
    # Comparativos MTD parciais (dia 1 ao dia atual)
    # -------------------------------------------------------------------------
    render_comparativos_mtd_funil(df_vendas, col_contrato_gerado)

    st.markdown(
        f'<div class="footer" style="text-align:center;padding:1rem 0;color:{COR_TEXTO_PRETO};font-size:0.82rem;">'
        f"Direcional Engenharia · Vendas — Acompanhamento de metas</div>",
        unsafe_allow_html=True,
    )


def preparar_metas_painel(df_metas_raw: pd.DataFrame) -> pd.DataFrame:
    """Metas fundidas e rateadas por coordenador (mesma lógica do painel)."""
    df_metas_melted = melt_metas(df_metas_raw)
    lixo = ["total", "geral", "não informado", "nao informado", "nan", "none", "null", ""]
    if "Região" in df_metas_melted.columns:
        df_metas_melted = df_metas_melted[
            ~df_metas_melted["Região"].astype(str).str.strip().str.lower().isin(lixo)
        ]
    if "Empreendimento" in df_metas_melted.columns:
        df_metas_melted = df_metas_melted[
            ~df_metas_melted["Empreendimento"].astype(str).str.strip().str.lower().isin(lixo)
        ]
    if "Coordenador" not in df_metas_melted.columns:
        df_metas_melted["Coordenador"] = "Não Informado"
    rows_metas = []
    for _, row in df_metas_melted.iterrows():
        coords = [
            c.strip()
            for c in re.split(r"\s+e\s+", str(row.get("Coordenador", "Não Informado")))
            if c.strip()
        ]
        if not coords:
            coords = ["Não Informado"]
        n = len(coords)
        for c in coords:
            new_row = row.copy()
            new_row["Coordenador"] = c
            new_row["Regiao_Coord"] = (
                f"{row['Região']} - {c}"
                if c not in ("Não Informado", "nan", "")
                else str(row["Região"])
            )
            new_row["Meta_Qtd"] = float(new_row["Meta_Qtd"]) / n
            new_row["Meta_VGV"] = float(new_row["Meta_VGV"]) / n
            new_row["_peso_coord"] = 1.0 / n
            rows_metas.append(new_row)
    return pd.DataFrame(rows_metas)


def preparar_vendas_painel(
    df_vendas_raw: pd.DataFrame,
    df_metas: pd.DataFrame,
) -> Dict[str, Any]:
    """Normaliza vendas e gera dataframe pronto para exibição no relatório."""
    avisos: List[str] = []
    df_vendas = normalizar_colunas(df_vendas_raw)
    col_ano = achar_coluna(df_vendas, ["Ano da Venda", "Ano Venda", "Ano"])
    col_mes_looker = achar_coluna(df_vendas, ["Mês da Venda - Looker"])
    col_mes_venda = achar_coluna(df_vendas, ["Mês Venda"])
    col_regiao = achar_coluna(df_vendas, ["Região", "Regiao"])
    col_imobiliaria = achar_coluna(df_vendas, ["Imobiliária", "Imobiliaria"])
    col_canal = achar_coluna(df_vendas, ["Canal"])
    col_valor = achar_coluna(
        df_vendas,
        ["Valor Real de Venda", "Valor Real", "Valor", "Valor_Real_de_Venda__c"],
    )
    col_emp = achar_coluna(df_vendas, ["Empreendimento", "Obra", "Nome do Empreendimento"])
    col_venda_comercial = achar_coluna(df_vendas, ALIASES_VENDA_COMERCIAL)
    col_venda_facilitada = achar_coluna(
        df_vendas, ["Venda facilitada", "Venda Facilitada", "Venda Facilitada?"]
    )
    col_data_venda = achar_coluna(
        df_vendas, ["Data da venda", "Data Venda", "Data de venda", "Data"]
    )
    col_contrato_gerado = achar_coluna(df_vendas, ALIASES_CONTRATO_GERADO)

    if col_emp and col_emp != "Empreendimento":
        df_vendas.rename(columns={col_emp: "Empreendimento"}, inplace=True)
        col_emp = "Empreendimento"
    if col_regiao and col_regiao != "Região":
        df_vendas.rename(columns={col_regiao: "Região"}, inplace=True)
    if col_imobiliaria and col_imobiliaria != "Imobiliária":
        df_vendas.rename(columns={col_imobiliaria: "Imobiliária"}, inplace=True)
        col_imobiliaria = "Imobiliária"
    if col_imobiliaria:
        df_vendas["Canal"] = df_vendas[col_imobiliaria].map(canal_de_imobiliaria)
        col_canal = "Canal"
    if col_emp:
        df_vendas = df_vendas[
            ~df_vendas[col_emp].astype(str).str.strip().str.lower().isin(
                ["total", "geral", "nan", "none", "null", ""]
            )
        ]
    if col_venda_comercial:
        df_vendas = filtrar_vendas_comerciais(df_vendas)
    else:
        avisos.append("Coluna 'Venda Comercial?' não encontrada na base.")
    if col_venda_facilitada:

        def check_facilitada(val: Any) -> str:
            if pd.isna(val):
                return "Normal"
            v_str = str(val).strip().upper()
            if v_str in ("1", "1.0", "SIM", "TRUE"):
                return "Facilitada"
            return "Normal"

        df_vendas["Tipo_Venda"] = df_vendas[col_venda_facilitada].apply(check_facilitada)
    else:
        df_vendas["Tipo_Venda"] = "Normal"

    cols_data_venda = [
        c for c in [col_data_venda, col_contrato_gerado] if c and c in df_vendas.columns
    ]
    df_vendas["_mes"], df_vendas["_ano"] = aplicar_mes_ano_vendas(
        df_vendas,
        cols_data=cols_data_venda,
        col_mes_venda=col_mes_venda,
        col_ano_venda=col_ano,
        col_mes_looker=col_mes_looker,
    )
    df_vendas["_vgv"] = (
        pd.to_numeric(df_vendas[col_valor], errors="coerce").fillna(0.0)
        if col_valor
        else 0.0
    )
    if col_valor:
        mask_txt = df_vendas["_vgv"] == 0
        if mask_txt.any():
            df_vendas.loc[mask_txt, "_vgv"] = (
                df_vendas.loc[mask_txt, col_valor].map(parse_valor_br)
            )
    if col_canal:

        def agrupar_canal(c: Any) -> str:
            bytes_str = str(c).strip().upper()
            prefixo = bytes_str.split("-")[0].strip()
            if prefixo in ["RJ", "RJG"] or bytes_str in ["RJ", "RJG"]:
                return "IMOB"
            return "DV RJ"

        df_vendas["Canal_Agrupado"] = df_vendas[col_canal].apply(agrupar_canal)
    else:
        df_vendas["Canal_Agrupado"] = "DV RJ"

    df_vendas = _distribuir_vendas_coordenador(df_vendas, df_metas)
    df_vendas["_qtd_venda"] = 1.0 * df_vendas["_peso_coord"]
    df_vendas["_vgv_venda"] = df_vendas["_vgv"] * df_vendas["_peso_coord"]
    return {
        "df_vendas": df_vendas,
        "df_vendas_painel": df_vendas.copy(),
        "col_contrato_gerado": col_contrato_gerado,
        "col_canal": col_canal,
        "col_data_venda": col_data_venda,
        "avisos": avisos,
    }


def main() -> None:
    fav = _resolver_png_raiz(FAVICON_ARQUIVO)
    st.set_page_config(
        page_title="Acompanhamento de Vendas | Direcional",
        page_icon=str(fav) if fav else None,
        layout="wide",
    )
    aplicar_estilo()
    _cabecalho_pagina()

    raw_gs = _secrets_connections_gsheets()
    info = montar_service_account_info(raw_gs)
    if not info:
        st.error("Credenciais Google em **[connections.gsheets]** incompletas. Preencha pelo menos **private_key** e **client_email**.")
        return

    sid = spreadsheet_id_de_secrets(raw_gs)
    cred_fp = _fingerprint_credenciais(info)

    try:
        df_metas_raw = ler_planilha_aba_df(sid, WS_METAS, cred_fp)
    except Exception as e:
        st.error(f"Erro ao ler metas na planilha: {str(e)}")
        return

    df_ag_vis_cache: Optional[pd.DataFrame] = None
    df_pastas_cache: Optional[pd.DataFrame] = None
    origem_ag_cache = ""
    origem_pastas_cache = ""

    manifest = vc.ler_manifest(info)
    col_contrato_gerado: Optional[str] = None
    col_canal: Optional[str] = None
    col_data_venda: Optional[str] = None
    df_vendas_raw = pd.DataFrame()
    origem_vendas_painel = ""
    df_vendas = pd.DataFrame()
    df_vendas_painel = pd.DataFrame()

    df_metas = preparar_metas_painel(df_metas_raw)
    df_vendas_painel, origem_vp = _ler_dado_painel("vendas_painel", cred_fp)
    if not df_vendas_painel.empty:
        df_vendas_painel = assegurar_metricas_vendas(df_vendas_painel)
        df_vendas = df_vendas_painel.copy()
        col_contrato_gerado = (
            str(manifest.get("col_contrato_gerado") or "").strip()
            or achar_coluna(df_vendas_painel, ALIASES_CONTRATO_GERADO)
        )
        col_canal = (
            str(manifest.get("col_canal") or "").strip()
            or achar_coluna(df_vendas_painel, ["Canal"])
        )
        col_data_venda = (
            str(manifest.get("col_data_venda") or "").strip()
            or achar_coluna(
                df_vendas_painel,
                ["Data da venda", "Data Venda", "Data de venda", "Data"],
            )
        )
        origem_vendas_painel = origem_vp
    else:
        df_vendas_raw, origem_vendas_painel = _ler_dado_painel("vendas_raw", cred_fp)
        if df_vendas_raw.empty:
            st.error(
                f"Não foi possível carregar vendas do Google Sheets "
                f"(Cache · * ou aba «{WS_VENDAS}»)."
            )
            return
        prep = preparar_vendas_painel(df_vendas_raw, df_metas)
        for aviso in prep.get("avisos") or []:
            st.warning(aviso)
        df_vendas = prep["df_vendas"]
        df_vendas_painel = assegurar_metricas_vendas(prep["df_vendas_painel"])
        col_contrato_gerado = prep.get("col_contrato_gerado")
        col_canal = prep.get("col_canal")
        col_data_venda = prep.get("col_data_venda")

    st.caption(f"Base de vendas: {origem_vendas_painel} · Google Sheets (sem Salesforce ao vivo)")

    df_metas_coord_opts, _ = carregar_metas_coordenadores_com_fallback(
        cred_fp, df_metas, date.today().year, date.today().month,
    )
    filtros_glob = render_filtros_globais(
        df_metas_coord_opts, df_vendas_painel, df_metas_fallback=df_metas,
    )
    df_metas_coord_g, aviso_metas = carregar_metas_coordenadores_com_fallback(
        cred_fp, df_metas, filtros_glob.ano_meta, filtros_glob.mes_meta,
    )
    if aviso_metas:
        st.warning(f"Metas coordenadores: {aviso_metas}")

    df_estoque_kpi = pd.DataFrame()
    try:
        df_estoque_kpi = carregar_estoque_painel_sf()
    except Exception as exc_est:
        st.warning(f"Estoque indisponível no Google Sheets: {exc_est}")
    kpi_est_global, _ = agregar_estoque(df_estoque_kpi)
    render_kpi_resumo_painel(
        kpi_est_global,
        df_metas_coord_g,
        df_vendas_painel,
        col_contrato_gerado,
        filtros_glob.mes_meta,
        filtros_glob.ano_meta,
        emps_sel=filtros_glob.emps_sel or None,
    )
    st.divider()

    tab_metas, tab_dashboard, tab_funil_emp, tab_poder_compra, tab_feedbacks, tab_previsao = st.tabs(
        [
            "Metas & Projeção",
            "Dashboard Comercial",
            "Funil por Empreendimento",
            "Poder de Compra",
            "Feedbacks Comerciais",
            "Previsão de Vendas",
        ]
    )
    with tab_feedbacks:
        if render_aba_feedbacks_comerciais and carregar_feedbacks_comerciais:
            try:
                df_fb = carregar_feedbacks_comerciais(cred_fp)
                render_aba_feedbacks_comerciais(df_fb)
            except Exception as exc:
                st.error(f"Não foi possível carregar Feedbacks Comerciais: {exc}")
        else:
            st.error("Módulo velocimetro_feedbacks_previsao.py não encontrado.")
    with tab_previsao:
        if render_aba_previsao_vendas and carregar_previsao_vendas:
            try:
                df_pr = carregar_previsao_vendas(cred_fp)
                render_aba_previsao_vendas(df_pr, df_vendas_painel, col_contrato_gerado or "")
            except Exception as exc:
                st.error(f"Não foi possível carregar Previsão de Vendas: {exc}")
        else:
            st.error("Módulo velocimetro_feedbacks_previsao.py não encontrado.")
    with tab_funil_emp:
        render_aba_funil_empreendimentos(
            df_metas=df_metas,
            df_vendas=df_vendas_painel,
            col_contrato_gerado=col_contrato_gerado,
            filtros_glob=filtros_glob,
        )
    with tab_poder_compra:
        try:
            pacote_pc = carregar_pacote_poder_compra_sf()
            df_est_pc = carregar_estoque_painel_sf()
            _, enr_pc = agregar_estoque(df_est_pc)
            render_aba_poder_compra(
                _coalesce_dict_df(pacote_pc, "pastas_aprovadas"),
                enr_pc,
                df_vendas_painel,
                _coalesce_dict_df(pacote_pc, "tabela_comprometimento"),
            )
        except Exception as exc:
            st.error(f"Não foi possível carregar aba Poder de Compra: {exc}")
    with tab_dashboard:
        df_est_dc = pd.DataFrame()
        df_cot_dc = pd.DataFrame()
        try:
            df_est_dc = carregar_estoque_painel_sf()
        except Exception as exc:
            st.warning(f"Estoque indisponível no Google Sheets: {exc}")
        try:
            df_cot_dc = carregar_cotacoes_painel_sf()
        except Exception:
            pass
        render_dashboard_comercial(
            df_vendas=df_vendas_painel,
            df_estoque=df_est_dc,
            df_cotacoes=df_cot_dc,
            cred_fp=cred_fp,
            col_data_venda=col_data_venda,
            filtros_glob=filtros_glob,
            df_metas_fallback=df_metas,
        )
    with tab_metas:
        _corpo_painel_metas(
            df_vendas=df_vendas,
            df_metas=df_metas,
            df_vendas_painel=df_vendas_painel,
            df_vendas_raw=df_vendas_raw,
            origem_vendas_painel=origem_vendas_painel,
            col_contrato_gerado=col_contrato_gerado,
            col_canal=col_canal,
            cred_fp=cred_fp,
            sid=sid,
            filtros_glob=filtros_glob,
        )


if __name__ == "__main__":
    main()
