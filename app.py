import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
from pathlib import Path
import glob
from io import BytesIO
import numpy as np
from sp_connector import get_sp_connector

# ============================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================
st.set_page_config(
    page_title="Análise CTOX - PCLs vs Empresas",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================
# CSS MODERNO (INSPIRADO NO SHADCN UI)
# ============================================
st.markdown("""
<style>
    /* Importar fonte Inter */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    /* Reset e Base */
    *, *::before, *::after {
        box-sizing: border-box;
    }
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }
    
    /* Fundo da aplicação */
    .stApp, .main, body {
        background: #FAFAFA !important;
    }
    
    .main .block-container {
        padding: 2rem 2rem 3rem 2rem;
        max-width: 1600px;
        background: #FAFAFA !important;
    }
    
    /* Headers */
    h1 {
        font-size: 2rem !important;
        font-weight: 700 !important;
        color: #09090B !important;
        letter-spacing: -0.02em !important;
    }
    
    h2 {
        font-size: 1.5rem !important;
        font-weight: 600 !important;
        color: #18181B !important;
        letter-spacing: -0.01em !important;
        margin-bottom: 1rem !important;
    }
    
    h3 {
        font-size: 1.125rem !important;
        font-weight: 600 !important;
        color: #27272A !important;
    }
    
    /* Texto padrão maior */
    p, span, div, label {
        font-size: 14px !important;
        color: #3F3F46 !important;
    }
    
    /* Sidebar moderna */
    [data-testid="stSidebar"] {
        background: #FFFFFF !important;
        border-right: 1px solid #E4E4E7 !important;
    }
    
    [data-testid="stSidebar"] > div {
        background: #FFFFFF !important;
        padding: 1.5rem 1rem !important;
    }
    
    /* Labels da sidebar */
    [data-testid="stSidebar"] label {
        font-size: 13px !important;
        font-weight: 600 !important;
        color: #18181B !important;
        text-transform: uppercase !important;
        letter-spacing: 0.05em !important;
    }
    
    [data-testid="stSidebar"] label p {
        font-size: 13px !important;
        font-weight: 600 !important;
        color: #18181B !important;
    }
    
    /* Selects modernos */
    div[data-baseweb="select"] > div {
        background: #FFFFFF !important;
        border: 1px solid #E4E4E7 !important;
        border-radius: 8px !important;
        color: #18181B !important;
        font-weight: 500 !important;
        transition: all 0.2s ease !important;
    }
    
    div[data-baseweb="select"] > div:hover {
        border-color: #22C55E !important;
    }
    
    div[data-baseweb="select"] > div:focus,
    div[data-baseweb="select"] > div[aria-expanded="true"] {
        border-color: #22C55E !important;
        box-shadow: 0 0 0 3px rgba(34, 197, 94, 0.15) !important;
    }
    
    div[data-baseweb="select"] span {
        color: #18181B !important;
        font-weight: 500 !important;
        font-size: 14px !important;
    }
    
    div[data-baseweb="select"] svg {
        fill: #71717A !important;
    }
    
    /* Dropdown do select */
    [data-baseweb="popover"], 
    div[data-baseweb="select"] ul,
    [role="listbox"] {
        background: #FFFFFF !important;
        border: 1px solid #E4E4E7 !important;
        border-radius: 8px !important;
        box-shadow: 0 10px 40px rgba(0,0,0,0.1) !important;
    }
    
    [role="option"] {
        background: #FFFFFF !important;
        color: #18181B !important;
        font-size: 14px !important;
        padding: 10px 12px !important;
    }
    
    [role="option"]:hover,
    [role="option"][aria-selected="true"] {
        background: #F0FDF4 !important;
        color: #166534 !important;
    }
    
    /* Botões */
    .stButton > button {
        background: #18181B !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.625rem 1.25rem !important;
        font-weight: 500 !important;
        font-size: 14px !important;
        transition: all 0.2s ease !important;
    }
    
    .stButton > button:hover {
        background: #27272A !important;
        transform: translateY(-1px) !important;
    }
    
    /* Botões de Download - verde claro */
    .stDownloadButton > button {
        background: #4CAF50 !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 6px !important;
        padding: 0.625rem 1.25rem !important;
        font-weight: 500 !important;
        font-size: 14px !important;
        transition: all 0.2s ease !important;
    }
    
    .stDownloadButton > button:hover {
        background: #45A049 !important;
        transform: translateY(-1px) !important;
    }
    
    /* Cards container */
    [data-testid="stHorizontalBlock"] {
        gap: 1rem !important;
    }
    
    /* Métricas */
    [data-testid="stMetricValue"] {
        font-size: 2rem !important;
        font-weight: 700 !important;
        color: #09090B !important;
    }
    
    [data-testid="stMetricLabel"] {
        font-size: 13px !important;
        font-weight: 600 !important;
        color: #71717A !important;
        text-transform: uppercase !important;
        letter-spacing: 0.05em !important;
    }
    
    /* DataFrames */
    .stDataFrame {
        border-radius: 12px !important;
        overflow: hidden !important;
        border: 1px solid #E4E4E7 !important;
        background: #FFFFFF !important;
    }
    
    .stDataFrame [data-testid="stDataFrameResizable"] {
        border-radius: 12px !important;
        background: #FFFFFF !important;
    }
    
    /* Estilização de tabelas com cores da referência - FORÇAR FUNDO BRANCO */
    .stDataFrame,
    .stDataFrame *,
    .stDataFrame table,
    .stDataFrame table * {
        background-color: #FFFFFF !important;
    }
    
    .stDataFrame table {
        background: #FFFFFF !important;
        background-color: #FFFFFF !important;
        border-collapse: collapse !important;
        width: 100% !important;
    }
    
    /* Cabeçalho da tabela - azul claro */
    .stDataFrame table thead {
        background: #E3F2FD !important;
        background-color: #E3F2FD !important;
    }
    
    .stDataFrame table thead th {
        background: #E3F2FD !important;
        background-color: #E3F2FD !important;
        color: #18181B !important;
        font-weight: 600 !important;
        border: 1px solid #BBDEFB !important;
        padding: 10px 12px !important;
    }
    
    /* Linhas do corpo - alternadas - FORÇAR FUNDO BRANCO/CINZA CLARO */
    .stDataFrame table tbody {
        background: #FFFFFF !important;
        background-color: #FFFFFF !important;
    }
    
    .stDataFrame table tbody tr {
        background: #FFFFFF !important;
        background-color: #FFFFFF !important;
    }
    
    .stDataFrame table tbody tr:nth-child(even) {
        background: #F5F5F5 !important;
        background-color: #F5F5F5 !important;
    }
    
    .stDataFrame table tbody tr:nth-child(odd) {
        background: #FFFFFF !important;
        background-color: #FFFFFF !important;
    }
    
    /* Células da tabela - FORÇAR FUNDO E COR DE TEXTO */
    .stDataFrame table tbody td {
        background: inherit !important;
        background-color: inherit !important;
        color: #18181B !important;
        border: 1px solid #E0E0E0 !important;
        padding: 8px 12px !important;
    }
    
    /* Garantir que células de linhas pares tenham fundo cinza claro */
    .stDataFrame table tbody tr:nth-child(even) td {
        background: #F5F5F5 !important;
        background-color: #F5F5F5 !important;
    }
    
    /* Garantir que células de linhas ímpares tenham fundo branco */
    .stDataFrame table tbody tr:nth-child(odd) td {
        background: #FFFFFF !important;
        background-color: #FFFFFF !important;
    }
    
    .stDataFrame table th {
        background: inherit !important;
        background-color: inherit !important;
        color: #18181B !important;
    }
    
    /* Valores negativos em vermelho - será aplicado via JavaScript */
    .stDataFrame table tbody td.negative-value {
        color: #DC2626 !important;
        font-weight: 500 !important;
    }
    
    /* Estilo padrão para células - garantir cor de texto */
    .stDataFrame table tbody td {
        color: #18181B !important;
    }
    
    /* Sobrescrever qualquer estilo dark mode do Streamlit */
    [data-testid="stDataFrameResizable"] table,
    [data-testid="stDataFrameResizable"] table tbody,
    [data-testid="stDataFrameResizable"] table tbody tr,
    [data-testid="stDataFrameResizable"] table tbody td {
        background-color: #FFFFFF !important;
        color: #18181B !important;
    }
    
    [data-testid="stDataFrameResizable"] table tbody tr:nth-child(even) {
        background-color: #F5F5F5 !important;
    }
    
    [data-testid="stDataFrameResizable"] table tbody tr:nth-child(even) td {
        background-color: #F5F5F5 !important;
    }
</style>
""", unsafe_allow_html=True)

# JavaScript para destacar valores negativos nas tabelas e garantir fundo branco
st.markdown("""
<script>
function highlightNegativeValues() {
    const tables = document.querySelectorAll('.stDataFrame table tbody td');
    tables.forEach(cell => {
        const text = cell.textContent.trim();
        // Verificar se contém sinal negativo ou é um número negativo
        if (text.startsWith('-') || text.match(/^-\d+/) || text.match(/^-\d+\.\d+/)) {
            cell.style.color = '#DC2626';
            cell.style.fontWeight = '500';
        }
        // Garantir que o fundo seja branco ou cinza claro (nunca preto)
        const row = cell.parentElement;
        if (row) {
            const rowIndex = Array.from(row.parentElement.children).indexOf(row);
            if (rowIndex % 2 === 0) {
                cell.style.backgroundColor = '#FFFFFF';
            } else {
                cell.style.backgroundColor = '#F5F5F5';
            }
        }
    });
}

function forceWhiteBackground() {
    // Forçar fundo branco em todas as tabelas
    const dataFrames = document.querySelectorAll('.stDataFrame, [data-testid="stDataFrameResizable"]');
    dataFrames.forEach(df => {
        df.style.backgroundColor = '#FFFFFF';
        const tables = df.querySelectorAll('table');
        tables.forEach(table => {
            table.style.backgroundColor = '#FFFFFF';
            const rows = table.querySelectorAll('tbody tr');
            rows.forEach((row, index) => {
                row.style.backgroundColor = index % 2 === 0 ? '#FFFFFF' : '#F5F5F5';
                const cells = row.querySelectorAll('td');
                cells.forEach(cell => {
                    cell.style.backgroundColor = index % 2 === 0 ? '#FFFFFF' : '#F5F5F5';
                    cell.style.color = '#18181B';
                });
            });
        });
    });
}

// Executar quando a página carregar
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function() {
        highlightNegativeValues();
        forceWhiteBackground();
    });
} else {
    highlightNegativeValues();
    forceWhiteBackground();
}

// Executar após atualizações do Streamlit
const observer = new MutationObserver(function(mutations) {
    highlightNegativeValues();
    forceWhiteBackground();
});
observer.observe(document.body, { childList: true, subtree: true });
</script>
""", unsafe_allow_html=True)

st.markdown("""
<style>
    /* Scrollbar moderna */
    ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    
    ::-webkit-scrollbar-track {
        background: #F4F4F5;
        border-radius: 4px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: #D4D4D8;
        border-radius: 4px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: #A1A1AA;
    }
    
    /* Tabs customizadas */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0 !important;
        background: #F4F4F5 !important;
        border-radius: 10px !important;
        padding: 4px !important;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: transparent !important;
        border-radius: 8px !important;
        color: #71717A !important;
        font-weight: 500 !important;
        font-size: 14px !important;
        padding: 8px 16px !important;
    }
    
    .stTabs [aria-selected="true"] {
        background: #FFFFFF !important;
        color: #18181B !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1) !important;
    }
    
    /* Alertas e mensagens */
    .stSuccess {
        background: #F0FDF4 !important;
        border: 1px solid #BBF7D0 !important;
        border-radius: 8px !important;
        color: #166534 !important;
    }
    
    .stWarning {
        background: #FFFBEB !important;
        border: 1px solid #FEF3C7 !important;
        border-radius: 8px !important;
        color: #92400E !important;
    }
    
    .stInfo {
        background: #EFF6FF !important;
        border: 1px solid #DBEAFE !important;
        border-radius: 8px !important;
        color: #1E40AF !important;
    }
    
    /* Divider */
    hr {
        border: none !important;
        border-top: 1px solid #E4E4E7 !important;
        margin: 2rem 0 !important;
    }
    
    /* Caption/Footer */
    .stCaption, small {
        font-size: 12px !important;
        color: #71717A !important;
    }
    
    /* Barra de navegação superior */
    .main .block-container {
        padding-top: 1rem !important;
    }
    
    /* Selectbox da navegação superior */
    div[data-testid="column"]:has(select[key="nav_selectbox"]) {
        display: flex;
        align-items: center;
        justify-content: flex-end;
    }
    
    div[data-testid="column"]:has(select[key="nav_selectbox"]) label {
        font-size: 14px !important;
        font-weight: 600 !important;
        color: #18181B !important;
        margin-bottom: 0.5rem !important;
    }
</style>
""", unsafe_allow_html=True)

# ============================================
# FUNÇÕES AUXILIARES
# ============================================

def format_number(num):
    """Formata números com separadores de milhar"""
    if num is None or pd.isna(num):
        return "0"
    try:
        num_float = float(num)
        if pd.isna(num_float) or not np.isfinite(num_float):
            return "0"
        return f"{int(num_float):,}".replace(",", ".")
    except (ValueError, TypeError, OverflowError):
        return "0"

def format_currency(value):
    """Formata valores monetários"""
    if pd.isna(value):
        return "-"
    try:
        return f"R$ {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(value)

def format_percentage(value, total):
    """Calcula e formata percentual"""
    if total == 0:
        return "0%"
    try:
        pct = (value / total) * 100
        return f"{pct:.1f}%"
    except:
        return "-"

# ============================================
# COMPONENTES DE UI
# ============================================

def create_metric_card(title, value, subtitle="", trend=None, trend_color="green"):
    """Cria um card de métrica usando st.metric"""
    try:
        # Usar st.metric nativo que é mais confiável
        delta = None
        if trend:
            delta = trend
        
        st.metric(
            label=title,
            value=value,
            delta=delta
        )
        
        if subtitle:
            st.caption(subtitle)
    except Exception as e:
        # Fallback para HTML simples
        try:
            html_card = f'<div style="background:#FFFFFF;border:1px solid #E4E4E7;border-radius:12px;padding:1rem;"><p style="font-size:11px;color:#71717A;margin:0 0 0.5rem 0;">{title}</p><h3 style="font-size:1.5rem;color:#09090B;margin:0;">{value}</h3><p style="font-size:12px;color:#71717A;margin:0.5rem 0 0 0;">{subtitle}</p></div>'
            st.markdown(html_card, unsafe_allow_html=True)
        except:
            st.write(f"**{title}**: {value}")
            if subtitle:
                st.caption(subtitle)

def create_section_header(icon, title, subtitle=""):
    """Cria um cabeçalho de seção moderno"""
    try:
        title_clean = str(title).replace('"', '').replace("'", "")
        icon_clean = str(icon).replace('"', '').replace("'", "")
        
        if subtitle:
            subtitle_clean = str(subtitle).replace('"', '').replace("'", "")
            sub_html = f'<p style="color:#71717A;font-size:14px;margin:0.25rem 0 0 0;">{subtitle_clean}</p>'
        else:
            sub_html = ""
        
        header_html = f'<div style="margin-bottom:1.5rem;"><h2 style="font-size:1.5rem;font-weight:600;color:#18181B;margin:0;">{icon_clean} {title_clean}</h2>{sub_html}</div>'
        st.markdown(header_html, unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Erro: {str(e)}")

# ============================================
# FUNÇÕES DE GRÁFICOS
# ============================================

def create_bar_chart(df, x_col, y_col, title, max_items=12, color='#22C55E'):
    """Cria gráfico de barras horizontal moderno"""
    df_chart = df.copy()
    df_chart = df_chart.sort_values(y_col, ascending=False).head(max_items)
    df_chart = df_chart.sort_values(y_col, ascending=True)
    
    # Limpar valores NaN/None
    df_chart[y_col] = df_chart[y_col].fillna(0)
    df_chart[y_col] = df_chart[y_col].replace([np.inf, -np.inf], 0)
    
    fig = go.Figure()
    
    valores_x = [float(x) if pd.notna(x) and np.isfinite(x) else 0.0 for x in df_chart[y_col]]
    valores_y = [str(y) if pd.notna(y) else "" for y in df_chart[x_col]]
    
    fig.add_trace(go.Bar(
        x=valores_x,
        y=valores_y,
        orientation='h',
        marker=dict(
            color=color,
            line=dict(width=0)
        ),
        text=[format_number(x) for x in valores_x],
        textposition='outside',
        textfont=dict(size=11, color='#3F3F46', family='Inter'),
        hovertemplate=f'<b>%{{y}}</b><br>{y_col}: %{{x:,.0f}}<extra></extra>'
    ))
    
    num_items = len(df_chart)
    height = max(350, num_items * 32 + 80)
    
    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=15, color='#18181B', family='Inter'),
            x=0.5,
            xanchor='center',
            y=0.97
        ),
        xaxis=dict(
            title="",
            showgrid=True,
            gridcolor='#F4F4F5',
            showline=False,
            zeroline=False,
            tickfont=dict(size=11, color='#71717A', family='Inter')
        ),
        yaxis=dict(
            title="",
            showline=False,
            tickfont=dict(size=12, color='#3F3F46', family='Inter')
        ),
        plot_bgcolor='#FFFFFF',
        paper_bgcolor='#FFFFFF',
        height=height,
        margin=dict(l=10, r=60, t=50, b=20),
        showlegend=False,
        font=dict(family='Inter')
    )
    
    return fig

def create_grouped_bar_chart(df, x_col, title, colors=None, max_items=10):
    """Cria gráfico de barras agrupadas"""
    if colors is None:
        colors = {'Ativo': '#22C55E', 'Inativo': '#EF4444'}
    
    try:
        df_chart = df.groupby([x_col, 'status']).size().reset_index(name='Quantidade')
        df_pivot = df_chart.pivot(index=x_col, columns='status', values='Quantidade').fillna(0)
        
        # Garantir que valores sejam numéricos válidos
        for col in df_pivot.columns:
            df_pivot[col] = pd.to_numeric(df_pivot[col], errors='coerce').fillna(0)
        
        if 'Ativo' in df_pivot.columns and 'Inativo' in df_pivot.columns:
            df_pivot['Total'] = df_pivot['Ativo'] + df_pivot['Inativo']
        elif 'Ativo' in df_pivot.columns:
            df_pivot['Total'] = df_pivot['Ativo']
        else:
            df_pivot['Total'] = 0
        
        df_pivot = df_pivot.sort_values('Total', ascending=False).head(max_items)
        df_pivot = df_pivot.drop('Total', axis=1)
        df_pivot = df_pivot.sort_values(df_pivot.columns[0] if len(df_pivot.columns) > 0 else df_pivot.index)
    except Exception as e:
        st.error(f"Erro ao processar dados para gráfico: {str(e)}")
        return go.Figure()
    
    fig = go.Figure()
    
    if 'Ativo' in df_pivot.columns:
        valores_ativos = [float(x) if pd.notna(x) and np.isfinite(x) else 0.0 for x in df_pivot['Ativo']]
        fig.add_trace(go.Bar(
            name='Ativos',
            x=valores_ativos,
            y=df_pivot.index,
            orientation='h',
            marker=dict(color=colors.get('Ativo', '#22C55E')),
            text=[format_number(int(x)) for x in valores_ativos],
            textposition='outside',
            textfont=dict(size=10, color='#3F3F46', family='Inter'),
            hovertemplate='<b>%{y}</b><br>Ativos: %{x:,.0f}<extra></extra>'
        ))
    
    if 'Inativo' in df_pivot.columns:
        valores_inativos = [float(x) if pd.notna(x) and np.isfinite(x) else 0.0 for x in df_pivot['Inativo']]
        fig.add_trace(go.Bar(
            name='Inativos',
            x=valores_inativos,
            y=df_pivot.index,
            orientation='h',
            marker=dict(color=colors.get('Inativo', '#EF4444')),
            text=[format_number(int(x)) for x in valores_inativos],
            textposition='outside',
            textfont=dict(size=10, color='#3F3F46', family='Inter'),
            hovertemplate='<b>%{y}</b><br>Inativos: %{x:,.0f}<extra></extra>'
        ))
    
    num_items = len(df_pivot)
    height = max(400, num_items * 40 + 100)
    
    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=15, color='#18181B', family='Inter'),
            x=0.5,
            xanchor='center',
            y=0.97
        ),
        xaxis=dict(
            title="",
            showgrid=True,
            gridcolor='#F4F4F5',
            showline=False,
            zeroline=False,
            tickfont=dict(size=11, color='#71717A', family='Inter')
        ),
        yaxis=dict(
            title="",
            showline=False,
            tickfont=dict(size=12, color='#3F3F46', family='Inter')
        ),
        barmode='group',
        plot_bgcolor='#FFFFFF',
        paper_bgcolor='#FFFFFF',
        height=height,
        margin=dict(l=10, r=70, t=50, b=20),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
            font=dict(size=12, color='#3F3F46', family='Inter'),
            bgcolor='rgba(255,255,255,0)'
        ),
        font=dict(family='Inter'),
        bargap=0.3,
        bargroupgap=0.1
    )
    
    return fig

def create_progress_card(title, value, total, color="#22C55E"):
    """Cria um card com barra de progresso usando st.progress"""
    try:
        value = float(value) if value is not None else 0
        total = float(total) if total is not None else 0
        percentage = (value / total * 100) if total > 0 else 0
        percentage = min(100, max(0, percentage))
        
        value_str = format_number(int(value))
        total_str = format_number(int(total))
        percentage_val = percentage / 100
        
        # Usar st.metric e st.progress nativos
        st.metric(
            label=title,
            value=value_str
        )
        
        st.progress(percentage_val)
        st.caption(f"{percentage:.1f}% de {total_str}")
    except Exception as e:
        st.error(f"Erro: {str(e)}")

def create_top_list_card(title, data_dict, color="#22C55E"):
    """Cria um card com lista de top items usando gráfico simples"""
    try:
        if not data_dict or len(data_dict) == 0:
            st.markdown(f'<p style="font-size:12px;font-weight:600;color:#71717A;">{title}</p><p style="color:#71717A;">Sem dados</p>', unsafe_allow_html=True)
            return
        
        # Converter dict para DataFrame
        df_top = pd.DataFrame(list(data_dict.items())[:5], columns=['UF', 'Quantidade'])
        df_top = df_top.sort_values('Quantidade', ascending=False)
        
        # Criar gráfico de barras horizontal simples
        fig = go.Figure()
        
        valores = [float(x) if pd.notna(x) and np.isfinite(x) else 0.0 for x in df_top['Quantidade']]
        ufs = [str(x) if pd.notna(x) else "" for x in df_top['UF']]
        
        fig.add_trace(go.Bar(
            x=valores,
            y=ufs,
            orientation='h',
            marker=dict(color=color),
            text=[format_number(int(x)) for x in valores],
            textposition='outside',
            textfont=dict(size=11, color='#3F3F46', family='Inter')
        ))
        
        fig.update_layout(
            title=dict(text=title, font=dict(size=13, color='#18181B', family='Inter'), x=0.5),
            xaxis=dict(title="", showgrid=True, gridcolor='#F4F4F5', tickfont=dict(size=10, color='#71717A')),
            yaxis=dict(title="", tickfont=dict(size=11, color='#3F3F46')),
            plot_bgcolor='#FFFFFF',
            paper_bgcolor='#FFFFFF',
            height=220,
            margin=dict(l=10, r=60, t=40, b=10),
            showlegend=False
        )
        
        st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.error(f"Erro: {str(e)}")

# ============================================
# FUNÇÕES DE DADOS
# ============================================

def get_latest_file_from_sp(sp_connector, folder_path: str):
    """
    Busca o arquivo Excel mais recente em uma pasta do OneDrive/SharePoint.
    
    Args:
        sp_connector: Instância de SPConnector
        folder_path: Caminho relativo a Documents (ex: "Data Analysis/Acumulado de Coletas - Empresas")
    
    Returns:
        Tuple (caminho_completo, nome_arquivo, data_modificacao) ou (None, None, None) se não encontrar
    """
    try:
        # Listar arquivos na pasta
        files = sp_connector.list_files(folder_path)
        
        # Debug: verificar o que foi retornado
        if not files:
            return None, None, None
        
        # Filtrar apenas arquivos .xlsx
        # Verificar se o item tem a propriedade "file" (indica que é arquivo, não pasta)
        excel_files = []
        for f in files:
            # Verificar se é arquivo: tem propriedade "file" (mesmo que vazia) e NÃO tem "folder"
            is_file = "file" in f and "folder" not in f
            name = f.get("name", "")
            # Verificar se termina com .xlsx (case insensitive)
            if is_file and name.lower().endswith(".xlsx"):
                excel_files.append(f)
        
        if not excel_files:
            return None, None, None
        
        # Encontrar o arquivo mais recente pela data de modificação
        latest = max(excel_files, key=lambda x: x.get("lastModifiedDateTime", ""))
        
        # Construir caminho completo
        file_path = f"{folder_path}/{latest['name']}"
        file_name = latest['name']
        last_modified = latest.get("lastModifiedDateTime", "")
        
        return file_path, file_name, last_modified
        
    except FileNotFoundError:
        # Pasta não encontrada
        return None, None, None
    except Exception as e:
        # Outros erros - logar para debug
        import traceback
        print(f"Erro em get_latest_file_from_sp para '{folder_path}': {e}")
        print(traceback.format_exc())
        return None, None, None

@st.cache_data
def load_data():
    """Carrega o arquivo Excel mais recente de cada pasta do SharePoint/OneDrive ou localmente"""
    empresas_data = None
    labs_data = None
    empresas_file = None
    labs_file = None
    errors = []
    
    def read_excel_safe(file_path):
        """Tenta ler Excel com tratamento de erro de permissão"""
        try:
            # Tentativa 1: Leitura normal com openpyxl
            return pd.read_excel(file_path, engine='openpyxl')
        except PermissionError as e:
            # Erro de permissão explícito
            raise PermissionError(f"O arquivo '{file_path.name}' está aberto em outro programa. Feche o arquivo no Excel ou outro programa e recarregue a página.")
        except OSError as e:
            # Verificar se é erro de permissão específico (Errno 13)
            error_str = str(e).lower()
            if 'permission denied' in error_str or 'errno 13' in error_str:
                raise PermissionError(f"O arquivo '{file_path.name}' está aberto em outro programa. Feche o arquivo no Excel ou outro programa e recarregue a página.")
            else:
                # Outro tipo de OSError - re-raise como está
                raise e
        except Exception as e:
            # Re-raise outros erros sem modificar
            raise e
    
    # Tentar conectar ao SharePoint/OneDrive
    sp_connector = None
    try:
        sp_connector = get_sp_connector()
    except Exception as e:
        # Erro ao criar conexão - usar fallback local
        pass
    
    # Se conectado ao SharePoint, buscar arquivos lá
    if sp_connector is not None:
        try:
            # Buscar arquivo mais recente de Empresas
            empresas_path_sp, empresas_name, empresas_date = get_latest_file_from_sp(
                sp_connector, 
                "Data Analysis/Acumulado de Coletas - Empresas"
            )
            
            if empresas_path_sp:
                try:
                    empresas_data = sp_connector.read_excel(empresas_path_sp, engine='openpyxl')
                    # Criar objeto simples para manter compatibilidade com file_info
                    empresas_file = type('FileInfo', (), {'name': empresas_name})()
                except Exception as e:
                    errors.append(f"Erro ao carregar {empresas_name} do SharePoint: {e}")
                    # Tentar fallback local
                    empresas_path_sp = None
            else:
                # Tentar listar arquivos para debug
                try:
                    debug_files = sp_connector.list_files("Data Analysis/Acumulado de Coletas - Empresas")
                    if debug_files:
                        file_names = [f.get('name', 'N/A') for f in debug_files]
                        errors.append(f"⚠️ Nenhum arquivo Excel encontrado em 'Data Analysis/Acumulado de Coletas - Empresas' no SharePoint. Arquivos encontrados: {', '.join(file_names[:5])}")
                    else:
                        errors.append("⚠️ Pasta 'Data Analysis/Acumulado de Coletas - Empresas' não encontrada ou vazia no SharePoint")
                except Exception as debug_e:
                    errors.append(f"⚠️ Erro ao acessar pasta 'Data Analysis/Acumulado de Coletas - Empresas' no SharePoint: {debug_e}")
            
            # Buscar arquivo mais recente de Labs
            labs_path_sp, labs_name, labs_date = get_latest_file_from_sp(
                sp_connector,
                "Data Analysis/Acumulado de Coletas - Labs"
            )
            
            if labs_path_sp:
                try:
                    labs_data = sp_connector.read_excel(labs_path_sp, engine='openpyxl')
                    # Criar objeto simples para manter compatibilidade com file_info
                    labs_file = type('FileInfo', (), {'name': labs_name})()
                except Exception as e:
                    errors.append(f"Erro ao carregar {labs_name} do SharePoint: {e}")
                    # Tentar fallback local
                    labs_path_sp = None
            else:
                # Tentar listar arquivos para debug
                try:
                    debug_files = sp_connector.list_files("Data Analysis/Acumulado de Coletas - Labs")
                    if debug_files:
                        file_names = [f.get('name', 'N/A') for f in debug_files]
                        errors.append(f"⚠️ Nenhum arquivo Excel encontrado em 'Data Analysis/Acumulado de Coletas - Labs' no SharePoint. Arquivos encontrados: {', '.join(file_names[:5])}")
                    else:
                        errors.append("⚠️ Pasta 'Data Analysis/Acumulado de Coletas - Labs' não encontrada ou vazia no SharePoint")
                except Exception as debug_e:
                    errors.append(f"⚠️ Erro ao acessar pasta 'Data Analysis/Acumulado de Coletas - Labs' no SharePoint: {debug_e}")
                
        except Exception as e:
            errors.append(f"Erro ao acessar SharePoint: {e}")
            # Continuar com fallback local
    
    # Fallback: buscar arquivos localmente se não encontrados no SharePoint
    if empresas_data is None:
        empresas_path = Path("Acumulado de Coletas - Empresas")
        if empresas_path.exists():
            excel_files = list(empresas_path.glob("*.xlsx"))
            if excel_files:
                empresas_file = max(excel_files, key=lambda f: f.stat().st_mtime)
                try:
                    empresas_data = read_excel_safe(empresas_file)
                except PermissionError as e:
                    errors.append(f"⚠️ **ERRO DE PERMISSÃO:** O arquivo '{empresas_file.name}' está aberto em outro programa (provavelmente Excel). Por favor, **feche o arquivo** e recarregue esta página (F5).")
                except Exception as e:
                    errors.append(f"Erro ao carregar {empresas_file.name}: {e}")
    
    if labs_data is None:
        labs_path = Path("Acumulado de Coletas - Labs")
        if labs_path.exists():
            excel_files = list(labs_path.glob("*.xlsx"))
            if excel_files:
                labs_file = max(excel_files, key=lambda f: f.stat().st_mtime)
                try:
                    labs_data = read_excel_safe(labs_file)
                except PermissionError as e:
                    errors.append(f"⚠️ **ERRO DE PERMISSÃO:** O arquivo '{labs_file.name}' está aberto em outro programa (provavelmente Excel). Por favor, **feche o arquivo** e recarregue esta página (F5).")
                except Exception as e:
                    errors.append(f"Erro ao carregar {labs_file.name}: {e}")
    
    df_empresas = empresas_data if empresas_data is not None else pd.DataFrame()
    df_labs = labs_data if labs_data is not None else pd.DataFrame()
    
    # Determinar origem dos arquivos (SharePoint ou Local)
    # Verificar se veio do SharePoint (objeto criado dinamicamente, não Path)
    empresas_source = None
    labs_source = None
    
    if empresas_file:
        # Se é um objeto criado dinamicamente (do SharePoint), não é Path
        if not isinstance(empresas_file, Path):
            empresas_source = "sharepoint"
        else:
            empresas_source = "local"
    
    if labs_file:
        # Se é um objeto criado dinamicamente (do SharePoint), não é Path
        if not isinstance(labs_file, Path):
            labs_source = "sharepoint"
        else:
            labs_source = "local"
    
    file_info = {
        'empresas_file': str(empresas_file.name) if empresas_file and hasattr(empresas_file, 'name') else None,
        'labs_file': str(labs_file.name) if labs_file and hasattr(labs_file, 'name') else None,
        'empresas_source': empresas_source,
        'labs_source': labs_source,
    }
    
    return df_empresas, df_labs, errors, file_info

def normalize_column_names(df):
    """Normaliza nomes de colunas baseado nas colunas reais dos arquivos Excel"""
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()
    
    # Mapeamento baseado nas colunas reais dos arquivos Excel
    column_mapping = {
        # Identificação - Empresas
        'cnpj da empresa': 'cnpj',
        'cnpj': 'cnpj',
        'nome da empresa': 'razao_social',
        'razao social': 'razao_social',
        'razão social': 'razao_social',
        'nome fantasia': 'nome_fantasia',
        # Datas
        'data de credenciamento': 'data_credenciamento',
        'data credenciamento': 'data_credenciamento',
        'data da última coleta': 'data_ultima_coleta',
        'data última coleta': 'data_ultima_coleta',
        'última coleta (voucher)': 'ultima_coleta_voucher',
        'última coleta (não-voucher)': 'ultima_coleta_nao_voucher',
        # Dias sem coleta
        'dias sem coleta (voucher)': 'dias_sem_coleta_voucher',
        'dias sem coleta (não-voucher)': 'dias_sem_coleta_nao_voucher',
        # Localização
        'cidade': 'cidade',
        'estado': 'uf',
        'uf': 'uf',
        'representante': 'representante',
        # Vouchers/Coletas - Empresas
        'acumulado coletas voucher': 'acumulado_vouchers',
        'acumulado coletas não-voucher': 'acumulado_coletas_nao_voucher',
        'total coletas voucher 2024': 'vouchers_2024',
        'total coletas voucher 2025': 'vouchers_2025',
        'total coletas não-voucher 2024': 'coletas_nao_voucher_2024',
        'total coletas não-voucher 2025': 'coletas_nao_voucher_2025',
        # Coletas - PCLs
        'acumulado de coletas': 'acumulado_coletas',
        'total de coletas 2024': 'coletas_2024',
        'total de coletas 2025': 'coletas_2025',
    }
    
    for old_name, new_name in column_mapping.items():
        if old_name in df.columns:
            df.rename(columns={old_name: new_name}, inplace=True)
    
    return df

def process_empresas(df_empresas):
    """
    Processa dados de empresas.
    
    CRITÉRIO DE ATIVIDADE:
    Uma empresa é considerada ATIVA se:
    - Última Coleta (Voucher) <= 365 dias OU
    - Última Coleta (Não-Voucher) <= 365 dias OU
    - Dias Sem Coleta (Voucher) <= 365 OU
    - Dias Sem Coleta (Não-Voucher) <= 365
    
    MÉTRICAS CALCULADAS:
    - acumulado_coletas_total: Voucher + Não-Voucher
    - coletas_2025: Voucher 2025 + Não-Voucher 2025
    - status: Ativo/Inativo baseado nos critérios acima
    """
    if df_empresas.empty:
        return df_empresas
    
    df = normalize_column_names(df_empresas)
    
    # Garantir que colunas numéricas existam e sejam numéricas
    if 'acumulado_vouchers' in df.columns:
        df['acumulado_vouchers'] = pd.to_numeric(df['acumulado_vouchers'], errors='coerce').fillna(0)
    else:
        df['acumulado_vouchers'] = 0
    
    if 'acumulado_coletas_nao_voucher' in df.columns:
        df['acumulado_coletas_nao_voucher'] = pd.to_numeric(df['acumulado_coletas_nao_voucher'], errors='coerce').fillna(0)
    else:
        df['acumulado_coletas_nao_voucher'] = 0
    
    # Total de coletas (Voucher + Não-Voucher)
    df['acumulado_coletas_total'] = df['acumulado_vouchers'] + df['acumulado_coletas_nao_voucher']
    
    # Coletas 2025 (Voucher + Não-Voucher)
    vouchers_2025 = pd.to_numeric(df.get('vouchers_2025', 0), errors='coerce').fillna(0)
    nao_voucher_2025 = pd.to_numeric(df.get('coletas_nao_voucher_2025', 0), errors='coerce').fillna(0)
    df['coletas_2025'] = vouchers_2025 + nao_voucher_2025
    
    # Calcular status baseado em AMBOS os tipos de coleta
    # Usar a coluna "Dias Sem Coleta" que já existe no Excel
    dias_voucher = pd.to_numeric(df.get('dias_sem_coleta_voucher', pd.Series([9999]*len(df))), errors='coerce').fillna(9999)
    dias_nao_voucher = pd.to_numeric(df.get('dias_sem_coleta_nao_voucher', pd.Series([9999]*len(df))), errors='coerce').fillna(9999)
    
    # Empresa ativa: menor dos dois dias <= 365
    df['dias_sem_coleta_min'] = pd.concat([dias_voucher, dias_nao_voucher], axis=1).min(axis=1)
    df['status'] = df['dias_sem_coleta_min'].apply(lambda x: 'Ativo' if x <= 365 else 'Inativo')
    
    # Fallback: se não tem dias mas tem coletas > 0, considerar ativo
    df.loc[(df['dias_sem_coleta_min'] > 365) & (df['acumulado_coletas_total'] > 0), 'status'] = 'Ativo'
    
    # Última coleta (a mais recente entre voucher e não-voucher)
    if 'ultima_coleta_voucher' in df.columns:
        df['ultima_coleta_voucher'] = pd.to_datetime(df['ultima_coleta_voucher'], errors='coerce', dayfirst=True)
    if 'ultima_coleta_nao_voucher' in df.columns:
        df['ultima_coleta_nao_voucher'] = pd.to_datetime(df['ultima_coleta_nao_voucher'], errors='coerce', dayfirst=True)
    
    # Criar coluna de última coleta geral (a mais recente)
    if 'ultima_coleta_voucher' in df.columns and 'ultima_coleta_nao_voucher' in df.columns:
        df['ultima_coleta'] = df[['ultima_coleta_voucher', 'ultima_coleta_nao_voucher']].max(axis=1)
    elif 'ultima_coleta_nao_voucher' in df.columns:
        df['ultima_coleta'] = df['ultima_coleta_nao_voucher']
    elif 'ultima_coleta_voucher' in df.columns:
        df['ultima_coleta'] = df['ultima_coleta_voucher']
    
    return df

def process_labs(df_labs):
    """
    Processa dados de labs (PCLs).
    
    CRITÉRIO DE ATIVIDADE:
    Um PCL é considerado ATIVO se:
    - Dias sem coleta <= 90 OU
    - Acumulado de Coletas > 0
    
    Usa as colunas do Excel:
    - 'Ativo em Coletas' (booleano)
    - 'Dias sem coleta' (número)
    - 'Acumulado de Coletas' (número)
    """
    if df_labs.empty:
        return df_labs
    
    df = normalize_column_names(df_labs)
    
    # Garantir que acumulado_coletas seja numérico
    if 'acumulado_coletas' in df.columns:
        df['acumulado_coletas'] = pd.to_numeric(df['acumulado_coletas'], errors='coerce').fillna(0)
    else:
        df['acumulado_coletas'] = 0
    
    # Usar coluna 'Ativo em Coletas' se existir (já vem do Excel)
    if 'ativo em coletas' in df.columns:
        df['status'] = df['ativo em coletas'].apply(lambda x: 'Ativo' if x == True or str(x).lower() == 'true' else 'Inativo')
    elif 'dias sem coleta' in df.columns:
        # Usar dias sem coleta
        dias = pd.to_numeric(df['dias sem coleta'], errors='coerce').fillna(9999)
        df['status'] = dias.apply(lambda x: 'Ativo' if x <= 90 else 'Inativo')
    else:
        # Fallback: usar acumulado
        df['status'] = df['acumulado_coletas'].apply(lambda x: 'Ativo' if x > 0 else 'Inativo')
    
    # Coletas do ano
    if 'coletas_2025' in df.columns:
        df['acumulado_coletas_ano'] = pd.to_numeric(df['coletas_2025'], errors='coerce').fillna(0)
    elif 'acumulado_coletas' in df.columns:
        df['acumulado_coletas_ano'] = df['acumulado_coletas']
    
    return df

def apply_filters(df, estado, cidade):
    """Aplica filtros ao dataframe"""
    df_filtered = df.copy()
    if estado != "Todos" and 'uf' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['uf'] == estado]
    if cidade != "Todas" and 'cidade' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['cidade'] == cidade]
    return df_filtered

def prepare_display_dataframe(df, colunas_desejadas, rename_map):
    """
    Prepara um DataFrame para exibição, evitando colunas duplicadas.
    
    Args:
        df: DataFrame original
        colunas_desejadas: Lista de colunas a exibir (em ordem)
        rename_map: Dicionário de renomeação {nome_original: nome_exibição}
    
    Returns:
        DataFrame pronto para exibição
    """
    # Remover colunas duplicadas do DataFrame original
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    
    # Selecionar apenas colunas que existem (sem duplicatas)
    colunas_existentes = []
    for col in colunas_desejadas:
        if col in df.columns and col not in colunas_existentes:
            colunas_existentes.append(col)
    
    if not colunas_existentes:
        return pd.DataFrame()
    
    # Criar novo DataFrame com colunas selecionadas
    df_display = df[colunas_existentes].copy()
    
    # Renomear colunas
    rename_final = {k: v for k, v in rename_map.items() if k in df_display.columns}
    df_display = df_display.rename(columns=rename_final)
    
    # Garantir que não há duplicatas após renomeação
    df_display = df_display.loc[:, ~df_display.columns.duplicated(keep='first')]
    
    return df_display

# ============================================
# CARREGAR DADOS
# ============================================

with st.spinner("Carregando dados..."):
    df_empresas_raw, df_labs_raw, load_errors, file_info = load_data()
    
    if load_errors:
        for error in load_errors:
            st.warning(error)
    
    if df_empresas_raw.empty and df_labs_raw.empty:
        st.error("⚠️ Nenhum arquivo encontrado nas pastas 'Acumulado de Coletas - Empresas' e 'Acumulado de Coletas - Labs'")
        st.stop()
    
    df_empresas = process_empresas(df_empresas_raw)
    df_labs = process_labs(df_labs_raw)

# ============================================
# SIDEBAR
# ============================================

with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 1rem 0;">
        <h3 style="margin: 0; font-size: 1.25rem;">📊 CTOX Analytics</h3>
        <p style="margin: 0.5rem 0 0 0; font-size: 12px; color: #71717A;">Painel de Análise</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Menu de navegação
    st.markdown("**NAVEGAÇÃO**")
    tipo_analise = st.selectbox(
        "Módulo",
        ["Visão Geral", "Análise de Coletas", "Listagem de PCLs", "Listagem de Empresas", "Análises Específicas"],
        label_visibility="collapsed",
        key="nav_selectbox"
    )
    
    st.markdown("---")
    
    # Filtros
    st.markdown("**FILTROS**")
    
    if not df_labs.empty:
        estados_disponiveis = sorted(df_labs['uf'].dropna().unique().tolist()) if 'uf' in df_labs.columns else []
        cidades_disponiveis = sorted(df_labs['cidade'].dropna().unique().tolist()) if 'cidade' in df_labs.columns else []
    else:
        estados_disponiveis = []
        cidades_disponiveis = []
    
    estado_selecionado = st.selectbox("Estado (UF)", ["Todos"] + estados_disponiveis) if estados_disponiveis else "Todos"
    cidade_selecionada = st.selectbox("Cidade", ["Todas"] + cidades_disponiveis) if cidades_disponiveis else "Todas"
    
    st.markdown("---")
    
    # Info dos arquivos
    st.markdown("**FONTES DE DADOS**")
    
    # Indicador de origem (SharePoint ou Local)
    if file_info.get('labs_source') == 'sharepoint' or file_info.get('empresas_source') == 'sharepoint':
        st.markdown("☁️ **SharePoint Online**")
    else:
        st.markdown("💻 **Local**")
    
    if file_info['labs_file']:
        source_icon = "☁️" if file_info.get('labs_source') == 'sharepoint' else "💻"
        st.caption(f"{source_icon} PCLs: {file_info['labs_file']}")
    if file_info['empresas_file']:
        source_icon = "☁️" if file_info.get('empresas_source') == 'sharepoint' else "💻"
        st.caption(f"{source_icon} Empresas: {file_info['empresas_file']}")

# ============================================
# CONTEÚDO PRINCIPAL
# ============================================

if tipo_analise == "Visão Geral":
    create_section_header("📈", "Visão Geral", "Métricas e indicadores principais da base CTOX")
    
    # Calcular métricas
    total_pcls = len(df_labs) if not df_labs.empty else 0
    pcls_ativos = len(df_labs[df_labs['status'] == 'Ativo']) if not df_labs.empty else 0
    pcls_inativos = total_pcls - pcls_ativos
    pct_pcls_ativos = (pcls_ativos / total_pcls * 100) if total_pcls > 0 else 0
    
    total_empresas = len(df_empresas) if not df_empresas.empty else 0
    empresas_ativas = len(df_empresas[df_empresas['status'] == 'Ativo']) if not df_empresas.empty else 0
    empresas_inativas = total_empresas - empresas_ativas
    pct_empresas_ativas = (empresas_ativas / total_empresas * 100) if total_empresas > 0 else 0
    
    total_coletas = 0
    if not df_labs.empty and 'acumulado_coletas' in df_labs.columns:
        try:
            coletas_sum = df_labs['acumulado_coletas'].sum()
            total_coletas = float(coletas_sum) if pd.notna(coletas_sum) and np.isfinite(coletas_sum) else 0
        except:
            total_coletas = 0
    
    total_vouchers = 0
    if not df_empresas.empty and 'acumulado_vouchers' in df_empresas.columns:
        try:
            vouchers_sum = df_empresas['acumulado_vouchers'].sum()
            total_vouchers = float(vouchers_sum) if pd.notna(vouchers_sum) and np.isfinite(vouchers_sum) else 0
        except:
            total_vouchers = 0
    
    # Linha 1: Métricas principais
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        create_metric_card("Total PCLs", format_number(total_pcls), f"{pct_pcls_ativos:.1f}% ativos", f"↗ {format_number(pcls_ativos)} ativos", "green")
    with col2:
        create_metric_card("PCLs Inativos", format_number(pcls_inativos), "Sem coleta há +90 dias", "", "gray")
    with col3:
        create_metric_card("Total Empresas", format_number(total_empresas), f"{pct_empresas_ativas:.1f}% ativas", f"↗ {format_number(empresas_ativas)} ativas", "green")
    with col4:
        create_metric_card("Empresas Inativas", format_number(empresas_inativas), "Sem uso há +365 dias", "", "gray")
    
    st.markdown("")
    
    # Linha 2: Volumes
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        media_coletas = total_coletas / total_pcls if total_pcls > 0 else 0
        create_metric_card("Total Coletas", format_number(int(total_coletas)), f"Média: {int(media_coletas)}/PCL", "", "gray")
    with col2:
        media_vouchers = total_vouchers / total_empresas if total_empresas > 0 else 0
        create_metric_card("Total Vouchers", format_number(int(total_vouchers)), f"Média: {int(media_vouchers)}/Empresa", "", "gray")
    with col3:
        estados_pcl = df_labs['uf'].nunique() if not df_labs.empty and 'uf' in df_labs.columns else 0
        cidades_pcl = df_labs['cidade'].nunique() if not df_labs.empty and 'cidade' in df_labs.columns else 0
        create_metric_card("Cobertura", f"{estados_pcl} UFs", f"{format_number(cidades_pcl)} cidades", "", "gray")
    with col4:
        razao = total_pcls / total_empresas if total_empresas > 0 else 0
        create_metric_card("Razão PCL/Empresa", f"{razao:.3f}", "PCLs por empresa", "", "gray")
    
    st.markdown("---")
    
    # Cards de Status e Top UFs
    create_section_header("🎯", "Status e Distribuição")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        try:
            pcls_ativos_val = int(pcls_ativos) if pd.notna(pcls_ativos) else 0
            total_pcls_val = int(total_pcls) if pd.notna(total_pcls) else 1
            create_progress_card("PCLs Ativos", pcls_ativos_val, total_pcls_val, "#22C55E")
        except:
            create_progress_card("PCLs Ativos", 0, 1, "#22C55E")
    
    with col2:
        try:
            empresas_ativas_val = int(empresas_ativas) if pd.notna(empresas_ativas) else 0
            total_empresas_val = int(total_empresas) if pd.notna(total_empresas) else 1
            create_progress_card("Empresas Ativas", empresas_ativas_val, total_empresas_val, "#3B82F6")
        except:
            create_progress_card("Empresas Ativas", 0, 1, "#3B82F6")
    
    with col3:
        if not df_labs.empty and 'uf' in df_labs.columns:
            try:
                top5_pcl = df_labs.groupby('uf').size().nlargest(5)
                top5_pcl_dict = {str(k): int(v) for k, v in top5_pcl.to_dict().items() if pd.notna(v) and np.isfinite(v)}
                if top5_pcl_dict:
                    create_top_list_card("Top 5 UFs (PCLs)", top5_pcl_dict, "#22C55E")
            except:
                pass
    
    with col4:
        if not df_empresas.empty and 'uf' in df_empresas.columns:
            try:
                top5_emp = df_empresas.groupby('uf').size().nlargest(5)
                top5_emp_dict = {str(k): int(v) for k, v in top5_emp.to_dict().items() if pd.notna(v) and np.isfinite(v)}
                if top5_emp_dict:
                    create_top_list_card("Top 5 UFs (Empresas)", top5_emp_dict, "#3B82F6")
            except:
                pass
    
    st.markdown("---")
    
    # Gráficos por Estado
    create_section_header("📊", "Distribuição por Estado", "Top 12 estados por volume")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if not df_labs.empty and 'uf' in df_labs.columns:
            df_estado = df_labs.groupby('uf').size().reset_index(name='Quantidade')
            fig = create_bar_chart(df_estado, 'uf', 'Quantidade', "PCLs por Estado", max_items=12, color='#22C55E')
            st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        if not df_empresas.empty and 'uf' in df_empresas.columns:
            df_estado = df_empresas.groupby('uf').size().reset_index(name='Quantidade')
            fig = create_bar_chart(df_estado, 'uf', 'Quantidade', "Empresas por Estado", max_items=12, color='#3B82F6')
            st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("")
    
    # Ativos vs Inativos
    create_section_header("📈", "Ativos vs Inativos por Estado", "Top 10 estados")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if not df_labs.empty and 'uf' in df_labs.columns:
            fig = create_grouped_bar_chart(df_labs, 'uf', "PCLs: Ativos vs Inativos", max_items=10)
            st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        if not df_empresas.empty and 'uf' in df_empresas.columns:
            fig = create_grouped_bar_chart(df_empresas, 'uf', "Empresas: Ativas vs Inativas", max_items=10)
            st.plotly_chart(fig, use_container_width=True)

elif tipo_analise == "Análise de Coletas":
    create_section_header("🔬", "Análise de Coletas", "Métricas detalhadas de coletas por estado e PCL")
    
    if df_labs.empty or 'acumulado_coletas' not in df_labs.columns:
        st.warning("Dados de coletas não disponíveis.")
    else:
        # Métricas de coletas
        try:
            total_coletas = float(df_labs['acumulado_coletas'].sum()) if pd.notna(df_labs['acumulado_coletas'].sum()) else 0.0
            media_coletas = float(df_labs['acumulado_coletas'].mean()) if pd.notna(df_labs['acumulado_coletas'].mean()) else 0.0
            mediana_coletas = float(df_labs['acumulado_coletas'].median()) if pd.notna(df_labs['acumulado_coletas'].median()) else 0.0
            max_coletas = float(df_labs['acumulado_coletas'].max()) if pd.notna(df_labs['acumulado_coletas'].max()) else 0.0
            pcls_sem_coleta = len(df_labs[df_labs['acumulado_coletas'].fillna(0) == 0])
            pcls_com_coleta = len(df_labs[df_labs['acumulado_coletas'].fillna(0) > 0])
        except:
            total_coletas = 0.0
            media_coletas = 0.0
            mediana_coletas = 0.0
            max_coletas = 0.0
            pcls_sem_coleta = 0
            pcls_com_coleta = 0
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_coletas_int = int(total_coletas) if pd.notna(total_coletas) and np.isfinite(total_coletas) else 0
            create_metric_card("Total de Coletas", format_number(total_coletas_int), "Acumulado geral", "", "gray")
        with col2:
            media_int = int(media_coletas) if pd.notna(media_coletas) and np.isfinite(media_coletas) else 0
            mediana_int = int(mediana_coletas) if pd.notna(mediana_coletas) and np.isfinite(mediana_coletas) else 0
            create_metric_card("Média por PCL", format_number(media_int), f"Mediana: {mediana_int}", "", "gray")
        with col3:
            max_int = int(max_coletas) if pd.notna(max_coletas) and np.isfinite(max_coletas) else 0
            create_metric_card("Máximo", format_number(max_int), "Maior volume", "", "gray")
        with col4:
            pct_com_coleta = (pcls_com_coleta / len(df_labs) * 100) if len(df_labs) > 0 else 0
            create_metric_card("PCLs com Coleta", format_number(pcls_com_coleta), f"{pct_com_coleta:.1f}% do total", "", "gray")
        
        st.markdown("---")
        
        # Coletas por Estado
        create_section_header("📊", "Coletas por Estado")
        
        if 'uf' in df_labs.columns:
            coletas_estado = df_labs.groupby('uf').agg({
                'acumulado_coletas': ['sum', 'mean', 'count']
            }).reset_index()
            coletas_estado.columns = ['UF', 'Total Coletas', 'Média por PCL', 'Qtd PCLs']
            coletas_estado = coletas_estado.sort_values('Total Coletas', ascending=False)
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Gráfico de Total de Coletas por Estado
                df_chart = coletas_estado[['UF', 'Total Coletas']].head(12)
                df_chart = df_chart.sort_values('Total Coletas', ascending=True)
                
                fig = go.Figure()
                valores_x = [float(x) if pd.notna(x) and np.isfinite(x) else 0.0 for x in df_chart['Total Coletas']]
                valores_y = [str(y) if pd.notna(y) else "" for y in df_chart['UF']]
                fig.add_trace(go.Bar(
                    x=valores_x,
                    y=valores_y,
                    orientation='h',
                    marker=dict(color='#22C55E'),
                    text=[format_number(int(x)) for x in valores_x],
                    textposition='outside',
                    textfont=dict(size=11, color='#3F3F46', family='Inter')
                ))
                
                fig.update_layout(
                    title=dict(text="Total de Coletas por Estado (Top 12)", font=dict(size=15, color='#18181B', family='Inter'), x=0.5),
                    xaxis=dict(title="", showgrid=True, gridcolor='#F4F4F5', tickfont=dict(size=11, color='#71717A')),
                    yaxis=dict(title="", tickfont=dict(size=12, color='#3F3F46')),
                    plot_bgcolor='#FFFFFF',
                    paper_bgcolor='#FFFFFF',
                    height=450,
                    margin=dict(l=10, r=70, t=50, b=20)
                )
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Gráfico de Média de Coletas por Estado
                df_chart = coletas_estado[['UF', 'Média por PCL']].head(12)
                df_chart = df_chart.sort_values('Média por PCL', ascending=True)
                
                fig = go.Figure()
                valores_x = [float(x) if pd.notna(x) and np.isfinite(x) else 0.0 for x in df_chart['Média por PCL']]
                valores_y = [str(y) if pd.notna(y) else "" for y in df_chart['UF']]
                fig.add_trace(go.Bar(
                    x=valores_x,
                    y=valores_y,
                    orientation='h',
                    marker=dict(color='#3B82F6'),
                    text=[f"{x:.1f}" for x in valores_x],
                    textposition='outside',
                    textfont=dict(size=11, color='#3F3F46', family='Inter')
                ))
                
                fig.update_layout(
                    title=dict(text="Média de Coletas por PCL (Top 12)", font=dict(size=15, color='#18181B', family='Inter'), x=0.5),
                    xaxis=dict(title="", showgrid=True, gridcolor='#F4F4F5', tickfont=dict(size=11, color='#71717A')),
                    yaxis=dict(title="", tickfont=dict(size=12, color='#3F3F46')),
                    plot_bgcolor='#FFFFFF',
                    paper_bgcolor='#FFFFFF',
                    height=450,
                    margin=dict(l=10, r=70, t=50, b=20)
                )
                st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("---")
            
            # Tabela detalhada
            create_section_header("📋", "Tabela Detalhada por Estado")
            
            coletas_estado['Total Coletas'] = coletas_estado['Total Coletas'].fillna(0).astype(int)
            coletas_estado['Média por PCL'] = coletas_estado['Média por PCL'].fillna(0).round(1)
            coletas_estado['Qtd PCLs'] = coletas_estado['Qtd PCLs'].fillna(0).astype(int)
            
            st.dataframe(
                coletas_estado,
                use_container_width=True,
                hide_index=True,
                height=400
            )
            
            st.markdown("---")
            
            # Top PCLs por coletas
            create_section_header("🏆", "Top 20 PCLs por Volume de Coletas")
            
            cols_display = ['razao_social', 'cidade', 'uf', 'acumulado_coletas', 'status']
            cols_available = [c for c in cols_display if c in df_labs.columns]
            
            if cols_available:
                top_pcls = df_labs.nlargest(20, 'acumulado_coletas')[cols_available]
                rename_map = {
                    'razao_social': 'Razão Social',
                    'cidade': 'Cidade',
                    'uf': 'UF',
                    'acumulado_coletas': 'Total Coletas',
                    'status': 'Status'
                }
                top_pcls = top_pcls.rename(columns=rename_map)
                
                st.dataframe(
                    top_pcls,
                    use_container_width=True,
                    hide_index=True,
                    height=500
                )

elif tipo_analise == "Listagem de PCLs":
    create_section_header("🏥", "Listagem de PCLs", "Base completa de laboratórios credenciados")
    
    df_labs_filtered = apply_filters(df_labs, estado_selecionado, cidade_selecionada)
    
    if df_labs_filtered.empty:
        st.warning("Nenhum PCL encontrado com os filtros selecionados.")
    else:
        # Métricas rápidas
        col1, col2, col3, col4 = st.columns(4)
        
        total = len(df_labs_filtered)
        ativos = len(df_labs_filtered[df_labs_filtered['status'] == 'Ativo']) if 'status' in df_labs_filtered.columns else 0
        
        with col1:
            create_metric_card("Total Filtrado", format_number(total), "", "", "gray")
        with col2:
            create_metric_card("Ativos", format_number(ativos), f"{(ativos/total*100):.1f}%" if total > 0 else "0%", "", "gray")
        with col3:
            create_metric_card("Inativos", format_number(total - ativos), "", "", "gray")
        with col4:
            coletas = df_labs_filtered['acumulado_coletas'].sum() if 'acumulado_coletas' in df_labs_filtered.columns else 0
            create_metric_card("Total Coletas", format_number(int(coletas)), "", "", "gray")
        
        st.markdown("---")
        
        # Enriquecer dados com campos calculados
        df_display = df_labs_filtered.copy()
        
        # Criar Cidade+UF
        if 'cidade' in df_display.columns and 'uf' in df_display.columns:
            df_display['cidade_uf'] = df_display['cidade'].fillna('') + '-' + df_display['uf'].fillna('')
        
        # Calcular quantidade de empresas na cidade do PCL
        if 'cidade' in df_display.columns and not df_empresas.empty and 'cidade' in df_empresas.columns:
            empresas_por_cidade = df_empresas.groupby('cidade').size().to_dict()
            df_display['qtd_empresas_cidade'] = df_display['cidade'].map(empresas_por_cidade).fillna(0).astype(int)
            
            # Empresas ativas na cidade
            empresas_ativas_cidade = df_empresas[df_empresas['status'] == 'Ativo'].groupby('cidade').size().to_dict() if 'status' in df_empresas.columns else {}
            df_display['qtd_empresas_ativas_cidade'] = df_display['cidade'].map(empresas_ativas_cidade).fillna(0).astype(int)
            
            # Empresas inativas na cidade
            df_display['qtd_empresas_inativas_cidade'] = df_display['qtd_empresas_cidade'] - df_display['qtd_empresas_ativas_cidade']
            
            # Empresas que utilizaram voucher (acumulado > 0)
            if 'acumulado_vouchers' in df_empresas.columns:
                empresas_com_voucher = df_empresas[df_empresas['acumulado_vouchers'].fillna(0) > 0].groupby('cidade').size().to_dict()
                df_display['qtd_empresas_usaram_voucher'] = df_display['cidade'].map(empresas_com_voucher).fillna(0).astype(int)
                
                empresas_sem_voucher = df_empresas[df_empresas['acumulado_vouchers'].fillna(0) == 0].groupby('cidade').size().to_dict()
                df_display['qtd_empresas_nunca_voucher'] = df_display['cidade'].map(empresas_sem_voucher).fillna(0).astype(int)
                
                # Empresas que nunca utilizaram - separado por representação (INTERNO/EXTERNO)
                if 'representacao' in df_empresas.columns:
                    empresas_nunca_interno = df_empresas[(df_empresas['acumulado_vouchers'].fillna(0) == 0) & 
                                                          (df_empresas['representacao'].str.upper().str.contains('INTERNO', na=False))].groupby('cidade').size().to_dict()
                    df_display['qtd_empresas_nunca_interno'] = df_display['cidade'].map(empresas_nunca_interno).fillna(0).astype(int)
                    
                    empresas_usaram_interno = df_empresas[(df_empresas['acumulado_vouchers'].fillna(0) > 0) & 
                                                           (df_empresas['representacao'].str.upper().str.contains('INTERNO', na=False))].groupby('cidade').size().to_dict()
                    df_display['qtd_empresas_usaram_interno'] = df_display['cidade'].map(empresas_usaram_interno).fillna(0).astype(int)
        
        # Preparar DataFrame para exibição
        colunas_pcl = [
            'cnpj', 'razao_social', 'nome_fantasia', 'cidade', 'uf',
            'data_credenciamento', 'representante',
            'acumulado_coletas', 'acumulado_coletas_ano', 'data_ultima_coleta', 'status',
            'qtd_empresas_cidade'
        ]
        
        rename_map_pcl = {
            'cnpj': 'CNPJ', 'razao_social': 'Razão Social', 'nome_fantasia': 'Nome Fantasia',
            'cidade': 'Cidade', 'uf': 'UF',
            'data_credenciamento': 'Data Credenciamento', 'representante': 'Representante',
            'acumulado_coletas': 'Coletas Total', 'acumulado_coletas_ano': 'Coletas 2025',
            'data_ultima_coleta': 'Última Coleta', 'status': 'Status',
            'qtd_empresas_cidade': 'Empresas na Cidade'
        }
        
        df_final = prepare_display_dataframe(df_display, colunas_pcl, rename_map_pcl)
        
        if not df_final.empty:
            st.dataframe(df_final, use_container_width=True, hide_index=True, height=500)
        else:
            st.warning("Nenhum dado disponível para exibição.")
        
        # Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_display.to_excel(writer, index=False, sheet_name='PCLs')
        
        st.download_button(
            label="📥 Download Excel",
            data=output.getvalue(),
            file_name=f'pcls_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

elif tipo_analise == "Listagem de Empresas":
    create_section_header("🏢", "Listagem de Empresas", "Base completa de empresas credenciadas")
    
    df_empresas_filtered = apply_filters(df_empresas, estado_selecionado, cidade_selecionada)
    
    if df_empresas_filtered.empty:
        st.warning("Nenhuma empresa encontrada com os filtros selecionados.")
    else:
        # Métricas rápidas
        col1, col2, col3, col4 = st.columns(4)
        
        total = len(df_empresas_filtered)
        ativas = len(df_empresas_filtered[df_empresas_filtered['status'] == 'Ativo']) if 'status' in df_empresas_filtered.columns else 0
        
        
        with col1:
            create_metric_card("Total Filtrado", format_number(total), "", "", "gray")
        with col2:
            create_metric_card("Ativas", format_number(ativas), f"{(ativas/total*100):.1f}%" if total > 0 else "0%", "", "gray")
        with col3:
            create_metric_card("Inativas", format_number(total - ativas), "", "", "gray")
        with col4:
            if 'acumulado_vouchers' in df_empresas_filtered.columns:
                vouchers = float(df_empresas_filtered['acumulado_vouchers'].fillna(0).sum())
            else:
                vouchers = 0
            create_metric_card("Total Vouchers", format_number(int(vouchers)), "", "", "gray")
        
        st.markdown("---")
        
        # Enriquecer dados com campos calculados
        df_display = df_empresas_filtered.copy()
        
        
        # Criar Cidade+UF
        if 'cidade' in df_display.columns and 'uf' in df_display.columns:
            df_display['cidade_uf'] = df_display['cidade'].fillna('') + '-' + df_display['uf'].fillna('')
        
        # Calcular quantidade de PCLs na cidade da empresa
        if 'cidade' in df_display.columns and not df_labs.empty and 'cidade' in df_labs.columns:
            pcls_por_cidade = df_labs.groupby('cidade').size().to_dict()
            df_display['qtd_pcls_cidade'] = df_display['cidade'].map(pcls_por_cidade).fillna(0).astype(int)
            
            # PCL na cidade? (Sim/Não)
            df_display['pcl_na_cidade'] = df_display['qtd_pcls_cidade'].apply(lambda x: 'Sim' if x > 0 else 'Não')
            
            # PCLs ativos na cidade
            if 'status' in df_labs.columns:
                pcls_ativos_cidade = df_labs[df_labs['status'] == 'Ativo'].groupby('cidade').size().to_dict()
                df_display['qtd_pcls_ativos_cidade'] = df_display['cidade'].map(pcls_ativos_cidade).fillna(0).astype(int)
                
                # PCLs inativos na cidade
                df_display['qtd_pcls_inativos_cidade'] = df_display['qtd_pcls_cidade'] - df_display['qtd_pcls_ativos_cidade']
                
                # PCLs ativos/inativos por representação (INTERNO/EXTERNO)
                if 'representacao' in df_labs.columns:
                    # PCLs ativos INTERNO
                    pcls_ativos_interno = df_labs[(df_labs['status'] == 'Ativo') & 
                                                   (df_labs['representacao'].str.upper().str.contains('INTERNO', na=False))].groupby('cidade').size().to_dict()
                    df_display['qtd_pcls_ativos_interno'] = df_display['cidade'].map(pcls_ativos_interno).fillna(0).astype(int)
                    
                    # PCLs inativos INTERNO
                    pcls_inativos_interno = df_labs[(df_labs['status'] == 'Inativo') & 
                                                     (df_labs['representacao'].str.upper().str.contains('INTERNO', na=False))].groupby('cidade').size().to_dict()
                    df_display['qtd_pcls_inativos_interno'] = df_display['cidade'].map(pcls_inativos_interno).fillna(0).astype(int)
        
        # Garantir que colunas existam para exibição
        if 'acumulado_vouchers' not in df_display.columns:
            df_display['acumulado_vouchers'] = 0
        if 'acumulado_coletas_nao_voucher' not in df_display.columns:
            df_display['acumulado_coletas_nao_voucher'] = 0
        if 'acumulado_coletas_total' not in df_display.columns:
            df_display['acumulado_coletas_total'] = df_display['acumulado_vouchers'] + df_display['acumulado_coletas_nao_voucher']
        if 'coletas_2025' not in df_display.columns:
            df_display['coletas_2025'] = 0
        
        # Formatar valores para exibição
        for col in ['acumulado_vouchers', 'acumulado_coletas_nao_voucher', 'acumulado_coletas_total', 'coletas_2025']:
            if col in df_display.columns:
                df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0).astype(int)
        
        # Preparar DataFrame para exibição
        colunas_empresa = [
            'cnpj', 'razao_social', 'cidade', 'uf',
            'data_credenciamento', 'representante',
            'acumulado_coletas_total', 'acumulado_vouchers', 'acumulado_coletas_nao_voucher',
            'coletas_2025', 'ultima_coleta', 'status',
            'qtd_pcls_cidade'
        ]
        
        rename_map_empresa = {
            'cnpj': 'CNPJ', 'razao_social': 'Razão Social',
            'cidade': 'Cidade', 'uf': 'UF',
            'data_credenciamento': 'Data Credenciamento', 'representante': 'Representante',
            'acumulado_coletas_total': 'Total Coletas', 'acumulado_vouchers': 'Coletas Voucher',
            'acumulado_coletas_nao_voucher': 'Coletas Não-Voucher', 'coletas_2025': 'Coletas 2025',
            'ultima_coleta': 'Última Coleta', 'status': 'Status',
            'qtd_pcls_cidade': 'PCLs na Cidade'
        }
        
        df_final = prepare_display_dataframe(df_display, colunas_empresa, rename_map_empresa)
        
        if not df_final.empty:
            st.dataframe(df_final, use_container_width=True, hide_index=True, height=500)
        else:
            st.warning("Nenhum dado disponível para exibição.")
        
        # Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_display.to_excel(writer, index=False, sheet_name='Empresas')
        
        st.download_button(
            label="📥 Download Excel",
            data=output.getvalue(),
            file_name=f'empresas_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

elif tipo_analise == "Análises Específicas":
    create_section_header("🔍", "Análises Específicas", "Consultas customizadas conforme demanda")
    
    analise_tipo = st.selectbox(
        "Selecione a análise",
        [
            "1. PCLs em cidades SEM Empresas credenciadas",
            "2. PCLs em cidades COM Empresas INATIVAS (365 dias)",
            "3. Empresas em cidades SEM PCL credenciado",
            "4. Empresas em cidades COM PCL INATIVO (90 dias)",
            "Top PCLs por volume de coletas",
            "Estados com menor cobertura"
        ]
    )
    
    # Análise 1: PCLs em cidades SEM Empresas
    if analise_tipo == "1. PCLs em cidades SEM Empresas credenciadas":
        st.markdown("**Descrição:** Lista de PCLs em cidades que não têm nenhuma empresa credenciada.")
        
        if not df_labs.empty and not df_empresas.empty:
            cidades_com_empresa = set(df_empresas['cidade'].dropna().unique()) if 'cidade' in df_empresas.columns else set()
            df_result = df_labs[~df_labs['cidade'].isin(cidades_com_empresa)] if 'cidade' in df_labs.columns else pd.DataFrame()
            
            if not df_result.empty:
                st.success(f"✅ Encontrados {len(df_result)} PCLs em cidades sem empresas credenciadas")
                
                cols = ['cnpj', 'razao_social', 'nome_fantasia', 'cidade', 'uf', 'status', 'acumulado_coletas', 'data_ultima_coleta']
                cols_available = [c for c in cols if c in df_result.columns]
                df_display = df_result[cols_available].copy()
                
                rename_map = {'cnpj': 'CNPJ', 'razao_social': 'Razão Social', 'nome_fantasia': 'Nome Fantasia',
                              'cidade': 'Cidade', 'uf': 'UF', 'status': 'Status', 
                              'acumulado_coletas': 'Coletas', 'data_ultima_coleta': 'Última Coleta'}
                df_display = df_display.rename(columns=rename_map)
                
                st.dataframe(df_display, use_container_width=True, hide_index=True, height=500)
                
                # Download
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_display.to_excel(writer, index=False, sheet_name='PCLs sem Empresas')
                st.download_button("📥 Download Excel", output.getvalue(), 
                                   f'pcls_sem_empresas_{datetime.now().strftime("%Y%m%d")}.xlsx',
                                   'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            else:
                st.info("✅ Todos os PCLs estão em cidades com empresas credenciadas.")
        else:
            st.warning("Dados insuficientes para análise.")
    
    # Análise 2: PCLs em cidades COM Empresas INATIVAS
    elif analise_tipo == "2. PCLs em cidades COM Empresas INATIVAS (365 dias)":
        st.markdown("**Descrição:** Lista de PCLs em cidades que têm empresas credenciadas, mas TODAS as empresas estão inativas (>365 dias sem voucher).")
        
        if not df_labs.empty and not df_empresas.empty:
            # Encontrar cidades onde TODAS as empresas são inativas
            if 'cidade' in df_empresas.columns and 'status' in df_empresas.columns:
                empresas_por_cidade = df_empresas.groupby('cidade').agg({
                    'status': lambda x: (x == 'Ativo').sum()
                }).reset_index()
                empresas_por_cidade.columns = ['cidade', 'empresas_ativas']
                
                # Cidades com empresas, mas nenhuma ativa
                cidades_empresas_inativas = set(empresas_por_cidade[empresas_por_cidade['empresas_ativas'] == 0]['cidade'])
                
                # PCLs nessas cidades
                df_result = df_labs[df_labs['cidade'].isin(cidades_empresas_inativas)] if 'cidade' in df_labs.columns else pd.DataFrame()
                
                if not df_result.empty:
                    st.warning(f"⚠️ Encontrados {len(df_result)} PCLs em cidades onde todas as empresas estão inativas")
                    
                    cols = ['cnpj', 'razao_social', 'nome_fantasia', 'cidade', 'uf', 'status', 'acumulado_coletas', 'data_ultima_coleta']
                    cols_available = [c for c in cols if c in df_result.columns]
                    df_display = df_result[cols_available].copy()
                    
                    # Adicionar quantidade de empresas inativas na cidade
                    emp_count = df_empresas[df_empresas['cidade'].isin(cidades_empresas_inativas)].groupby('cidade').size().to_dict()
                    df_display['empresas_inativas_cidade'] = df_display['cidade'].map(emp_count).fillna(0).astype(int)
                    
                    rename_map = {'cnpj': 'CNPJ', 'razao_social': 'Razão Social', 'nome_fantasia': 'Nome Fantasia',
                                  'cidade': 'Cidade', 'uf': 'UF', 'status': 'Status PCL', 
                                  'acumulado_coletas': 'Coletas', 'data_ultima_coleta': 'Última Coleta',
                                  'empresas_inativas_cidade': 'Empresas Inativas na Cidade'}
                    df_display = df_display.rename(columns=rename_map)
                    
                    st.dataframe(df_display, use_container_width=True, hide_index=True, height=500)
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_display.to_excel(writer, index=False, sheet_name='PCLs Empresas Inativas')
                    st.download_button("📥 Download Excel", output.getvalue(), 
                                       f'pcls_empresas_inativas_{datetime.now().strftime("%Y%m%d")}.xlsx',
                                       'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                else:
                    st.success("✅ Não há PCLs em cidades onde todas as empresas estão inativas.")
            else:
                st.warning("Colunas necessárias não encontradas nos dados.")
        else:
            st.warning("Dados insuficientes para análise.")
    
    # Análise 3: Empresas sem PCL na cidade
    elif analise_tipo == "3. Empresas em cidades SEM PCL credenciado":
        st.markdown("**Descrição:** Lista de empresas em cidades que não têm nenhum PCL credenciado.")
        
        if not df_labs.empty and not df_empresas.empty:
            cidades_com_pcl = set(df_labs['cidade'].dropna().unique()) if 'cidade' in df_labs.columns else set()
            df_result = df_empresas[~df_empresas['cidade'].isin(cidades_com_pcl)] if 'cidade' in df_empresas.columns else pd.DataFrame()
            
            if not df_result.empty:
                st.error(f"❌ Encontradas {len(df_result)} empresas em cidades sem PCL credenciado")
                
                cols = ['cnpj', 'razao_social', 'nome_fantasia', 'cidade', 'uf', 'status', 'acumulado_vouchers', 'data_ultima_utilizacao']
                cols_available = [c for c in cols if c in df_result.columns]
                df_display = df_result[cols_available].copy()
                
                rename_map = {'cnpj': 'CNPJ', 'razao_social': 'Razão Social', 'nome_fantasia': 'Nome Fantasia',
                              'cidade': 'Cidade', 'uf': 'UF', 'status': 'Status', 
                              'acumulado_vouchers': 'Vouchers', 'data_ultima_utilizacao': 'Última Utilização'}
                df_display = df_display.rename(columns=rename_map)
                
                st.dataframe(df_display, use_container_width=True, hide_index=True, height=500)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_display.to_excel(writer, index=False, sheet_name='Empresas sem PCL')
                st.download_button("📥 Download Excel", output.getvalue(), 
                                   f'empresas_sem_pcl_{datetime.now().strftime("%Y%m%d")}.xlsx',
                                   'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            else:
                st.success("✅ Todas as empresas estão em cidades com PCL credenciado.")
        else:
            st.warning("Dados insuficientes para análise.")
    
    # Análise 4: Empresas em cidades COM PCL INATIVO
    elif analise_tipo == "4. Empresas em cidades COM PCL INATIVO (90 dias)":
        st.markdown("**Descrição:** Lista de empresas em cidades que têm PCL credenciado, mas TODOS os PCLs estão inativos (>90 dias sem coleta).")
        
        if not df_labs.empty and not df_empresas.empty:
            # Encontrar cidades onde TODOS os PCLs são inativos
            if 'cidade' in df_labs.columns and 'status' in df_labs.columns:
                pcls_por_cidade = df_labs.groupby('cidade').agg({
                    'status': lambda x: (x == 'Ativo').sum()
                }).reset_index()
                pcls_por_cidade.columns = ['cidade', 'pcls_ativos']
                
                # Cidades com PCLs, mas nenhum ativo
                cidades_pcls_inativos = set(pcls_por_cidade[pcls_por_cidade['pcls_ativos'] == 0]['cidade'])
                
                # Empresas nessas cidades
                df_result = df_empresas[df_empresas['cidade'].isin(cidades_pcls_inativos)] if 'cidade' in df_empresas.columns else pd.DataFrame()
                
                if not df_result.empty:
                    st.warning(f"⚠️ Encontradas {len(df_result)} empresas em cidades onde todos os PCLs estão inativos")
                    
                    cols = ['cnpj', 'razao_social', 'nome_fantasia', 'cidade', 'uf', 'status', 'acumulado_vouchers', 'data_ultima_utilizacao']
                    cols_available = [c for c in cols if c in df_result.columns]
                    df_display = df_result[cols_available].copy()
                    
                    # Adicionar quantidade de PCLs inativos na cidade
                    pcl_count = df_labs[df_labs['cidade'].isin(cidades_pcls_inativos)].groupby('cidade').size().to_dict()
                    df_display['pcls_inativos_cidade'] = df_display['cidade'].map(pcl_count).fillna(0).astype(int)
                    
                    rename_map = {'cnpj': 'CNPJ', 'razao_social': 'Razão Social', 'nome_fantasia': 'Nome Fantasia',
                                  'cidade': 'Cidade', 'uf': 'UF', 'status': 'Status Empresa', 
                                  'acumulado_vouchers': 'Vouchers', 'data_ultima_utilizacao': 'Última Utilização',
                                  'pcls_inativos_cidade': 'PCLs Inativos na Cidade'}
                    df_display = df_display.rename(columns=rename_map)
                    
                    st.dataframe(df_display, use_container_width=True, hide_index=True, height=500)
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_display.to_excel(writer, index=False, sheet_name='Empresas PCLs Inativos')
                    st.download_button("📥 Download Excel", output.getvalue(), 
                                       f'empresas_pcls_inativos_{datetime.now().strftime("%Y%m%d")}.xlsx',
                                       'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                else:
                    st.success("✅ Não há empresas em cidades onde todos os PCLs estão inativos.")
            else:
                st.warning("Colunas necessárias não encontradas nos dados.")
        else:
            st.warning("Dados insuficientes para análise.")
    
    # Top PCLs por volume
    elif analise_tipo == "Top PCLs por volume de coletas":
        if not df_labs.empty and 'acumulado_coletas' in df_labs.columns:
            top_pcls = df_labs.nlargest(50, 'acumulado_coletas')
            cols = ['cnpj', 'razao_social', 'cidade', 'uf', 'acumulado_coletas', 'status']
            cols_available = [c for c in cols if c in top_pcls.columns]
            df_display = top_pcls[cols_available].copy()
            rename_map = {'cnpj': 'CNPJ', 'razao_social': 'Razão Social', 'cidade': 'Cidade', 
                          'uf': 'UF', 'acumulado_coletas': 'Total Coletas', 'status': 'Status'}
            df_display = df_display.rename(columns=rename_map)
            st.dataframe(df_display, use_container_width=True, hide_index=True, height=500)
    
    # Estados com menor cobertura
    elif analise_tipo == "Estados com menor cobertura":
        if not df_labs.empty and 'uf' in df_labs.columns:
            cobertura = df_labs.groupby('uf').agg({
                'cidade': 'nunique',
                'acumulado_coletas': 'sum' if 'acumulado_coletas' in df_labs.columns else 'count'
            }).reset_index()
            cobertura.columns = ['UF', 'Cidades Atendidas', 'Total Coletas']
            cobertura = cobertura.sort_values('Cidades Atendidas')
            st.dataframe(cobertura, use_container_width=True, hide_index=True)

# ============================================
# RODAPÉ
# ============================================

st.markdown("---")
st.caption(f"📅 Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M')} | CTOX Analytics v2.0")
