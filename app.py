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
from sp_connector import get_sp_connector, SPConnector
from auth_microsoft import MicrosoftAuth, AuthManager, create_login_page, create_user_header

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
# AUTENTICAÇÃO MICROSOFT
# ============================================
try:
    auth = MicrosoftAuth()
    
    # Verificar e renovar token se necessário
    AuthManager.check_and_refresh_token(auth)
    
    # Mostrar página de login se não autenticado
    if not create_login_page(auth):
        st.stop()
        
except Exception as e:
    st.error(f"⚠️ Erro ao inicializar autenticação: {e}")
    st.info("💡 Verifique as configurações em `.streamlit/secrets.toml`")
    st.stop()

# ============================================
# CSS PARA TELA DE CARREGAMENTO
# ============================================
st.markdown("""
<style>
/* Esconder indicador "Running" do Streamlit */
[data-testid="stStatusWidget"] {
    display: none !important;
}

/* Esconder mensagem de cache "Running function..." */
.stSpinner {
    display: none !important;
}

/* Estilo para o container de carregamento */
.loading-container {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 2rem;
    margin: 2rem 0;
}

.loading-title {
    font-size: 1.5rem;
    font-weight: 600;
    margin-bottom: 1rem;
    color: var(--text-color);
}

.loading-subtitle {
    font-size: 1rem;
    color: #6b7280;
    margin-bottom: 0.5rem;
}

/* Animação de pulso para indicador de carregamento */
@keyframes pulse {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.5; }
}

.loading-indicator {
    animation: pulse 1.5s ease-in-out infinite;
}

/* Estilo para toast de sucesso customizado */
.filter-feedback {
    position: fixed;
    top: 80px;
    right: 20px;
    padding: 0.75rem 1.5rem;
    background: linear-gradient(135deg, #22C55E 0%, #16A34A 100%);
    color: white;
    border-radius: 8px;
    font-weight: 500;
    box-shadow: 0 4px 12px rgba(34, 197, 94, 0.3);
    z-index: 9999;
    animation: slideIn 0.3s ease-out;
}

@keyframes slideIn {
    from {
        transform: translateX(100%);
        opacity: 0;
    }
    to {
        transform: translateX(0);
        opacity: 1;
    }
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

def format_cnpj(cnpj):
    """Formata CNPJ com zeros à esquerda e pontuação: XX.XXX.XXX/XXXX-XX"""
    if pd.isna(cnpj) or cnpj is None or str(cnpj).strip() == '':
        return ""
    try:
        # Remover qualquer formatação existente (pontos, barras, hífens)
        cnpj_limpo = str(cnpj).replace('.', '').replace('/', '').replace('-', '').replace(' ', '')
        # Remover possível .0 de números float
        if cnpj_limpo.endswith('.0'):
            cnpj_limpo = cnpj_limpo[:-2]
        # Converter para inteiro e depois para string com zeros à esquerda
        cnpj_numeros = ''.join(filter(str.isdigit, cnpj_limpo))
        # Garantir 14 dígitos com zeros à esquerda
        cnpj_padded = cnpj_numeros.zfill(14)
        # Formatar: XX.XXX.XXX/XXXX-XX
        return f"{cnpj_padded[:2]}.{cnpj_padded[2:5]}.{cnpj_padded[5:8]}/{cnpj_padded[8:12]}-{cnpj_padded[12:14]}"
    except:
        return str(cnpj)

def extrair_bairro(endereco):
    """Extrai o bairro do endereço (última parte após a vírgula)"""
    if pd.isna(endereco) or endereco is None or str(endereco).strip() == '':
        return ""
    try:
        endereco_str = str(endereco).strip()
        # Dividir por vírgula e pegar a última parte
        partes = [p.strip() for p in endereco_str.split(',') if p.strip()]
        if len(partes) >= 2:
            return partes[-1].upper()
        return ""
    except:
        return ""

def extrair_endereco_sem_bairro(endereco):
    """Extrai o endereço sem o bairro (remove última parte após vírgula)"""
    if pd.isna(endereco) or endereco is None or str(endereco).strip() == '':
        return ""
    try:
        endereco_str = str(endereco).strip()
        partes = [p.strip() for p in endereco_str.split(',') if p.strip()]
        if len(partes) >= 2:
            return ', '.join(partes[:-1])
        return endereco_str
    except:
        return str(endereco)

# ============================================
# ADD-961: FUNÇÕES DE VALIDAÇÃO DE QUALIDADE DE DADOS
# ============================================

def validar_cnpj(cnpj):
    """
    Valida CNPJ e retorna (válido, mensagem).
    Verifica se tem 14 dígitos numéricos.
    """
    if pd.isna(cnpj) or cnpj is None or str(cnpj).strip() == '':
        return False, "CNPJ vazio"

    # Remover formatação
    cnpj_limpo = ''.join(filter(str.isdigit, str(cnpj)))

    if len(cnpj_limpo) == 0:
        return False, "CNPJ vazio"

    if len(cnpj_limpo) != 14:
        return False, f"CNPJ com {len(cnpj_limpo)} dígitos (esperado 14)"

    # Verificar se não é uma sequência inválida (todos iguais)
    if len(set(cnpj_limpo)) == 1:
        return False, "CNPJ inválido (dígitos repetidos)"

    return True, "OK"

def validar_cep(cep):
    """
    Valida CEP e retorna (válido, mensagem).
    Verifica se tem 8 dígitos numéricos.
    """
    if pd.isna(cep) or cep is None or str(cep).strip() == '':
        return False, "CEP vazio"

    # Remover formatação
    cep_limpo = ''.join(filter(str.isdigit, str(cep)))

    if len(cep_limpo) == 0:
        return False, "CEP vazio"

    if len(cep_limpo) != 8:
        return False, f"CEP com {len(cep_limpo)} dígitos (esperado 8)"

    return True, "OK"

def validar_uf(uf):
    """
    Valida UF e retorna (válido, mensagem).
    Verifica se é uma sigla válida de estado brasileiro.
    """
    UFS_VALIDAS = {
        'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA',
        'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN',
        'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
    }

    if pd.isna(uf) or uf is None or str(uf).strip() == '':
        return False, "UF vazia"

    uf_upper = str(uf).strip().upper()

    if len(uf_upper) != 2:
        return False, f"UF com {len(uf_upper)} caracteres (esperado 2)"

    if uf_upper not in UFS_VALIDAS:
        return False, f"UF '{uf_upper}' não é válida"

    return True, "OK"

def validar_data(data, nome_campo="Data"):
    """
    Valida uma data e retorna (válido, mensagem).
    Verifica se é uma data válida e não está no futuro.
    """
    if pd.isna(data) or data is None:
        return False, f"{nome_campo} vazia"

    try:
        data_parsed = pd.to_datetime(data, errors='coerce')
        if pd.isna(data_parsed):
            return False, f"{nome_campo} inválida"

        if data_parsed > pd.Timestamp.now():
            return False, f"{nome_campo} no futuro"

        return True, "OK"
    except:
        return False, f"{nome_campo} inválida"

def identificar_ofensores_pcls(df_pcls):
    """
    ADD-961: Identifica problemas de qualidade em PCLs.
    Retorna DataFrame com todos os ofensores encontrados.
    """
    if df_pcls.empty:
        return pd.DataFrame()

    ofensores = []

    for idx, row in df_pcls.iterrows():
        cnpj = row.get('cnpj', '')
        nome = row.get('razao_social', row.get('nome_fantasia', 'N/A'))
        cidade = row.get('cidade', '')
        uf = row.get('uf', '')
        representante = row.get('representante', '')

        # Validar CNPJ (Crítico)
        valido, msg = validar_cnpj(cnpj)
        if not valido:
            ofensores.append({
                'entidade': 'PCL',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'CNPJ',
                'tipo_problema': msg,
                'valor_atual': str(cnpj) if cnpj else '(vazio)',
                'severidade': 'Crítico',
                'severidade_ordem': 1
            })

        # Validar Razão Social (Crítico)
        razao = row.get('razao_social', '')
        if pd.isna(razao) or str(razao).strip() == '' or str(razao).lower() in ['nan', 'none', 'null']:
            ofensores.append({
                'entidade': 'PCL',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'Razão Social',
                'tipo_problema': 'Razão Social vazia',
                'valor_atual': '(vazio)',
                'severidade': 'Crítico',
                'severidade_ordem': 1
            })

        # Validar Cidade (Crítico)
        if pd.isna(cidade) or str(cidade).strip() == '' or str(cidade).lower() in ['nan', 'none', 'null']:
            ofensores.append({
                'entidade': 'PCL',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'Cidade',
                'tipo_problema': 'Cidade vazia',
                'valor_atual': '(vazio)',
                'severidade': 'Crítico',
                'severidade_ordem': 1
            })

        # Validar UF (Crítico)
        valido, msg = validar_uf(uf)
        if not valido:
            ofensores.append({
                'entidade': 'PCL',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'UF',
                'tipo_problema': msg,
                'valor_atual': str(uf) if uf else '(vazio)',
                'severidade': 'Crítico',
                'severidade_ordem': 1
            })

        # Validar Endereço (Alto)
        endereco = row.get('endereco', row.get('endereco_logradouro', ''))
        if pd.isna(endereco) or str(endereco).strip() == '' or str(endereco).lower() in ['nan', 'none', 'null']:
            ofensores.append({
                'entidade': 'PCL',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'Endereço',
                'tipo_problema': 'Endereço vazio',
                'valor_atual': '(vazio)',
                'severidade': 'Alto',
                'severidade_ordem': 2
            })

        # Validar CEP (Alto)
        cep = row.get('cep', '')
        valido, msg = validar_cep(cep)
        if not valido:
            ofensores.append({
                'entidade': 'PCL',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'CEP',
                'tipo_problema': msg,
                'valor_atual': str(cep) if cep else '(vazio)',
                'severidade': 'Alto',
                'severidade_ordem': 2
            })

        # Validar Representante (Médio)
        if pd.isna(representante) or str(representante).strip() == '' or str(representante).lower() in ['nan', 'none', 'null']:
            ofensores.append({
                'entidade': 'PCL',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'Representante',
                'tipo_problema': 'Representante vazio',
                'valor_atual': '(vazio)',
                'severidade': 'Médio',
                'severidade_ordem': 3
            })

        # Validar Data Credenciamento (Médio)
        data_cred = row.get('data_credenciamento', None)
        valido, msg = validar_data(data_cred, "Data Credenciamento")
        if not valido:
            ofensores.append({
                'entidade': 'PCL',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'Data Credenciamento',
                'tipo_problema': msg,
                'valor_atual': str(data_cred) if data_cred else '(vazio)',
                'severidade': 'Médio',
                'severidade_ordem': 3
            })

        # Validar Nome Fantasia (Baixo)
        nome_fantasia = row.get('nome_fantasia', '')
        if pd.isna(nome_fantasia) or str(nome_fantasia).strip() == '' or str(nome_fantasia).lower() in ['nan', 'none', 'null']:
            ofensores.append({
                'entidade': 'PCL',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'Nome Fantasia',
                'tipo_problema': 'Nome Fantasia vazio',
                'valor_atual': '(vazio)',
                'severidade': 'Baixo',
                'severidade_ordem': 4
            })

        # Validar Transportadora (Médio) - PCL sem matriz logística
        transportadora = row.get('transportadora', '')
        if pd.isna(transportadora) or str(transportadora).strip() == '' or str(transportadora).lower() in ['nan', 'none', 'null', 'não cadastrado']:
            ofensores.append({
                'entidade': 'PCL',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'Transportadora',
                'tipo_problema': 'Sem transportadora na matriz logística',
                'valor_atual': '(não cadastrado)',
                'severidade': 'Médio',
                'severidade_ordem': 3
            })

    return pd.DataFrame(ofensores)

def identificar_ofensores_empresas(df_empresas):
    """
    ADD-961: Identifica problemas de qualidade em Empresas.
    Retorna DataFrame com todos os ofensores encontrados.
    """
    if df_empresas.empty:
        return pd.DataFrame()

    ofensores = []

    for idx, row in df_empresas.iterrows():
        cnpj = row.get('cnpj', '')
        nome = row.get('razao_social', row.get('nome', 'N/A'))
        cidade = row.get('cidade', '')
        uf = row.get('uf', '')
        representante = row.get('representante', '')

        # Validar CNPJ (Crítico)
        valido, msg = validar_cnpj(cnpj)
        if not valido:
            ofensores.append({
                'entidade': 'Empresa',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'CNPJ',
                'tipo_problema': msg,
                'valor_atual': str(cnpj) if cnpj else '(vazio)',
                'severidade': 'Crítico',
                'severidade_ordem': 1
            })

        # Validar Nome/Razão Social (Crítico)
        if pd.isna(nome) or str(nome).strip() == '' or str(nome).lower() in ['nan', 'none', 'null']:
            ofensores.append({
                'entidade': 'Empresa',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'Nome/Razão Social',
                'tipo_problema': 'Nome vazio',
                'valor_atual': '(vazio)',
                'severidade': 'Crítico',
                'severidade_ordem': 1
            })

        # Validar Cidade (Crítico)
        if pd.isna(cidade) or str(cidade).strip() == '' or str(cidade).lower() in ['nan', 'none', 'null']:
            ofensores.append({
                'entidade': 'Empresa',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'Cidade',
                'tipo_problema': 'Cidade vazia',
                'valor_atual': '(vazio)',
                'severidade': 'Crítico',
                'severidade_ordem': 1
            })

        # Validar UF (Crítico)
        valido, msg = validar_uf(uf)
        if not valido:
            ofensores.append({
                'entidade': 'Empresa',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'UF',
                'tipo_problema': msg,
                'valor_atual': str(uf) if uf else '(vazio)',
                'severidade': 'Crítico',
                'severidade_ordem': 1
            })

        # Validar Endereço (Alto)
        endereco = row.get('endereco', row.get('endereco_logradouro', ''))
        if pd.isna(endereco) or str(endereco).strip() == '' or str(endereco).lower() in ['nan', 'none', 'null']:
            ofensores.append({
                'entidade': 'Empresa',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'Endereço',
                'tipo_problema': 'Endereço vazio',
                'valor_atual': '(vazio)',
                'severidade': 'Alto',
                'severidade_ordem': 2
            })

        # Validar CEP (Alto)
        cep = row.get('cep', '')
        valido, msg = validar_cep(cep)
        if not valido:
            ofensores.append({
                'entidade': 'Empresa',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'CEP',
                'tipo_problema': msg,
                'valor_atual': str(cep) if cep else '(vazio)',
                'severidade': 'Alto',
                'severidade_ordem': 2
            })

        # Validar Representante (Médio)
        if pd.isna(representante) or str(representante).strip() == '' or str(representante).lower() in ['nan', 'none', 'null']:
            ofensores.append({
                'entidade': 'Empresa',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'Representante',
                'tipo_problema': 'Representante vazio',
                'valor_atual': '(vazio)',
                'severidade': 'Médio',
                'severidade_ordem': 3
            })

        # Validar Data Credenciamento (Médio)
        data_cred = row.get('data_credenciamento', None)
        valido, msg = validar_data(data_cred, "Data Credenciamento")
        if not valido:
            ofensores.append({
                'entidade': 'Empresa',
                'cnpj': cnpj,
                'nome': nome,
                'cidade': cidade,
                'uf': uf,
                'representante': representante,
                'campo_problema': 'Data Credenciamento',
                'tipo_problema': msg,
                'valor_atual': str(data_cred) if data_cred else '(vazio)',
                'severidade': 'Médio',
                'severidade_ordem': 3
            })

    return pd.DataFrame(ofensores)

def identificar_duplicados(df_pcls, df_empresas):
    """
    ADD-961: Identifica CNPJs duplicados em PCLs e Empresas.
    Retorna DataFrame com os duplicados encontrados.
    """
    ofensores = []

    # Duplicados em PCLs
    if not df_pcls.empty and 'cnpj' in df_pcls.columns:
        cnpjs_pcl = df_pcls['cnpj'].dropna()
        duplicados_pcl = cnpjs_pcl[cnpjs_pcl.duplicated(keep=False)]
        for cnpj in duplicados_pcl.unique():
            registros = df_pcls[df_pcls['cnpj'] == cnpj]
            for idx, row in registros.iterrows():
                ofensores.append({
                    'entidade': 'PCL',
                    'cnpj': cnpj,
                    'nome': row.get('razao_social', row.get('nome_fantasia', 'N/A')),
                    'cidade': row.get('cidade', ''),
                    'uf': row.get('uf', ''),
                    'representante': row.get('representante', ''),
                    'campo_problema': 'CNPJ',
                    'tipo_problema': f'CNPJ duplicado ({len(registros)} ocorrências)',
                    'valor_atual': cnpj,
                    'severidade': 'Crítico',
                    'severidade_ordem': 1
                })

    # Duplicados em Empresas
    if not df_empresas.empty and 'cnpj' in df_empresas.columns:
        cnpjs_emp = df_empresas['cnpj'].dropna()
        duplicados_emp = cnpjs_emp[cnpjs_emp.duplicated(keep=False)]
        for cnpj in duplicados_emp.unique():
            registros = df_empresas[df_empresas['cnpj'] == cnpj]
            for idx, row in registros.iterrows():
                ofensores.append({
                    'entidade': 'Empresa',
                    'cnpj': cnpj,
                    'nome': row.get('razao_social', row.get('nome', 'N/A')),
                    'cidade': row.get('cidade', ''),
                    'uf': row.get('uf', ''),
                    'representante': row.get('representante', ''),
                    'campo_problema': 'CNPJ',
                    'tipo_problema': f'CNPJ duplicado ({len(registros)} ocorrências)',
                    'valor_atual': cnpj,
                    'severidade': 'Crítico',
                    'severidade_ordem': 1
                })

    return pd.DataFrame(ofensores)

def consolidar_ofensores(df_pcls, df_empresas):
    """
    ADD-961: Consolida todos os ofensores em um único DataFrame.
    """
    # Identificar ofensores de cada tipo
    ofensores_pcls = identificar_ofensores_pcls(df_pcls)
    ofensores_empresas = identificar_ofensores_empresas(df_empresas)
    ofensores_duplicados = identificar_duplicados(df_pcls, df_empresas)

    # Concatenar todos
    dfs = [df for df in [ofensores_pcls, ofensores_empresas, ofensores_duplicados] if not df.empty]

    if not dfs:
        return pd.DataFrame(columns=[
            'entidade', 'cnpj', 'nome', 'cidade', 'uf', 'representante',
            'campo_problema', 'tipo_problema', 'valor_atual', 'severidade', 'severidade_ordem'
        ])

    df_consolidado = pd.concat(dfs, ignore_index=True)

    # Ordenar por severidade e entidade
    df_consolidado = df_consolidado.sort_values(
        ['severidade_ordem', 'entidade', 'cnpj', 'campo_problema'],
        ascending=[True, True, True, True]
    )

    return df_consolidado

def calcular_metricas_qualidade(df_ofensores, total_pcls, total_empresas):
    """
    ADD-961: Calcula métricas de qualidade de dados.
    """
    total_registros = total_pcls + total_empresas

    if df_ofensores.empty:
        return {
            'total_ofensores': 0,
            'criticos': 0,
            'altos': 0,
            'medios': 0,
            'baixos': 0,
            'pcls_com_problema': 0,
            'empresas_com_problema': 0,
            'score_qualidade': 100.0,
            'problemas_por_tipo': {},
            'problemas_por_representante': {}
        }

    # Contar por severidade
    criticos = len(df_ofensores[df_ofensores['severidade'] == 'Crítico'])
    altos = len(df_ofensores[df_ofensores['severidade'] == 'Alto'])
    medios = len(df_ofensores[df_ofensores['severidade'] == 'Médio'])
    baixos = len(df_ofensores[df_ofensores['severidade'] == 'Baixo'])

    # Registros únicos com problemas
    pcls_com_problema = df_ofensores[df_ofensores['entidade'] == 'PCL']['cnpj'].nunique()
    empresas_com_problema = df_ofensores[df_ofensores['entidade'] == 'Empresa']['cnpj'].nunique()

    # Score de qualidade (penaliza mais problemas críticos)
    registros_com_problema_critico = len(df_ofensores[df_ofensores['severidade'] == 'Crítico']['cnpj'].unique())
    if total_registros > 0:
        score_qualidade = max(0, 100 - (registros_com_problema_critico / total_registros * 100))
    else:
        score_qualidade = 100.0

    # Problemas por tipo
    problemas_por_tipo = df_ofensores.groupby('tipo_problema').size().to_dict()

    # Problemas por representante
    df_rep = df_ofensores[df_ofensores['representante'].notna() & (df_ofensores['representante'] != '')]
    if not df_rep.empty:
        problemas_por_representante = df_rep.groupby('representante').size().to_dict()
    else:
        problemas_por_representante = {}

    return {
        'total_ofensores': len(df_ofensores),
        'criticos': criticos,
        'altos': altos,
        'medios': medios,
        'baixos': baixos,
        'pcls_com_problema': pcls_com_problema,
        'empresas_com_problema': empresas_com_problema,
        'score_qualidade': round(score_qualidade, 1),
        'problemas_por_tipo': problemas_por_tipo,
        'problemas_por_representante': problemas_por_representante
    }

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
        # Fallback simples sem HTML
        st.write(f"**{title}**: {value}")
        if subtitle:
            st.caption(subtitle)

def create_section_header(icon, title, subtitle=""):
    """Cria um cabeçalho de seção moderno"""
    try:
        title_clean = str(title).replace('"', '').replace("'", "")
        icon_clean = str(icon).replace('"', '').replace("'", "")

        # Usar markdown nativo do Streamlit que respeita o tema
        st.markdown(f"## {icon_clean} {title_clean}")
        if subtitle:
            st.caption(subtitle)
    except Exception as e:
        st.error(f"Erro: {str(e)}")

# ============================================
# FUNÇÕES DE GRÁFICOS
# ============================================

def create_bar_chart(df, x_col, y_col, title, max_items=12, color='#22C55E', sort_by_alpha=False):
    """Cria gráfico de barras horizontal moderno
    
    Args:
        sort_by_alpha: Se True, ordena alfabeticamente por x_col. Se False, ordena por y_col (quantidade)
    """
    df_chart = df.copy()
    if sort_by_alpha:
        # Ordenar alfabeticamente por x_col (nome da cidade/estado)
        df_chart = df_chart.sort_values(x_col, ascending=True)
        if max_items > 0:
            df_chart = df_chart.head(max_items)
    else:
        # Ordenar por quantidade (comportamento padrão)
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
        hovertemplate=f'<b>%{{y}}</b><br>{y_col}: %{{x:,.0f}}<extra></extra>'
    ))

    num_items = len(df_chart)
    height = max(350, num_items * 32 + 80)

    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=15),
            x=0.5,
            xanchor='center',
            y=0.97
        ),
        xaxis=dict(
            title="",
            showgrid=True,
            gridcolor='rgba(128,128,128,0.2)',
            showline=False,
            zeroline=False
        ),
        yaxis=dict(
            title="",
            showline=False
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        height=height,
        margin=dict(l=10, r=60, t=50, b=20),
        showlegend=False
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
            hovertemplate='<b>%{y}</b><br>Inativos: %{x:,.0f}<extra></extra>'
        ))

    num_items = len(df_pivot)
    height = max(400, num_items * 40 + 100)

    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=15),
            x=0.5,
            xanchor='center',
            y=0.97
        ),
        xaxis=dict(
            title="",
            showgrid=True,
            gridcolor='rgba(128,128,128,0.2)',
            showline=False,
            zeroline=False
        ),
        yaxis=dict(
            title="",
            showline=False
        ),
        barmode='group',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        height=height,
        margin=dict(l=10, r=70, t=50, b=20),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
            bgcolor='rgba(0,0,0,0)'
        ),
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
            st.markdown(f"**{title}**")
            st.caption("Sem dados")
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
            textposition='outside'
        ))

        fig.update_layout(
            title=dict(text=title, font=dict(size=13), x=0.5),
            xaxis=dict(title="", showgrid=True, gridcolor='rgba(128,128,128,0.2)'),
            yaxis=dict(title=""),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
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

@st.cache_data(ttl=3600)
def load_matriz_logistica():
    """Carrega dados da matriz logística do SharePoint para enriquecer PCLs com transportadora/frequência.

    Retorna:
        tuple: (DataFrame, status_info)
            - DataFrame: dados da matriz ou DataFrame vazio se falhar
            - status_info: dict com 'loaded' (bool), 'source' (str), 'records' (int), 'error' (str ou None)
    """
    df = pd.DataFrame()
    status_info = {
        'loaded': False,
        'source': 'SharePoint',
        'records': 0,
        'error': None,
        'loaded_at': None
    }

    def process_matriz_df(df_raw):
        """Processa o DataFrame da matriz logística - busca colunas por nome exato"""
        colunas_originais = df_raw.columns.tolist()

        # Mapear colunas por nome exato (case insensitive)
        col_map = {}
        for col in colunas_originais:
            col_upper = str(col).upper().strip()
            if col_upper == 'TRANSPORTE':
                col_map['transporte'] = col
            elif col_upper in ['FREQUENCIA', 'FREQUÊNCIA']:
                col_map['frequencia'] = col
            elif col_upper in ['MUNICÍPIO', 'MUNICIPIO']:
                col_map['municipio'] = col
            elif col_upper == 'UF':
                col_map['uf'] = col

        # Debug: mostrar colunas encontradas
        print(f"[MATRIZ LOGISTICA] Colunas do arquivo: {colunas_originais}")
        print(f"[MATRIZ LOGISTICA] Mapeamento: {col_map}")

        # Verificar se encontrou as colunas essenciais
        required = ['transporte', 'frequencia', 'municipio', 'uf']
        missing = [r for r in required if r not in col_map]
        if missing:
            raise ValueError(f"Colunas não encontradas: {missing}. Disponíveis: {colunas_originais}")

        # Criar DataFrame processado com as colunas corretas
        df_processed = pd.DataFrame()
        df_processed['municipio'] = df_raw[col_map['municipio']].astype(str).str.strip().str.upper()
        df_processed['uf'] = df_raw[col_map['uf']].astype(str).str.strip().str.upper()
        df_processed['transporte'] = df_raw[col_map['transporte']].astype(str).str.strip()
        df_processed['frequencia'] = df_raw[col_map['frequencia']].astype(str).str.strip()

        return df_processed

    # Carregar apenas do SharePoint
    try:
        sp_logistica = st.secrets.get("sharepoint_logistica", {})
        if not sp_logistica:
            status_info['error'] = "Configuração [sharepoint_logistica] não encontrada"
        else:
            graph_cfg = st.secrets.get("graph", {})
            connector = SPConnector(
                tenant_id=graph_cfg.get("tenant_id"),
                client_id=graph_cfg.get("client_id"),
                client_secret=graph_cfg.get("client_secret"),
                hostname=sp_logistica.get("hostname"),
                site_path=sp_logistica.get("site_path"),
                library_name=sp_logistica.get("library_name")
            )
            file_path = sp_logistica.get("file_path", "")
            if not file_path:
                status_info['error'] = "Caminho do arquivo não configurado"
            else:
                df_raw = connector.read_excel(file_path)
                df = process_matriz_df(df_raw)
                status_info['loaded'] = True
                status_info['records'] = len(df)
                status_info['loaded_at'] = datetime.now()
    except Exception as e:
        status_info['error'] = str(e)
        print(f"Erro ao carregar matriz logística do SharePoint: {e}")

    return df, status_info

def enrich_pcls_with_logistics(df_labs, df_matriz):
    """Adiciona colunas transportadora e frequencia aos PCLs baseado na cidade/UF"""
    if df_labs.empty or df_matriz.empty:
        df_labs = df_labs.copy()
        df_labs['transportadora'] = 'Não cadastrado'
        df_labs['frequencia'] = 'Não cadastrado'
        return df_labs

    df = df_labs.copy()

    # Normalizar cidade/UF para match
    df['_cidade_norm'] = df['cidade'].fillna('').str.strip().str.upper()
    df['_uf_norm'] = df['uf'].fillna('').str.strip().str.upper()

    # Agrupar transportadoras e frequências por cidade/UF
    def join_unique(series):
        values = series.dropna().astype(str)
        values = [v for v in values if v and v != '-' and v != 'nan' and v.lower() != 'none']
        return ' | '.join(sorted(set(values))) if values else ''

    logistica_grouped = df_matriz.groupby(['municipio', 'uf']).agg({
        'transporte': join_unique,
        'frequencia': join_unique
    }).reset_index()

    # Criar chave para merge
    df['_merge_key'] = df['_cidade_norm'] + '_' + df['_uf_norm']
    logistica_grouped['_merge_key'] = logistica_grouped['municipio'] + '_' + logistica_grouped['uf']

    # Merge via dicionário
    logistica_dict_transporte = dict(zip(logistica_grouped['_merge_key'], logistica_grouped['transporte']))
    logistica_dict_frequencia = dict(zip(logistica_grouped['_merge_key'], logistica_grouped['frequencia']))

    df['transportadora'] = df['_merge_key'].map(logistica_dict_transporte).fillna('')
    df['frequencia'] = df['_merge_key'].map(logistica_dict_frequencia).fillna('')

    # Substituir valores vazios e inválidos por mensagem informativa
    valores_invalidos = ['', 'nan', 'None', 'NaN', 'none', 'NAN']
    df['transportadora'] = df['transportadora'].replace(valores_invalidos, 'Não cadastrado')
    df['frequencia'] = df['frequencia'].replace(valores_invalidos, 'Não cadastrado')

    # Limpar colunas temporárias
    df.drop(columns=['_cidade_norm', '_uf_norm', '_merge_key'], inplace=True)

    return df

@st.cache_data
def get_empresas_por_cidade(df_empresas):
    """Cache de contagem de empresas por cidade"""
    if df_empresas.empty or 'cidade' not in df_empresas.columns:
        return {}, {}, {}

    empresas_por_cidade = df_empresas.groupby('cidade').size().to_dict()

    if 'status' in df_empresas.columns:
        empresas_ativas_cidade = df_empresas[df_empresas['status'] == 'Ativo'].groupby('cidade').size().to_dict()
    else:
        empresas_ativas_cidade = {}

    if 'acumulado_vouchers' in df_empresas.columns:
        empresas_com_voucher = df_empresas[df_empresas['acumulado_vouchers'].fillna(0) > 0].groupby('cidade').size().to_dict()
    else:
        empresas_com_voucher = {}

    return empresas_por_cidade, empresas_ativas_cidade, empresas_com_voucher

@st.cache_data
def get_pcls_por_cidade(df_labs):
    """Cache de contagem de PCLs por cidade"""
    if df_labs.empty or 'cidade' not in df_labs.columns:
        return {}, {}

    pcls_por_cidade = df_labs.groupby('cidade').size().to_dict()

    if 'status' in df_labs.columns:
        pcls_ativos_cidade = df_labs[df_labs['status'] == 'Ativo'].groupby('cidade').size().to_dict()
    else:
        pcls_ativos_cidade = {}

    return pcls_por_cidade, pcls_ativos_cidade

@st.cache_data
def get_listas_filtros(df_labs, df_empresas):
    """Cache das listas de estados e cidades para filtros"""
    # PCLs
    if not df_labs.empty and 'uf' in df_labs.columns:
        estados_pcl = sorted(df_labs['uf'].dropna().unique().tolist())
        cidades_pcl = sorted(df_labs['cidade'].dropna().unique().tolist()) if 'cidade' in df_labs.columns else []
        cidades_por_estado_pcl = df_labs.groupby('uf')['cidade'].apply(lambda x: sorted(x.dropna().unique().tolist())).to_dict()
    else:
        estados_pcl = []
        cidades_pcl = []
        cidades_por_estado_pcl = {}

    # Empresas
    if not df_empresas.empty and 'uf' in df_empresas.columns:
        estados_emp = sorted(df_empresas['uf'].dropna().unique().tolist())
        cidades_emp = sorted(df_empresas['cidade'].dropna().unique().tolist()) if 'cidade' in df_empresas.columns else []
        cidades_por_estado_emp = df_empresas.groupby('uf')['cidade'].apply(lambda x: sorted(x.dropna().unique().tolist())).to_dict()
    else:
        estados_emp = []
        cidades_emp = []
        cidades_por_estado_emp = {}

    return {
        'pcl': {'estados': estados_pcl, 'cidades': cidades_pcl, 'cidades_por_estado': cidades_por_estado_pcl},
        'empresa': {'estados': estados_emp, 'cidades': cidades_emp, 'cidades_por_estado': cidades_por_estado_emp}
    }

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
    # Inclui versões com e sem acentos para compatibilidade de encoding
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
        'data da ultima coleta': 'data_ultima_coleta',
        'data última coleta': 'data_ultima_coleta',
        'última coleta (voucher)': 'ultima_coleta_voucher',
        'ultima coleta (voucher)': 'ultima_coleta_voucher',
        'última coleta (não-voucher)': 'ultima_coleta_nao_voucher',
        'ultima coleta (nao-voucher)': 'ultima_coleta_nao_voucher',
        # Dias sem coleta
        'dias sem coleta': 'dias_sem_coleta',
        'dias sem coleta (voucher)': 'dias_sem_coleta_voucher',
        'dias sem coleta (não-voucher)': 'dias_sem_coleta_nao_voucher',
        'dias sem coleta (nao-voucher)': 'dias_sem_coleta_nao_voucher',
        # Localização
        'endereco': 'endereco',
        'endereço': 'endereco',
        'cidade': 'cidade',
        'estado': 'uf',
        'uf': 'uf',
        'cep': 'cep',
        'representante': 'representante',
        # Vouchers/Coletas - Empresas
        'acumulado coletas voucher': 'acumulado_vouchers',
        'acumulado coletas não-voucher': 'acumulado_coletas_nao_voucher',
        'acumulado coletas nao-voucher': 'acumulado_coletas_nao_voucher',
        'total coletas voucher 2024': 'vouchers_2024',
        'total coletas voucher 2025': 'vouchers_2025',
        'total coletas não-voucher 2024': 'coletas_nao_voucher_2024',
        'total coletas nao-voucher 2024': 'coletas_nao_voucher_2024',
        'total coletas não-voucher 2025': 'coletas_nao_voucher_2025',
        'total coletas nao-voucher 2025': 'coletas_nao_voucher_2025',
        # Coletas - PCLs
        'acumulado de coletas': 'acumulado_coletas',
        'total de coletas 2024': 'coletas_2024',
        'total de coletas 2025': 'coletas_2025',
    }
    
    # Aplicar mapeamento direto
    for old_name, new_name in column_mapping.items():
        if old_name in df.columns:
            df.rename(columns={old_name: new_name}, inplace=True)
    
    # Fallback: procurar por colunas que contenham termos-chave (para lidar com encoding)
    for col in list(df.columns):  # Usar list() para evitar modificar durante iteração
        col_lower = col.lower()
        # Mapeamentos específicos por substring (lidar com diferentes encodings)
        if 'razao_social' not in df.columns and ('raz' in col_lower and 'social' in col_lower):
            df.rename(columns={col: 'razao_social'}, inplace=True)
        elif 'nome_fantasia' not in df.columns and 'nome fantasia' in col_lower:
            df.rename(columns={col: 'nome_fantasia'}, inplace=True)
        elif 'data_ultima_coleta' not in df.columns and 'ltima coleta' in col_lower and 'voucher' not in col_lower:
            df.rename(columns={col: 'data_ultima_coleta'}, inplace=True)
        elif 'ultima_coleta_voucher' not in df.columns and 'ltima coleta' in col_lower and 'voucher' in col_lower and 'n' not in col_lower.split('voucher')[0][-5:]:
            df.rename(columns={col: 'ultima_coleta_voucher'}, inplace=True)
        elif 'ultima_coleta_nao_voucher' not in df.columns and 'ltima coleta' in col_lower and 'voucher' in col_lower and ('n' in col_lower.split('voucher')[0][-5:] or 'nao' in col_lower or 'não' in col_lower):
            df.rename(columns={col: 'ultima_coleta_nao_voucher'}, inplace=True)
        elif 'dias_sem_coleta_voucher' not in df.columns and 'dias sem coleta' in col_lower and 'voucher' in col_lower and 'n' not in col_lower.split('voucher')[0][-5:]:
            df.rename(columns={col: 'dias_sem_coleta_voucher'}, inplace=True)
        elif 'dias_sem_coleta_nao_voucher' not in df.columns and 'dias sem coleta' in col_lower and 'voucher' in col_lower and ('n' in col_lower.split('voucher')[0][-5:] or 'nao' in col_lower or 'não' in col_lower):
            df.rename(columns={col: 'dias_sem_coleta_nao_voucher'}, inplace=True)
    
    return df

def padronizar_nome(nome):
    """
    ADD-960: Padroniza nome de cidade/bairro para exibição consistente.
    Converte para Title Case, tratando preposições em português.
    """
    if pd.isna(nome) or nome is None or str(nome).strip() == '':
        return ''

    nome_str = str(nome).strip()
    # Remover espaços múltiplos
    nome_str = ' '.join(nome_str.split())

    # Converter para Title Case
    palavras = nome_str.lower().split()
    preposicoes = {'de', 'da', 'do', 'das', 'dos', 'e'}

    resultado = []
    for i, palavra in enumerate(palavras):
        # Primeira palavra sempre capitalizada, preposições em minúsculo (exceto no início)
        if i == 0 or palavra not in preposicoes:
            resultado.append(palavra.capitalize())
        else:
            resultado.append(palavra)

    return ' '.join(resultado)

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

    # Formatar CNPJ com zeros à esquerda e pontuação
    if 'cnpj' in df.columns:
        df['cnpj'] = df['cnpj'].apply(format_cnpj)

    # Extrair bairro do endereço (última parte após vírgula)
    if 'endereco' in df.columns:
        df['bairro'] = df['endereco'].apply(extrair_bairro)
        df['endereco_logradouro'] = df['endereco'].apply(extrair_endereco_sem_bairro)
    else:
        df['bairro'] = ''
        df['endereco_logradouro'] = ''
        df['endereco'] = ''

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

    # ADD-960: Padronizar nomes de cidade e bairro
    if 'cidade' in df.columns:
        df['cidade'] = df['cidade'].apply(padronizar_nome)
    if 'bairro' in df.columns:
        df['bairro'] = df['bairro'].apply(padronizar_nome)

    return df

def process_labs(df_labs):
    """
    Processa dados de labs (PCLs).
    
    FILTRO INICIAL:
    - PCLs com "Ativo em Coletas" = False são EXCLUÍDOS completamente (não são considerados)
    
    CRITÉRIO DE ATIVIDADE (para PCLs restantes):
    Um PCL é considerado ATIVO se:
    - "Ativo em Coletas" = True (todos os restantes após filtro) OU
    - Dias sem coleta <= 90 OU
    - Acumulado de Coletas > 0
    
    Usa as colunas do Excel:
    - 'Ativo em Coletas' (booleano) - PCLs com False são excluídos
    - 'Dias sem coleta' (número)
    - 'Acumulado de Coletas' (número)
    """
    if df_labs.empty:
        return df_labs

    df = normalize_column_names(df_labs)

    # EXCLUIR PCLs com "Ativo em Coletas" = False (não devem ser considerados)
    if 'ativo em coletas' in df.columns:
        total_antes = len(df)
        # Filtrar apenas PCLs onde "Ativo em Coletas" é True
        # Considerar True: True, 'true', 'True', 1, '1', 'sim', 'Sim', 'yes', 'Yes'
        df = df[df['ativo em coletas'].apply(
            lambda x: x == True or str(x).lower() in ['true', '1', 'sim', 'yes', 's', 'y']
        )]
        pcls_excluidos_ativos = total_antes - len(df)
        if pcls_excluidos_ativos > 0:
            print(f"Excluídos {pcls_excluidos_ativos} PCLs com 'Ativo em Coletas' = False")

    # Formatar CNPJ com zeros à esquerda e pontuação
    if 'cnpj' in df.columns:
        df['cnpj'] = df['cnpj'].apply(format_cnpj)

    # Extrair bairro do endereço (última parte após vírgula)
    if 'endereco' in df.columns:
        df['bairro'] = df['endereco'].apply(extrair_bairro)
        df['endereco_logradouro'] = df['endereco'].apply(extrair_endereco_sem_bairro)
    else:
        df['bairro'] = ''
        df['endereco_logradouro'] = ''
        df['endereco'] = ''

    # Garantir que acumulado_coletas seja numérico
    if 'acumulado_coletas' in df.columns:
        df['acumulado_coletas'] = pd.to_numeric(df['acumulado_coletas'], errors='coerce').fillna(0)
    else:
        df['acumulado_coletas'] = 0
    
    # Usar coluna 'Ativo em Coletas' se existir (já vem do Excel)
    # Agora todos os PCLs restantes devem ser considerados Ativos (já filtramos os False)
    if 'ativo em coletas' in df.columns:
        df['status'] = 'Ativo'  # Todos os PCLs restantes são ativos (já filtramos os False)
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

    # ADD-960: Padronizar nomes de cidade e bairro
    if 'cidade' in df.columns:
        df['cidade'] = df['cidade'].apply(padronizar_nome)
    if 'bairro' in df.columns:
        df['bairro'] = df['bairro'].apply(padronizar_nome)

    return df

def normalize_city_name(city):
    """Normaliza nome da cidade para comparação (remove espaços extras, converte para minúsculas, remove acentos básicos)"""
    if pd.isna(city) or city == '':
        return ''
    city_str = str(city).strip().lower()
    # Remover espaços múltiplos
    city_str = ' '.join(city_str.split())
    return city_str

def normalize_city_column(df, column='cidade'):
    """Normaliza a coluna de cidade em um dataframe"""
    if column in df.columns:
        df = df.copy()
        df[column] = df[column].apply(normalize_city_name)
    return df

def apply_filters(df, estado, cidade):
    """Aplica filtros ao dataframe (sem copiar desnecessariamente)"""
    if estado == "Todos" and cidade == "Todas":
        return df

    mask = pd.Series([True] * len(df), index=df.index)

    if estado != "Todos" and 'uf' in df.columns:
        mask &= (df['uf'] == estado)
    if cidade != "Todas" and 'cidade' in df.columns:
        mask &= (df['cidade'] == cidade)

    return df[mask]

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

    # Substituir valores vazios/nulos por "Não cadastrado" em colunas de texto
    for col in df_display.columns:
        if df_display[col].dtype == 'object':
            # Substituir None, NaN, strings vazias e variações de "nan"/"none"
            df_display[col] = df_display[col].fillna('Não cadastrado')
            df_display[col] = df_display[col].replace(['', 'nan', 'None', 'NaN', 'none', 'NAN', 'null', 'NULL'], 'Não cadastrado')
            # Tratar strings que são apenas espaços
            df_display[col] = df_display[col].apply(lambda x: 'Não cadastrado' if isinstance(x, str) and x.strip() == '' else x)

    return df_display

# ============================================
# CARREGAR DADOS
# ============================================

# Placeholder para tela de carregamento
loading_placeholder = st.empty()

def show_loading(title, subtitle):
    """Exibe tela de carregamento limpa"""
    loading_placeholder.markdown(f"""
    <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 4rem 2rem; margin: 2rem 0;">
        <div style="width: 50px; height: 50px; border: 4px solid #f3f3f3; border-top: 4px solid #3b82f6; border-radius: 50%; animation: spin 1s linear infinite; margin-bottom: 1.5rem;"></div>
        <h2 style="margin: 0 0 0.5rem 0; font-weight: 500; color: inherit;">{title}</h2>
        <p style="margin: 0; color: #6b7280; font-size: 0.95rem;">{subtitle}</p>
    </div>
    <style>
        @keyframes spin {{
            0% {{ transform: rotate(0deg); }}
            100% {{ transform: rotate(360deg); }}
        }}
    </style>
    """, unsafe_allow_html=True)

show_loading("Carregando dados.", "Conectando ao servidor")
df_empresas_raw, df_labs_raw, load_errors, file_info = load_data()

if load_errors:
    loading_placeholder.empty()
    for error in load_errors:
        st.warning(error)

if df_empresas_raw.empty and df_labs_raw.empty:
    loading_placeholder.empty()
    st.error("⚠️ Nenhum arquivo encontrado nas pastas 'Acumulado de Coletas - Empresas' e 'Acumulado de Coletas - Labs'")
    st.stop()

show_loading("Carregando dados.", "Processando empresas")
df_empresas = process_empresas(df_empresas_raw)

show_loading("Carregando dados.", "Processando PCLs")
df_labs = process_labs(df_labs_raw)

show_loading("Carregando dados.", "Carregando matriz logística")
# Enriquecer PCLs com dados de logística (transportadora e frequência)
df_matriz_logistica, matriz_status = load_matriz_logistica()
df_labs = enrich_pcls_with_logistics(df_labs, df_matriz_logistica)

# ============================================
# LISTA DE EXCEÇÕES - CNPJs a serem excluídos das análises
# ============================================
# CNPJs da própria empresa que não devem ser contabilizados
CNPJS_EXCLUIDOS = [
    '07.339.867/0001-15',  # CAEP - CENTRO AVANÇADO DE ESTUDOS E PESQUISA LTDA (nosso CNPJ)
]

# Remover CNPJs excluídos dos dados de Labs (PCLs)
pcls_excluidos = 0
if not df_labs.empty and 'cnpj' in df_labs.columns:
    total_antes = len(df_labs)
    df_labs = df_labs[~df_labs['cnpj'].isin(CNPJS_EXCLUIDOS)]
    pcls_excluidos = total_antes - len(df_labs)

# Limpar tela de carregamento
loading_placeholder.empty()
    
    # Remover CNPJs excluídos dos dados de Empresas (se necessário)
    # if not df_empresas.empty and 'cnpj' in df_empresas.columns:
    #     df_empresas = df_empresas[~df_empresas['cnpj'].isin(CNPJS_EXCLUIDOS)]

# ============================================
# PRÉ-CALCULAR DADOS PARA PERFORMANCE
# ============================================
# Cachear cálculos pesados para evitar reprocessamento
empresas_por_cidade, empresas_ativas_cidade, empresas_com_voucher = get_empresas_por_cidade(df_empresas)
pcls_por_cidade, pcls_ativos_cidade = get_pcls_por_cidade(df_labs)
listas_filtros = get_listas_filtros(df_labs, df_empresas)

# ============================================
# HEADER - USUÁRIO E LOGOUT NO TOPO
# ============================================

# Header com título e usuário
col_titulo, col_user = st.columns([3, 1])

with col_titulo:
    st.markdown("## 📊 CTOX Analytics")

with col_user:
    user = AuthManager.get_current_user()
    if user:
        display_name = user.get('displayName', 'Usuário')
        col_nome, col_btn = st.columns([2, 1])
        with col_nome:
            st.markdown(f"👤 {display_name}")
        with col_btn:
            if st.button("Sair", key="logout_btn", type="secondary"):
                AuthManager.logout()
                st.rerun()

st.markdown("---")

# ============================================
# CONTEÚDO PRINCIPAL - NAVEGAÇÃO POR ABAS
# ============================================

def _limpar_filtro_visao_estado():
    """Callback para limpar filtro da aba Por Estado"""
    st.session_state["estado_visao_uf"] = "Todos"
    st.toast("Filtros limpos!", icon="🔄")

@st.fragment
def _visao_estado_fragment():
    """Conteúdo da aba Por Estado. Fragment para que ao mudar o selectbox apenas esta parte rerun, mantendo a aba ativa."""
    create_section_header("🗺️", "Visão por Estado", "Painel consolidado com métricas por UF")
    estados_lista = []
    if not df_labs.empty and 'uf' in df_labs.columns:
        estados_lista = sorted(df_labs['uf'].dropna().unique().tolist())
    elif not df_empresas.empty and 'uf' in df_empresas.columns:
        estados_lista = sorted(df_empresas['uf'].dropna().unique().tolist())
    col_filtro, col_btn = st.columns([3, 1])
    with col_filtro:
        estado_filtro = st.selectbox(
            "Selecione o Estado",
            ["Todos"] + estados_lista,
            key="estado_visao_uf"
        )
    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("Limpar Filtros", key="limpar_visao_estado", on_click=_limpar_filtro_visao_estado)
    st.markdown("---")
    if estado_filtro == "Todos":
        df_labs_uf = df_labs.copy()
        df_empresas_uf = df_empresas.copy()
    else:
        df_labs_uf = df_labs[df_labs['uf'] == estado_filtro] if not df_labs.empty and 'uf' in df_labs.columns else pd.DataFrame()
        df_empresas_uf = df_empresas[df_empresas['uf'] == estado_filtro] if not df_empresas.empty and 'uf' in df_empresas.columns else pd.DataFrame()
    total_pcls_uf = len(df_labs_uf) if not df_labs_uf.empty else 0
    pcls_ativos_uf = len(df_labs_uf[df_labs_uf['status'] == 'Ativo']) if not df_labs_uf.empty and 'status' in df_labs_uf.columns else 0
    pcls_inativos_uf = total_pcls_uf - pcls_ativos_uf
    total_empresas_uf = len(df_empresas_uf) if not df_empresas_uf.empty else 0
    empresas_ativas_uf = len(df_empresas_uf[df_empresas_uf['status'] == 'Ativo']) if not df_empresas_uf.empty and 'status' in df_empresas_uf.columns else 0
    empresas_inativas_uf = total_empresas_uf - empresas_ativas_uf
    total_coletas_uf = 0
    if not df_labs_uf.empty and 'acumulado_coletas' in df_labs_uf.columns:
        try:
            coletas_sum = df_labs_uf['acumulado_coletas'].sum()
            total_coletas_uf = float(coletas_sum) if pd.notna(coletas_sum) and np.isfinite(coletas_sum) else 0
        except Exception:
            total_coletas_uf = 0
    total_vouchers_uf = 0
    if not df_empresas_uf.empty and 'acumulado_vouchers' in df_empresas_uf.columns:
        try:
            vouchers_sum = df_empresas_uf['acumulado_vouchers'].sum()
            total_vouchers_uf = float(vouchers_sum) if pd.notna(vouchers_sum) and np.isfinite(vouchers_sum) else 0
        except Exception:
            total_vouchers_uf = 0
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        pct_ativos = (pcls_ativos_uf / total_pcls_uf * 100) if total_pcls_uf > 0 else 0
        create_metric_card("Total PCLs", format_number(total_pcls_uf), f"{pct_ativos:.1f}% ativos", f"↗ {format_number(pcls_ativos_uf)} ativos", "green")
    with col2:
        create_metric_card("PCLs Inativos", format_number(pcls_inativos_uf), "Sem coleta há +90 dias", "", "gray")
    with col3:
        pct_ativas = (empresas_ativas_uf / total_empresas_uf * 100) if total_empresas_uf > 0 else 0
        create_metric_card("Total Empresas", format_number(total_empresas_uf), f"{pct_ativas:.1f}% ativas", f"↗ {format_number(empresas_ativas_uf)} ativas", "green")
    with col4:
        create_metric_card("Empresas Inativas", format_number(empresas_inativas_uf), "Sem uso há +365 dias", "", "gray")
    st.markdown("")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        create_metric_card("Total Coletas", format_number(int(total_coletas_uf)), "Acumulado de coletas", "", "gray")
    with col2:
        create_metric_card("Total Vouchers", format_number(int(total_vouchers_uf)), "Acumulado de vouchers", "", "gray")
    with col3:
        cidades_pcl = df_labs_uf['cidade'].nunique() if not df_labs_uf.empty and 'cidade' in df_labs_uf.columns else 0
        create_metric_card("Cidades c/ PCL", format_number(cidades_pcl), "Cobertura de PCLs", "", "gray")
    with col4:
        cidades_emp = df_empresas_uf['cidade'].nunique() if not df_empresas_uf.empty and 'cidade' in df_empresas_uf.columns else 0
        create_metric_card("Cidades c/ Empresa", format_number(cidades_emp), "Cobertura de empresas", "", "gray")
    st.markdown("---")
    if estado_filtro != "Todos":
        create_section_header("📊", f"Distribuição por Cidade - {estado_filtro}")
        # Constante: gráfico mostra só as top N cidades para ficar legível; lista completa fica na tabela
        MAX_CIDADES_GRAFICO = 15
        col1, col2 = st.columns(2)
        with col1:
            if not df_labs_uf.empty and 'cidade' in df_labs_uf.columns:
                # Normalizar cidade para unificar grafias (ex: "ÁGUAS DE SANTA BARBARA" e "Águas de Santa Bárbara")
                df_pcl = df_labs_uf.copy()
                df_pcl["cidade_norm"] = df_pcl["cidade"].apply(normalize_city_name)
                df_pcl = df_pcl[df_pcl["cidade_norm"].astype(str).str.strip() != ""]
                agg_pcl = df_pcl.groupby("cidade_norm").agg(
                    Quantidade=("cidade", "count"),
                    cidade=("cidade", "first")
                ).reset_index(drop=True)
                df_cidade_pcl = agg_pcl[["cidade", "Quantidade"]].copy()
                df_cidade_pcl = df_cidade_pcl.sort_values("cidade", ascending=True)
                # Gráfico: apenas Top 15 por quantidade (mais legível)
                fig = create_bar_chart(
                    df_cidade_pcl, "cidade", "Quantidade",
                    f"Top {MAX_CIDADES_GRAFICO} cidades (PCLs) - {estado_filtro}",
                    max_items=MAX_CIDADES_GRAFICO, color="#22C55E", sort_by_alpha=False
                )
                st.plotly_chart(fig, use_container_width=True)
                # Tabela com todas as cidades (ordenada alfabeticamente)
                st.caption("Todas as cidades (PCLs)")
                st.dataframe(
                    df_cidade_pcl.rename(columns={"cidade": "Cidade", "Quantidade": "PCLs"}),
                    use_container_width=True, hide_index=True, height=280
                )
            else:
                st.info("Sem dados de PCLs para este estado.")
        with col2:
            if not df_empresas_uf.empty and 'cidade' in df_empresas_uf.columns:
                df_emp = df_empresas_uf.copy()
                df_emp["cidade_norm"] = df_emp["cidade"].apply(normalize_city_name)
                df_emp = df_emp[df_emp["cidade_norm"].astype(str).str.strip() != ""]
                agg_emp = df_emp.groupby("cidade_norm").agg(
                    Quantidade=("cidade", "count"),
                    cidade=("cidade", "first")
                ).reset_index(drop=True)
                df_cidade_emp = agg_emp[["cidade", "Quantidade"]].copy()
                df_cidade_emp = df_cidade_emp.sort_values("cidade", ascending=True)
                fig = create_bar_chart(
                    df_cidade_emp, "cidade", "Quantidade",
                    f"Top {MAX_CIDADES_GRAFICO} cidades (Empresas) - {estado_filtro}",
                    max_items=MAX_CIDADES_GRAFICO, color="#3B82F6", sort_by_alpha=False
                )
                st.plotly_chart(fig, use_container_width=True)
                st.caption("Todas as cidades (Empresas)")
                st.dataframe(
                    df_cidade_emp.rename(columns={"cidade": "Cidade", "Quantidade": "Empresas"}),
                    use_container_width=True, hide_index=True, height=280
                )
            else:
                st.info("Sem dados de empresas para este estado.")
        st.markdown("---")
    create_section_header("📋", "Tabela Resumo por Estado")
    if not df_labs.empty and 'uf' in df_labs.columns:
        pcl_agg = df_labs.groupby('uf').agg({
            'cnpj': 'count',
            'status': lambda x: (x == 'Ativo').sum(),
            'acumulado_coletas': 'sum' if 'acumulado_coletas' in df_labs.columns else 'count',
            'cidade': 'nunique'
        }).reset_index()
        pcl_agg.columns = ['UF', 'PCLs', 'PCLs Ativos', 'Coletas', 'Cidades c/ PCL']
        pcl_agg['PCLs Inativos'] = pcl_agg['PCLs'] - pcl_agg['PCLs Ativos']
    else:
        pcl_agg = pd.DataFrame(columns=['UF', 'PCLs', 'PCLs Ativos', 'PCLs Inativos', 'Coletas', 'Cidades c/ PCL'])
    if not df_empresas.empty and 'uf' in df_empresas.columns:
        emp_agg = df_empresas.groupby('uf').agg({
            'cnpj': 'count',
            'status': lambda x: (x == 'Ativo').sum(),
            'acumulado_vouchers': 'sum' if 'acumulado_vouchers' in df_empresas.columns else 'count',
            'cidade': 'nunique'
        }).reset_index()
        emp_agg.columns = ['UF', 'Empresas', 'Empresas Ativas', 'Vouchers', 'Cidades c/ Empresa']
        emp_agg['Empresas Inativas'] = emp_agg['Empresas'] - emp_agg['Empresas Ativas']
    else:
        emp_agg = pd.DataFrame(columns=['UF', 'Empresas', 'Empresas Ativas', 'Empresas Inativas', 'Vouchers', 'Cidades c/ Empresa'])
    if not pcl_agg.empty and not emp_agg.empty:
        tabela_resumo = pd.merge(pcl_agg, emp_agg, on='UF', how='outer').fillna(0)
    elif not pcl_agg.empty:
        tabela_resumo = pcl_agg.copy()
        tabela_resumo['Empresas'] = 0
        tabela_resumo['Empresas Ativas'] = 0
        tabela_resumo['Empresas Inativas'] = 0
        tabela_resumo['Vouchers'] = 0
        tabela_resumo['Cidades c/ Empresa'] = 0
    elif not emp_agg.empty:
        tabela_resumo = emp_agg.copy()
        tabela_resumo['PCLs'] = 0
        tabela_resumo['PCLs Ativos'] = 0
        tabela_resumo['PCLs Inativos'] = 0
        tabela_resumo['Coletas'] = 0
        tabela_resumo['Cidades c/ PCL'] = 0
    else:
        tabela_resumo = pd.DataFrame()
    if not tabela_resumo.empty:
        colunas_ordem = ['UF', 'PCLs', 'PCLs Ativos', 'PCLs Inativos', 'Empresas', 'Empresas Ativas', 'Empresas Inativas', 'Coletas', 'Vouchers', 'Cidades c/ PCL', 'Cidades c/ Empresa']
        colunas_existentes = [c for c in colunas_ordem if c in tabela_resumo.columns]
        tabela_resumo = tabela_resumo[colunas_existentes]
        for col in tabela_resumo.columns:
            if col != 'UF':
                tabela_resumo[col] = pd.to_numeric(tabela_resumo[col], errors='coerce').fillna(0).astype(int)
        tabela_resumo = tabela_resumo.sort_values('UF', ascending=True)
        st.dataframe(tabela_resumo, use_container_width=True, hide_index=True, height=500)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            tabela_resumo.to_excel(writer, index=False, sheet_name='Resumo por UF')
        st.download_button(
            label="📥 Download Resumo por UF (Excel)",
            data=output.getvalue(),
            file_name=f'resumo_por_uf_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.warning("Nenhum dado disponível para exibição.")

# Callbacks para limpar filtros
def _limpar_filtro_coletas():
    st.session_state["filtro_estado_coletas"] = "Todos"
    st.session_state["filtro_cidade_coletas"] = "Todas"
    st.session_state["filtro_bairro_coletas"] = "Todos"
    st.toast("Filtros limpos!", icon="🔄")

def _limpar_filtro_pcls():
    st.session_state["filtro_estado_pcl"] = "Todos"
    st.session_state["filtro_cidade_pcl"] = "Todas"
    st.session_state["filtro_bairro_pcl"] = "Todos"
    st.toast("Filtros limpos!", icon="🔄")

def _limpar_filtro_empresas():
    st.session_state["filtro_estado_emp"] = "Todos"
    st.session_state["filtro_cidade_emp"] = "Todas"
    st.session_state["filtro_bairro_emp"] = "Todos"
    st.toast("Filtros limpos!", icon="🔄")

@st.fragment
def _analise_coletas_fragment():
    """Conteúdo da aba Coletas. Fragment para manter a aba ativa ao mudar filtros."""
    create_section_header("🔬", "Análise de Coletas", "Métricas detalhadas de coletas por estado e PCL")

    # Filtros
    col_filtro1, col_filtro2, col_filtro3, col_filtro4 = st.columns([1, 1, 1, 1])

    with col_filtro1:
        estados_coletas_lista = ['Todos'] + listas_filtros['pcl']['estados']
        estado_coletas_selecionado = st.selectbox("Estado", estados_coletas_lista, key="filtro_estado_coletas")

    with col_filtro2:
        if estado_coletas_selecionado != "Todos":
            cidades_coletas_lista = ['Todas'] + listas_filtros['pcl']['cidades_por_estado'].get(estado_coletas_selecionado, [])
        else:
            cidades_coletas_lista = ['Todas'] + listas_filtros['pcl']['cidades']
        cidade_coletas_selecionada = st.selectbox("Cidade", cidades_coletas_lista, key="filtro_cidade_coletas")

    with col_filtro3:
        df_temp_coletas = df_labs.copy()
        if estado_coletas_selecionado != "Todos":
            df_temp_coletas = df_temp_coletas[df_temp_coletas['uf'] == estado_coletas_selecionado]
        if cidade_coletas_selecionada != "Todas":
            df_temp_coletas = df_temp_coletas[df_temp_coletas['cidade'] == cidade_coletas_selecionada]
        bairros_coletas_lista = ['Todos'] + sorted(df_temp_coletas['bairro'].dropna().unique().tolist()) if 'bairro' in df_temp_coletas.columns else ['Todos']
        bairros_coletas_lista = [b for b in bairros_coletas_lista if b and str(b).strip()]
        bairro_coletas_selecionado = st.selectbox("Bairro", bairros_coletas_lista, key="filtro_bairro_coletas")

    with col_filtro4:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("Limpar Filtros", key="limpar_coletas", on_click=_limpar_filtro_coletas)

    # Aplicar filtros
    df_labs_coletas = apply_filters(df_labs, estado_coletas_selecionado, cidade_coletas_selecionada)
    if bairro_coletas_selecionado != "Todos" and 'bairro' in df_labs_coletas.columns:
        df_labs_coletas = df_labs_coletas[df_labs_coletas['bairro'] == bairro_coletas_selecionado]

    if df_labs_coletas.empty or 'acumulado_coletas' not in df_labs_coletas.columns:
        st.warning("Dados de coletas não disponíveis para os filtros selecionados.")
    else:
        # Métricas de coletas
        try:
            total_coletas = float(df_labs_coletas['acumulado_coletas'].sum()) if pd.notna(df_labs_coletas['acumulado_coletas'].sum()) else 0.0
            media_coletas = float(df_labs_coletas['acumulado_coletas'].mean()) if pd.notna(df_labs_coletas['acumulado_coletas'].mean()) else 0.0
            mediana_coletas = float(df_labs_coletas['acumulado_coletas'].median()) if pd.notna(df_labs_coletas['acumulado_coletas'].median()) else 0.0
            max_coletas = float(df_labs_coletas['acumulado_coletas'].max()) if pd.notna(df_labs_coletas['acumulado_coletas'].max()) else 0.0
            pcls_sem_coleta = len(df_labs_coletas[df_labs_coletas['acumulado_coletas'].fillna(0) == 0])
            pcls_com_coleta = len(df_labs_coletas[df_labs_coletas['acumulado_coletas'].fillna(0) > 0])
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
            pct_com_coleta = (pcls_com_coleta / len(df_labs_coletas) * 100) if len(df_labs_coletas) > 0 else 0
            create_metric_card("PCLs com Coleta", format_number(pcls_com_coleta), f"{pct_com_coleta:.1f}% do total", "", "gray")

        st.markdown("---")

        # Coletas por Estado
        create_section_header("📊", "Coletas por Estado")

        if 'uf' in df_labs_coletas.columns:
            coletas_estado = df_labs_coletas.groupby('uf').agg({
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
                    textposition='outside'
                ))

                fig.update_layout(
                    title=dict(text="Total de Coletas por Estado (Top 12)", font=dict(size=15), x=0.5),
                    xaxis=dict(title="", showgrid=True, gridcolor='rgba(128,128,128,0.2)'),
                    yaxis=dict(title=""),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
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
                    textposition='outside'
                ))

                fig.update_layout(
                    title=dict(text="Média de Coletas por PCL (Top 12)", font=dict(size=15), x=0.5),
                    xaxis=dict(title="", showgrid=True, gridcolor='rgba(128,128,128,0.2)'),
                    yaxis=dict(title=""),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
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

            cols_display = ['razao_social', 'representante', 'cidade', 'uf', 'acumulado_coletas', 'status']
            cols_available = [c for c in cols_display if c in df_labs_coletas.columns]

            if cols_available:
                top_pcls = df_labs_coletas.nlargest(20, 'acumulado_coletas')[cols_available]
                rename_map = {
                    'razao_social': 'Razão Social',
                    'representante': 'Representante',
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

@st.fragment
def _listagem_pcls_fragment():
    """Conteúdo da aba PCLs. Fragment para manter a aba ativa ao mudar filtros."""
    create_section_header("🏥", "Listagem de PCLs", "Base completa de laboratórios credenciados")

    # Filtros dentro da aba (usando listas cacheadas)
    col_filtro1, col_filtro2, col_filtro3, col_filtro4 = st.columns([1, 1, 1, 1])

    with col_filtro1:
        estados_pcl_lista = ['Todos'] + listas_filtros['pcl']['estados']
        estado_pcl_selecionado = st.selectbox("Estado", estados_pcl_lista, key="filtro_estado_pcl")

    with col_filtro2:
        if estado_pcl_selecionado != "Todos":
            cidades_pcl_lista = ['Todas'] + listas_filtros['pcl']['cidades_por_estado'].get(estado_pcl_selecionado, [])
        else:
            cidades_pcl_lista = ['Todas'] + listas_filtros['pcl']['cidades']
        cidade_pcl_selecionada = st.selectbox("Cidade", cidades_pcl_lista, key="filtro_cidade_pcl")

    with col_filtro3:
        # Filtrar bairros baseado na cidade/estado selecionado
        df_temp_pcl = df_labs.copy()
        if estado_pcl_selecionado != "Todos":
            df_temp_pcl = df_temp_pcl[df_temp_pcl['uf'] == estado_pcl_selecionado]
        if cidade_pcl_selecionada != "Todas":
            df_temp_pcl = df_temp_pcl[df_temp_pcl['cidade'] == cidade_pcl_selecionada]
        bairros_pcl_lista = ['Todos'] + sorted(df_temp_pcl['bairro'].dropna().unique().tolist()) if 'bairro' in df_temp_pcl.columns else ['Todos']
        bairros_pcl_lista = [b for b in bairros_pcl_lista if b and str(b).strip()]  # Remover vazios
        bairro_pcl_selecionado = st.selectbox("Bairro", bairros_pcl_lista, key="filtro_bairro_pcl")

    with col_filtro4:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("Limpar Filtros", key="limpar_pcls", on_click=_limpar_filtro_pcls)

    # Aplicar filtros
    df_labs_filtered = apply_filters(df_labs, estado_pcl_selecionado, cidade_pcl_selecionada)
    if bairro_pcl_selecionado != "Todos" and 'bairro' in df_labs_filtered.columns:
        df_labs_filtered = df_labs_filtered[df_labs_filtered['bairro'] == bairro_pcl_selecionado]

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

        # Usar dados pré-calculados (cacheados) em vez de recalcular
        df_display = df_labs_filtered.copy()

        # Adicionar contagens usando dicionários cacheados
        if 'cidade' in df_display.columns:
            df_display['qtd_empresas_cidade'] = df_display['cidade'].map(empresas_por_cidade).fillna(0).astype(int)
            df_display['qtd_empresas_ativas_cidade'] = df_display['cidade'].map(empresas_ativas_cidade).fillna(0).astype(int)
            df_display['qtd_empresas_inativas_cidade'] = df_display['qtd_empresas_cidade'] - df_display['qtd_empresas_ativas_cidade']
            df_display['qtd_empresas_usaram_voucher'] = df_display['cidade'].map(empresas_com_voucher).fillna(0).astype(int)

        # ADD-959: Adicionar dias desde última coleta para validação de status
        if 'data_ultima_coleta' in df_display.columns:
            hoje = pd.Timestamp.now().normalize()
            df_display['dias_sem_coleta'] = df_display['data_ultima_coleta'].apply(
                lambda x: (hoje - pd.to_datetime(x, errors='coerce')).days if pd.notna(x) else None
            )
        else:
            df_display['dias_sem_coleta'] = None

        # Preparar DataFrame para exibição
        colunas_pcl = [
            'cnpj', 'razao_social', 'nome_fantasia', 'endereco_logradouro', 'bairro', 'cidade', 'uf', 'cep',
            'data_credenciamento', 'representante',
            'acumulado_coletas', 'acumulado_coletas_ano', 'data_ultima_coleta', 'dias_sem_coleta', 'status',
            'transportadora', 'frequencia',
            'qtd_empresas_cidade'
        ]

        rename_map_pcl = {
            'cnpj': 'CNPJ', 'razao_social': 'Razão Social', 'nome_fantasia': 'Nome Fantasia',
            'endereco_logradouro': 'Endereço', 'bairro': 'Bairro',
            'cidade': 'Cidade', 'uf': 'UF', 'cep': 'CEP',
            'data_credenciamento': 'Data Credenciamento', 'representante': 'Representante',
            'acumulado_coletas': 'Coletas Total', 'acumulado_coletas_ano': 'Coletas 2025',
            'data_ultima_coleta': 'Última Coleta', 'dias_sem_coleta': 'Dias s/ Coleta', 'status': 'Status',
            'transportadora': 'Transportadora', 'frequencia': 'Frequência',
            'qtd_empresas_cidade': 'Empresas na Cidade'
        }

        df_final = prepare_display_dataframe(df_display, colunas_pcl, rename_map_pcl)

        if not df_final.empty:
            st.dataframe(df_final, use_container_width=True, hide_index=True, height=500)
        else:
            st.warning("Nenhum dado disponível para exibição.")

        # Download - ADD-956: incluir campos logísticos (transportadora, frequência)
        colunas_download_pcl = [
            'cnpj', 'razao_social', 'nome_fantasia', 'endereco_logradouro', 'bairro', 'cidade', 'uf', 'cep',
            'data_credenciamento', 'representante',
            'acumulado_coletas', 'acumulado_coletas_ano', 'data_ultima_coleta', 'dias_sem_coleta', 'status',
            'transportadora', 'frequencia',
            'qtd_empresas_cidade', 'qtd_empresas_ativas_cidade', 'qtd_empresas_inativas_cidade', 'qtd_empresas_usaram_voucher'
        ]

        rename_map_download_pcl = {
            'cnpj': 'CNPJ', 'razao_social': 'Razão Social', 'nome_fantasia': 'Nome Fantasia',
            'endereco_logradouro': 'Endereço', 'bairro': 'Bairro',
            'cidade': 'Cidade', 'uf': 'UF', 'cep': 'CEP',
            'data_credenciamento': 'Data Credenciamento', 'representante': 'Representante',
            'acumulado_coletas': 'Coletas Total', 'acumulado_coletas_ano': 'Coletas 2025',
            'data_ultima_coleta': 'Última Coleta', 'dias_sem_coleta': 'Dias s/ Coleta', 'status': 'Status',
            'transportadora': 'Transportadora', 'frequencia': 'Frequência',
            'qtd_empresas_cidade': 'Empresas na Cidade', 'qtd_empresas_ativas_cidade': 'Empresas Ativas',
            'qtd_empresas_inativas_cidade': 'Empresas Inativas', 'qtd_empresas_usaram_voucher': 'Empresas c/ Voucher'
        }

        # Selecionar apenas colunas existentes para download
        colunas_existentes = [c for c in colunas_download_pcl if c in df_display.columns]
        df_download = df_display[colunas_existentes].copy()
        df_download = df_download.rename(columns={k: v for k, v in rename_map_download_pcl.items() if k in df_download.columns})

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_download.to_excel(writer, index=False, sheet_name='PCLs')

        st.download_button(
            label="📥 Download Excel",
            data=output.getvalue(),
            file_name=f'pcls_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

@st.fragment
def _listagem_empresas_fragment():
    """Conteúdo da aba Empresas. Fragment para manter a aba ativa ao mudar filtros."""
    create_section_header("🏢", "Listagem de Empresas", "Base completa de empresas credenciadas")

    # Filtros dentro da aba
    col_filtro1, col_filtro2, col_filtro3, col_filtro4 = st.columns([1, 1, 1, 1])

    with col_filtro1:
        estados_emp_lista = ['Todos'] + listas_filtros['empresa']['estados']
        estado_emp_selecionado = st.selectbox("Estado", estados_emp_lista, key="filtro_estado_emp")

    with col_filtro2:
        if estado_emp_selecionado != "Todos":
            cidades_emp_lista = ['Todas'] + listas_filtros['empresa']['cidades_por_estado'].get(estado_emp_selecionado, [])
        else:
            cidades_emp_lista = ['Todas'] + listas_filtros['empresa']['cidades']
        cidade_emp_selecionada = st.selectbox("Cidade", cidades_emp_lista, key="filtro_cidade_emp")

    with col_filtro3:
        # Filtrar bairros baseado na cidade/estado selecionado
        df_temp = df_empresas.copy()
        if estado_emp_selecionado != "Todos":
            df_temp = df_temp[df_temp['uf'] == estado_emp_selecionado]
        if cidade_emp_selecionada != "Todas":
            df_temp = df_temp[df_temp['cidade'] == cidade_emp_selecionada]
        bairros_lista = ['Todos'] + sorted(df_temp['bairro'].dropna().unique().tolist()) if 'bairro' in df_temp.columns else ['Todos']
        bairros_lista = [b for b in bairros_lista if b and b.strip()]  # Remover vazios
        bairro_emp_selecionado = st.selectbox("Bairro", bairros_lista, key="filtro_bairro_emp")

    with col_filtro4:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("Limpar Filtros", key="limpar_empresas", on_click=_limpar_filtro_empresas)

    # Aplicar filtros
    df_empresas_filtered = apply_filters(df_empresas, estado_emp_selecionado, cidade_emp_selecionada)
    if bairro_emp_selecionado != "Todos" and 'bairro' in df_empresas_filtered.columns:
        df_empresas_filtered = df_empresas_filtered[df_empresas_filtered['bairro'] == bairro_emp_selecionado]

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

        # Usar dados pré-calculados (cacheados) em vez de recalcular
        df_display = df_empresas_filtered.copy()

        # Adicionar contagens usando dicionários cacheados
        if 'cidade' in df_display.columns:
            df_display['qtd_pcls_cidade'] = df_display['cidade'].map(pcls_por_cidade).fillna(0).astype(int)
            df_display['pcl_na_cidade'] = df_display['qtd_pcls_cidade'].apply(lambda x: 'Sim' if x > 0 else 'Não')
            df_display['qtd_pcls_ativos_cidade'] = df_display['cidade'].map(pcls_ativos_cidade).fillna(0).astype(int)
            df_display['qtd_pcls_inativos_cidade'] = df_display['qtd_pcls_cidade'] - df_display['qtd_pcls_ativos_cidade']

        # Garantir que colunas existam para exibição
        if 'acumulado_vouchers' not in df_display.columns:
            df_display['acumulado_vouchers'] = 0
        if 'acumulado_coletas_nao_voucher' not in df_display.columns:
            df_display['acumulado_coletas_nao_voucher'] = 0
        if 'acumulado_coletas_total' not in df_display.columns:
            df_display['acumulado_coletas_total'] = df_display['acumulado_vouchers'] + df_display['acumulado_coletas_nao_voucher']
        if 'coletas_2025' not in df_display.columns:
            df_display['coletas_2025'] = 0

        # ADD-958: Adicionar indicador se empresa já usou voucher
        df_display['ja_usou_voucher'] = df_display['acumulado_vouchers'].apply(lambda x: 'Sim' if x > 0 else 'Não')

        # ADD-959: Adicionar dias desde última coleta para validação de status
        if 'ultima_coleta' in df_display.columns:
            hoje = pd.Timestamp.now().normalize()
            df_display['dias_sem_coleta'] = df_display['ultima_coleta'].apply(
                lambda x: (hoje - pd.to_datetime(x, errors='coerce')).days if pd.notna(x) else None
            )
        else:
            df_display['dias_sem_coleta'] = None

        # Formatar valores para exibição
        for col in ['acumulado_vouchers', 'acumulado_coletas_nao_voucher', 'acumulado_coletas_total', 'coletas_2025']:
            if col in df_display.columns:
                df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0).astype(int)

        # Preparar DataFrame para exibição
        colunas_empresa = [
            'cnpj', 'razao_social', 'endereco_logradouro', 'bairro', 'cidade', 'uf', 'cep',
            'data_credenciamento', 'representante',
            'ja_usou_voucher', 'acumulado_coletas_total', 'acumulado_vouchers', 'acumulado_coletas_nao_voucher',
            'coletas_2025', 'ultima_coleta', 'ultima_coleta_voucher', 'dias_sem_coleta', 'status',
            'qtd_pcls_cidade'
        ]

        rename_map_empresa = {
            'cnpj': 'CNPJ', 'razao_social': 'Razão Social',
            'endereco_logradouro': 'Endereço', 'bairro': 'Bairro',
            'cidade': 'Cidade', 'uf': 'UF', 'cep': 'CEP',
            'data_credenciamento': 'Data Credenciamento', 'representante': 'Representante',
            'ja_usou_voucher': 'Já Usou Voucher', 'acumulado_coletas_total': 'Total Coletas',
            'acumulado_vouchers': 'Coletas Voucher', 'acumulado_coletas_nao_voucher': 'Coletas Não-Voucher',
            'coletas_2025': 'Coletas 2025', 'ultima_coleta': 'Última Coleta',
            'ultima_coleta_voucher': 'Último Voucher', 'dias_sem_coleta': 'Dias s/ Coleta', 'status': 'Status',
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

@st.fragment
def _analises_especificas_fragment():
    """Conteúdo da aba Análises Específicas. Fragment para manter a aba ativa ao mudar filtros."""
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
            # Normalizar cidades antes de comparar
            df_labs_norm = normalize_city_column(df_labs.copy().reset_index(drop=True), 'cidade')
            df_empresas_norm = normalize_city_column(df_empresas.copy().reset_index(drop=True), 'cidade')

            cidades_com_empresa = set(df_empresas_norm['cidade'].dropna().unique()) if 'cidade' in df_empresas_norm.columns else set()
            cidades_com_empresa = {c for c in cidades_com_empresa if c != ''}  # Remover strings vazias

            df_result_norm = df_labs_norm[~df_labs_norm['cidade'].isin(cidades_com_empresa)] if 'cidade' in df_labs_norm.columns else pd.DataFrame()
            # Usar o dataframe original para manter os dados originais
            if not df_result_norm.empty:
                df_result = df_labs.iloc[df_result_norm.index].copy()
            else:
                df_result = pd.DataFrame()

            if not df_result.empty:
                st.success(f"✅ Encontrados {len(df_result)} PCLs em cidades sem empresas credenciadas")

                cols = ['cnpj', 'razao_social', 'nome_fantasia', 'cidade', 'uf', 'transportadora', 'frequencia', 'status', 'acumulado_coletas', 'data_ultima_coleta']
                cols_available = [c for c in cols if c in df_result.columns]

                # Fallback: usar todas as colunas se nenhuma das esperadas existir
                if not cols_available:
                    cols_available = df_result.columns.tolist()

                df_display = df_result[cols_available].copy()

                rename_map = {'cnpj': 'CNPJ', 'razao_social': 'Razão Social', 'nome_fantasia': 'Nome Fantasia',
                              'cidade': 'Cidade', 'uf': 'UF', 'transportadora': 'Transportadora', 'frequencia': 'Frequência',
                              'status': 'Status', 'acumulado_coletas': 'Coletas', 'data_ultima_coleta': 'Última Coleta'}
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
            # Normalizar cidades antes de comparar
            df_labs_norm = normalize_city_column(df_labs.copy().reset_index(drop=True), 'cidade')
            df_empresas_norm = normalize_city_column(df_empresas.copy().reset_index(drop=True), 'cidade')

            # Encontrar cidades onde TODAS as empresas são inativas
            if 'cidade' in df_empresas_norm.columns and 'status' in df_empresas_norm.columns:
                empresas_por_cidade_analise = df_empresas_norm.groupby('cidade').agg({
                    'status': lambda x: (x == 'Ativo').sum()
                }).reset_index()
                empresas_por_cidade_analise.columns = ['cidade', 'empresas_ativas']

                # Cidades com empresas, mas nenhuma ativa
                cidades_empresas_inativas = set(empresas_por_cidade_analise[empresas_por_cidade_analise['empresas_ativas'] == 0]['cidade'])
                cidades_empresas_inativas = {c for c in cidades_empresas_inativas if c != ''}  # Remover strings vazias

                # PCLs nessas cidades
                df_result_norm = df_labs_norm[df_labs_norm['cidade'].isin(cidades_empresas_inativas)] if 'cidade' in df_labs_norm.columns else pd.DataFrame()
                # Usar o dataframe original para manter os dados originais
                if not df_result_norm.empty:
                    df_result = df_labs.iloc[df_result_norm.index].copy()
                else:
                    df_result = pd.DataFrame()

                if not df_result.empty:
                    st.warning(f"⚠️ Encontrados {len(df_result)} PCLs em cidades onde todas as empresas estão inativas")

                    cols = ['cnpj', 'razao_social', 'nome_fantasia', 'cidade', 'uf', 'transportadora', 'frequencia', 'status', 'acumulado_coletas', 'data_ultima_coleta']
                    cols_available = [c for c in cols if c in df_result.columns]

                    # Fallback: usar todas as colunas se nenhuma das esperadas existir
                    if not cols_available:
                        cols_available = df_result.columns.tolist()

                    df_display = df_result[cols_available].copy()

                    # Adicionar quantidade de empresas inativas na cidade
                    # Criar mapeamento de cidade normalizada para original
                    cidade_map_norm_to_orig = dict(zip(df_empresas_norm['cidade'], df_empresas['cidade']))
                    # Mapear cidades normalizadas de volta para originais
                    cidades_originais = {cidade_map_norm_to_orig.get(c, c) for c in cidades_empresas_inativas if c in cidade_map_norm_to_orig}
                    emp_count = df_empresas[df_empresas['cidade'].isin(cidades_originais)].groupby('cidade').size().to_dict()
                    df_display['empresas_inativas_cidade'] = df_display['cidade'].map(emp_count).fillna(0).astype(int)

                    rename_map = {'cnpj': 'CNPJ', 'razao_social': 'Razão Social', 'nome_fantasia': 'Nome Fantasia',
                                  'cidade': 'Cidade', 'uf': 'UF', 'transportadora': 'Transportadora', 'frequencia': 'Frequência',
                                  'status': 'Status PCL', 'acumulado_coletas': 'Coletas', 'data_ultima_coleta': 'Última Coleta',
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
            # Normalizar cidades antes de comparar
            df_labs_norm = normalize_city_column(df_labs.copy().reset_index(drop=True), 'cidade')
            df_empresas_norm = normalize_city_column(df_empresas.copy().reset_index(drop=True), 'cidade')

            cidades_com_pcl = set(df_labs_norm['cidade'].dropna().unique()) if 'cidade' in df_labs_norm.columns else set()
            cidades_com_pcl = {c for c in cidades_com_pcl if c != ''}  # Remover strings vazias

            df_result_norm = df_empresas_norm[~df_empresas_norm['cidade'].isin(cidades_com_pcl)] if 'cidade' in df_empresas_norm.columns else pd.DataFrame()
            # Usar o dataframe original para manter os dados originais
            if not df_result_norm.empty:
                df_result = df_empresas.iloc[df_result_norm.index].copy()
            else:
                df_result = pd.DataFrame()

            if not df_result.empty:
                st.error(f"❌ Encontradas {len(df_result)} empresas em cidades sem PCL credenciado")

                cols = ['cnpj', 'razao_social', 'nome_fantasia', 'cidade', 'uf', 'status', 'acumulado_vouchers', 'data_ultima_utilizacao']
                cols_available = [c for c in cols if c in df_result.columns]

                # Fallback: usar todas as colunas se nenhuma das esperadas existir
                if not cols_available:
                    cols_available = df_result.columns.tolist()

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
            # Normalizar cidades antes de comparar
            df_labs_norm = normalize_city_column(df_labs.copy().reset_index(drop=True), 'cidade')
            df_empresas_norm = normalize_city_column(df_empresas.copy().reset_index(drop=True), 'cidade')

            # Encontrar cidades onde TODOS os PCLs são inativos
            if 'cidade' in df_labs_norm.columns and 'status' in df_labs_norm.columns:
                pcls_por_cidade_analise = df_labs_norm.groupby('cidade').agg({
                    'status': lambda x: (x == 'Ativo').sum()
                }).reset_index()
                pcls_por_cidade_analise.columns = ['cidade', 'pcls_ativos']

                # Cidades com PCLs, mas nenhum ativo
                cidades_pcls_inativos = set(pcls_por_cidade_analise[pcls_por_cidade_analise['pcls_ativos'] == 0]['cidade'])
                cidades_pcls_inativos = {c for c in cidades_pcls_inativos if c != ''}  # Remover strings vazias

                # Empresas nessas cidades
                df_result_norm = df_empresas_norm[df_empresas_norm['cidade'].isin(cidades_pcls_inativos)] if 'cidade' in df_empresas_norm.columns else pd.DataFrame()
                # Usar o dataframe original para manter os dados originais
                if not df_result_norm.empty:
                    df_result = df_empresas.iloc[df_result_norm.index].copy()
                else:
                    df_result = pd.DataFrame()

                if not df_result.empty:
                    st.warning(f"⚠️ Encontradas {len(df_result)} empresas em cidades onde todos os PCLs estão inativos")

                    cols = ['cnpj', 'razao_social', 'nome_fantasia', 'cidade', 'uf', 'status', 'acumulado_vouchers', 'data_ultima_utilizacao']
                    cols_available = [c for c in cols if c in df_result.columns]

                    # Fallback: usar todas as colunas se nenhuma das esperadas existir
                    if not cols_available:
                        cols_available = df_result.columns.tolist()

                    df_display = df_result[cols_available].copy()

                    # Adicionar quantidade de PCLs inativos na cidade
                    # Criar mapeamento de cidade normalizada para original
                    cidade_map_norm_to_orig = dict(zip(df_labs_norm['cidade'], df_labs['cidade']))
                    # Mapear cidades normalizadas de volta para originais
                    cidades_originais = {cidade_map_norm_to_orig.get(c, c) for c in cidades_pcls_inativos if c in cidade_map_norm_to_orig}
                    pcl_count = df_labs[df_labs['cidade'].isin(cidades_originais)].groupby('cidade').size().to_dict()
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
            cols = ['cnpj', 'razao_social', 'representante', 'cidade', 'uf', 'transportadora', 'frequencia', 'acumulado_coletas', 'status']
            cols_available = [c for c in cols if c in top_pcls.columns]
            df_display = top_pcls[cols_available].copy()
            rename_map = {'cnpj': 'CNPJ', 'razao_social': 'Razão Social', 'representante': 'Representante',
                          'cidade': 'Cidade', 'uf': 'UF', 'transportadora': 'Transportadora', 'frequencia': 'Frequência',
                          'acumulado_coletas': 'Total Coletas', 'status': 'Status'}
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

@st.fragment
def _qualidade_dados_fragment():
    """ADD-961: Conteúdo da aba Qualidade de Dados. Fragment para manter a aba ativa ao mudar filtros."""
    create_section_header("🛡️", "Qualidade de Dados", "Identificação e consolidação de ofensores de cadastro")

    # Consolidar ofensores
    df_ofensores = consolidar_ofensores(df_labs, df_empresas)

    # Calcular métricas
    total_pcls = len(df_labs) if not df_labs.empty else 0
    total_empresas = len(df_empresas) if not df_empresas.empty else 0
    metricas = calcular_metricas_qualidade(df_ofensores, total_pcls, total_empresas)

    # Cards de métricas principais
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        # Score de qualidade com cor baseada no valor
        score = metricas['score_qualidade']
        if score >= 90:
            score_color = "green"
        elif score >= 70:
            score_color = "orange"
        else:
            score_color = "red"
        st.metric("Score de Qualidade", f"{score}%", delta=None)

    with col2:
        st.metric("🔴 Críticos", format_number(metricas['criticos']))

    with col3:
        st.metric("🟠 Altos", format_number(metricas['altos']))

    with col4:
        st.metric("🟡 Médios", format_number(metricas['medios']))

    with col5:
        st.metric("🟢 Baixos", format_number(metricas['baixos']))

    st.markdown("---")

    # Segunda linha de métricas
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Total de Problemas", format_number(metricas['total_ofensores']))

    with col2:
        st.metric("PCLs com Problema", format_number(metricas['pcls_com_problema']))

    with col3:
        st.metric("Empresas com Problema", format_number(metricas['empresas_com_problema']))

    with col4:
        total_registros = total_pcls + total_empresas
        registros_ok = total_registros - metricas['pcls_com_problema'] - metricas['empresas_com_problema']
        st.metric("Registros OK", format_number(registros_ok))

    st.markdown("---")

    # Filtros
    col_filtro1, col_filtro2, col_filtro3 = st.columns(3)

    with col_filtro1:
        filtro_entidade = st.selectbox(
            "Entidade",
            ["Todas", "PCL", "Empresa"],
            key="filtro_qualidade_entidade"
        )

    with col_filtro2:
        filtro_severidade = st.selectbox(
            "Severidade",
            ["Todas", "Crítico", "Alto", "Médio", "Baixo"],
            key="filtro_qualidade_severidade"
        )

    with col_filtro3:
        # Lista de tipos de problema únicos
        tipos_problema = ["Todos"] + sorted(df_ofensores['tipo_problema'].unique().tolist()) if not df_ofensores.empty else ["Todos"]
        filtro_tipo = st.selectbox(
            "Tipo de Problema",
            tipos_problema,
            key="filtro_qualidade_tipo"
        )

    # Aplicar filtros
    df_filtrado = df_ofensores.copy()

    if filtro_entidade != "Todas" and not df_filtrado.empty:
        df_filtrado = df_filtrado[df_filtrado['entidade'] == filtro_entidade]

    if filtro_severidade != "Todas" and not df_filtrado.empty:
        df_filtrado = df_filtrado[df_filtrado['severidade'] == filtro_severidade]

    if filtro_tipo != "Todos" and not df_filtrado.empty:
        df_filtrado = df_filtrado[df_filtrado['tipo_problema'] == filtro_tipo]

    st.markdown("---")

    # Tabs para diferentes visões
    subtab1, subtab2, subtab3 = st.tabs(["📋 Lista de Ofensores", "📊 Por Tipo de Problema", "👤 Por Representante"])

    with subtab1:
        if df_filtrado.empty:
            st.success("✅ Nenhum problema de qualidade encontrado com os filtros selecionados!")
        else:
            st.warning(f"⚠️ {len(df_filtrado)} problemas encontrados")

            # Preparar DataFrame para exibição
            colunas_exibir = ['entidade', 'cnpj', 'nome', 'cidade', 'uf', 'representante', 'campo_problema', 'tipo_problema', 'severidade']
            df_display = df_filtrado[colunas_exibir].copy()
            df_display.columns = ['Entidade', 'CNPJ', 'Nome', 'Cidade', 'UF', 'Representante', 'Campo', 'Problema', 'Severidade']

            # Adicionar ícone de severidade
            def add_severidade_icon(sev):
                icons = {'Crítico': '🔴', 'Alto': '🟠', 'Médio': '🟡', 'Baixo': '🟢'}
                return f"{icons.get(sev, '')} {sev}"

            df_display['Severidade'] = df_display['Severidade'].apply(add_severidade_icon)

            st.dataframe(df_display, use_container_width=True, hide_index=True, height=500)

    with subtab2:
        if not df_ofensores.empty:
            # Gráfico de problemas por tipo
            problemas_tipo = df_ofensores.groupby(['tipo_problema', 'severidade']).size().reset_index(name='quantidade')
            problemas_tipo = problemas_tipo.sort_values('quantidade', ascending=False).head(15)

            if not problemas_tipo.empty:
                fig = px.bar(
                    problemas_tipo,
                    x='quantidade',
                    y='tipo_problema',
                    color='severidade',
                    orientation='h',
                    title='Top 15 Tipos de Problemas',
                    color_discrete_map={'Crítico': '#EF4444', 'Alto': '#F97316', 'Médio': '#EAB308', 'Baixo': '#22C55E'}
                )
                fig.update_layout(yaxis={'categoryorder': 'total ascending'}, height=500)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Nenhum dado para exibir.")

    with subtab3:
        if metricas['problemas_por_representante']:
            # Gráfico de problemas por representante
            df_rep = pd.DataFrame([
                {'representante': k, 'quantidade': v}
                for k, v in metricas['problemas_por_representante'].items()
            ]).sort_values('quantidade', ascending=False).head(15)

            if not df_rep.empty:
                fig = px.bar(
                    df_rep,
                    x='quantidade',
                    y='representante',
                    orientation='h',
                    title='Top 15 Representantes com Mais Problemas',
                    color='quantidade',
                    color_continuous_scale='Reds'
                )
                fig.update_layout(yaxis={'categoryorder': 'total ascending'}, height=500, showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Nenhum problema associado a representantes.")

    st.markdown("---")

    # Download
    if not df_ofensores.empty:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Aba Resumo
            resumo_data = {
                'Métrica': [
                    'Score de Qualidade (%)',
                    'Total de Problemas',
                    'Problemas Críticos',
                    'Problemas Altos',
                    'Problemas Médios',
                    'Problemas Baixos',
                    'PCLs com Problema',
                    'Empresas com Problema',
                    'Total PCLs',
                    'Total Empresas'
                ],
                'Valor': [
                    metricas['score_qualidade'],
                    metricas['total_ofensores'],
                    metricas['criticos'],
                    metricas['altos'],
                    metricas['medios'],
                    metricas['baixos'],
                    metricas['pcls_com_problema'],
                    metricas['empresas_com_problema'],
                    total_pcls,
                    total_empresas
                ]
            }
            pd.DataFrame(resumo_data).to_excel(writer, index=False, sheet_name='Resumo')

            # Aba Ofensores PCLs
            df_pcls_ofensores = df_ofensores[df_ofensores['entidade'] == 'PCL'].copy()
            if not df_pcls_ofensores.empty:
                df_pcls_ofensores = df_pcls_ofensores.drop(columns=['severidade_ordem'])
                df_pcls_ofensores.columns = ['Entidade', 'CNPJ', 'Nome', 'Cidade', 'UF', 'Representante', 'Campo', 'Problema', 'Valor Atual', 'Severidade']
                df_pcls_ofensores.to_excel(writer, index=False, sheet_name='Ofensores PCLs')

            # Aba Ofensores Empresas
            df_emp_ofensores = df_ofensores[df_ofensores['entidade'] == 'Empresa'].copy()
            if not df_emp_ofensores.empty:
                df_emp_ofensores = df_emp_ofensores.drop(columns=['severidade_ordem'])
                df_emp_ofensores.columns = ['Entidade', 'CNPJ', 'Nome', 'Cidade', 'UF', 'Representante', 'Campo', 'Problema', 'Valor Atual', 'Severidade']
                df_emp_ofensores.to_excel(writer, index=False, sheet_name='Ofensores Empresas')

            # Aba Por Tipo de Problema
            problemas_tipo_resumo = df_ofensores.groupby(['tipo_problema', 'severidade']).size().reset_index(name='quantidade')
            problemas_tipo_resumo = problemas_tipo_resumo.sort_values(['severidade', 'quantidade'], ascending=[True, False])
            problemas_tipo_resumo.columns = ['Tipo de Problema', 'Severidade', 'Quantidade']
            problemas_tipo_resumo.to_excel(writer, index=False, sheet_name='Por Tipo de Problema')

        st.download_button(
            label="📥 Download Relatório de Qualidade (Excel)",
            data=output.getvalue(),
            file_name=f'qualidade_dados_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.success("✅ Nenhum problema de qualidade encontrado! Parabéns pela excelência dos dados.")

# Abas de navegação (visual original). A aba "Por Estado" usa fragment para manter a aba ativa ao mudar o selectbox.
tab_visao_geral, tab_visao_estado, tab_analise_coletas, tab_listagem_pcls, tab_listagem_empresas, tab_analises_especificas, tab_qualidade, tab_ajuda = st.tabs([
    "📈 Visão Geral",
    "🗺️ Por Estado",
    "📊 Coletas",
    "🏥 PCLs",
    "🏢 Empresas",
    "🔍 Análises",
    "🛡️ Qualidade",
    "❓ Ajuda"
])

with tab_visao_geral:
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
        media_coletas_por_pcl = total_coletas / total_pcls if total_pcls > 0 else 0
        create_metric_card("Média de Coletas", format_number(int(media_coletas_por_pcl)), "Coletas por PCL", "", "gray")
    
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

with tab_visao_estado:
    _visao_estado_fragment()

with tab_analise_coletas:
    _analise_coletas_fragment()

with tab_listagem_pcls:
    _listagem_pcls_fragment()

with tab_listagem_empresas:
    _listagem_empresas_fragment()

with tab_analises_especificas:
    _analises_especificas_fragment()

with tab_qualidade:
    _qualidade_dados_fragment()

with tab_ajuda:
    create_section_header("❓", "Ajuda / FAQ", "Perguntas frequentes e orientações de uso")

    # Sub-tabs para organizar o conteúdo de ajuda
    subtab1, subtab2, subtab3, subtab4, subtab5, subtab6 = st.tabs(["Navegação", "Entendendo os Dados", "Colunas e Métricas", "Análises", "Problemas Comuns", "Fontes de Dados"])

    with subtab1:
        st.markdown("### Como navegar no dashboard?")
        st.markdown("""
        Use as **abas na parte superior** da página para acessar as diferentes seções:

        | Aba | Descrição |
        |-----|-----------|
        | **📈 Visão Geral** | Métricas gerais e gráficos comparativos por estado |
        | **🗺️ Por Estado** | Análise detalhada por UF |
        | **📊 Coletas** | Estatísticas de coletas realizadas |
        | **🏥 PCLs** | Tabela completa de laboratórios/pontos de coleta |
        | **🏢 Empresas** | Tabela completa de empresas credenciadas |
        | **🔍 Análises** | Consultas customizadas para cenários específicos |
        | **🛡️ Qualidade** | Identificação de problemas de cadastro (ADD-961) |
        | **❓ Ajuda** | Esta seção de ajuda |
        """)

        st.markdown("### Como filtrar os dados?")
        st.markdown("""
        Os filtros **ESTADO** e **CIDADE** estão disponíveis diretamente nas abas de **PCLs** e **Empresas**.
        - Selecione "Todos" para ver todos os registros
        - Os filtros afetam os dados exibidos e também os downloads em Excel
        """)

    with subtab2:
        st.markdown("### O que é um PCL?")
        st.markdown("""
        **PCL** (Ponto de Coleta/Laboratório) é um estabelecimento credenciado para realizar coletas de exames toxicológicos.
        """)

        st.markdown("### Qual a diferença entre PCL Ativo e Inativo?")
        st.warning("""
        ⚠️ **IMPORTANTE:** PCLs **descredenciados** são excluídos completamente do sistema e não aparecem em nenhuma análise.
        
        **O que são PCLs descredenciados?**
        - São PCLs que tiveram seu credenciamento revogado ou cancelado
        - Na planilha Excel, esses PCLs aparecem com o campo técnico **"Ativo em Coletas" = False**
        - Este é um **dado técnico** da planilha que identifica PCLs que não devem mais ser considerados no sistema
        - Esses PCLs são automaticamente filtrados e não entram em nenhuma contagem, análise ou listagem
        """)
        col1, col2 = st.columns(2)
        with col1:
            st.success("**ATIVO**: Realizou coleta nos últimos **90 dias**")
        with col2:
            st.error("**INATIVO**: Sem coletas há mais de **90 dias**")
        
        st.markdown("""
        **Critérios de classificação:**
        - **PCLs descredenciados** (campo técnico "Ativo em Coletas" = False) são **excluídos** antes de qualquer análise
        - Para os PCLs restantes, se a coluna "Ativo em Coletas" existir no Excel, apenas PCLs com valor **True** são considerados
        - Se a coluna "Ativo em Coletas" não existir, usa-se "Dias sem coleta": ≤ 90 dias = Ativo, > 90 dias = Inativo
        """)

        st.markdown("### Qual a diferença entre Empresa Ativa e Inativa?")
        col1, col2 = st.columns(2)
        with col1:
            st.success("**ATIVA**: Utilizou voucher nos últimos **365 dias**")
        with col2:
            st.error("**INATIVA**: Sem atividade há mais de **365 dias**")

        st.markdown("### O que são Vouchers?")
        st.markdown("""
        Vouchers são créditos que empresas utilizam para pagar coletas de exames toxicológicos de seus funcionários.
        """)

        st.markdown("### Os dados são atualizados em tempo real?")
        st.markdown("""
        Não. Os dados são atualizados de acordo com a disponibilidade nas planilhas no SharePoint corporativo. 
        As planilhas estão localizadas no **SharePoint corporativo** na pasta **"Data Analysis"**, dentro das subpastas:
        - **"Acumulado de Coletas - Empresas"** (para dados de empresas)
        - **"Acumulado de Coletas - Labs"** (para dados de PCLs)
        
        O sistema carrega automaticamente o arquivo mais recente de cada pasta sempre que a página é recarregada.
        """)

    with subtab3:
        st.markdown("### Métricas da Visão Geral")
        st.markdown("""
        A aba **Visão Geral** exibe as seguintes métricas principais:
        
        | Métrica | Descrição |
        |---------|-----------|
        | **Total PCLs** | Quantidade total de PCLs no sistema (excluindo PCLs descredenciados - campo técnico "Ativo em Coletas" = False) |
        | **PCLs Inativos** | PCLs sem coleta há mais de 90 dias |
        | **Total Empresas** | Quantidade total de empresas credenciadas |
        | **Empresas Inativas** | Empresas sem uso de voucher há mais de 365 dias |
        | **Total Coletas** | Soma de todas as coletas realizadas (histórico) |
        | **Total Vouchers** | Soma de todos os vouchers utilizados (histórico) |
        | **Média de Coletas** | Média de coletas por PCL (Total Coletas ÷ Total PCLs) |
        | **Cobertura** | Quantidade de estados e cidades atendidas |
        """)
        
        st.markdown("### Colunas da Listagem de PCLs")
        st.markdown("""
        | Coluna | Descrição |
        |--------|-----------|
        | **CNPJ** | Número de identificação fiscal do laboratório |
        | **Razão Social** | Nome oficial registrado |
        | **Nome Fantasia** | Nome comercial do estabelecimento |
        | **Cidade / UF** | Localização do PCL |
        | **Data Credenciamento** | Data em que o PCL foi credenciado |
        | **Representante** | Representante comercial responsável |
        | **Transportadora** | Empresas de transporte disponíveis na cidade |
        | **Frequência** | Frequência de coleta (DIARIO, SEMANAL, etc.) |
        | **Coletas Total** | Total histórico de coletas realizadas |
        | **Coletas 2025** | Coletas realizadas no ano de 2025 |
        | **Última Coleta** | Data da última coleta |
        | **Status** | Ativo ou Inativo |
        | **Empresas na Cidade** | Qtd. de empresas na mesma cidade |
        """)

        st.markdown("### O que significa quando Transportadora mostra múltiplos valores?")
        st.info("Quando uma cidade possui mais de uma transportadora, elas são separadas por \" | \". Exemplo: **AIRLAB | BIOMED LOG | CORREIOS AP**")

        st.markdown("### Valores de Frequência")
        st.markdown("""
        | Valor | Significado |
        |-------|-------------|
        | **DIARIO** | Coleta todos os dias úteis |
        | **SEMANAL** | Coleta uma vez por semana |
        | **2ª, 4ª E 6ª** | Coleta segundas, quartas e sextas |
        | **3ª E 5ª** | Coleta terças e quintas |
        | **ALTERNADO** | Coleta em dias alternados |
        """)

    with subtab4:
        st.markdown("### Análises Específicas Disponíveis")

        with st.expander("1. PCLs em cidades SEM Empresas credenciadas"):
            st.markdown("""
            Lista laboratórios em cidades onde **não há nenhuma empresa cliente**.

            **Utilidade:** Identificar PCLs que podem precisar de prospecção comercial na região.
            """)

        with st.expander("2. PCLs em cidades COM Empresas INATIVAS (365 dias)"):
            st.markdown("""
            Lista laboratórios em cidades onde **todas as empresas estão inativas** (sem vouchers há mais de 365 dias).

            **Utilidade:** Indica oportunidades de reativação de clientes.
            """)

        with st.expander("3. Empresas em cidades SEM PCL credenciado"):
            st.markdown("""
            Lista empresas que **não têm laboratório disponível** em sua cidade.

            **Utilidade:** Indica necessidade de credenciar novos PCLs para atender essas empresas.
            """)

        with st.expander("4. Empresas em cidades COM PCL INATIVO (90 dias)"):
            st.markdown("""
            Lista empresas em cidades onde **todos os PCLs estão inativos** (sem coletas há mais de 90 dias).

            **Utilidade:** Indica risco de perda de clientes por falta de atendimento ativo.
            """)

        with st.expander("5. Top PCLs por volume de coletas"):
            st.markdown("""
            Ranking dos **50 PCLs com maior volume** de coletas realizadas.

            **Utilidade:** Identificar os principais parceiros e laboratórios mais ativos.
            """)

        with st.expander("6. Estados com menor cobertura"):
            st.markdown("""
            Lista estados ordenados por **quantidade de cidades atendidas** (menor para maior).

            **Utilidade:** Identificar oportunidades de expansão territorial.
            """)

    with subtab5:
        st.markdown("### Problemas Comuns e Soluções")

        with st.expander("O dashboard está lento"):
            st.markdown("""
            - Verifique sua conexão com a internet
            - Recarregue a página (F5)
            - Limpe o cache do navegador se o problema persistir
            """)

        with st.expander("Os dados parecem desatualizados"):
            st.markdown("""
            1. Verifique a data de atualização no rodapé
            2. Pressione F5 para forçar nova leitura dos dados
            3. Verifique se os arquivos fonte foram atualizados
            """)

        with st.expander("Erro: Arquivo está aberto em outro programa"):
            st.markdown("""
            Isso ocorre quando o arquivo Excel está aberto em outro programa.

            **Solução:** Feche o Excel e recarregue a página.
            """)

        with st.expander("Transportadora/Frequência aparecem vazias"):
            st.markdown("""
            Possíveis causas:
            - O arquivo `CONSULTA MATRIZ LOGISTICA.1.xlsx` não está presente
            - A cidade do PCL não está cadastrada na matriz logística
            - O nome da cidade pode estar escrito de forma diferente
            """)
        
        with st.expander("Alguns PCLs não aparecem no sistema"):
            st.markdown("""
            **Causa:** PCLs **descredenciados** são excluídos completamente do sistema.
            
            **O que são PCLs descredenciados?**
            - São PCLs que tiveram seu credenciamento revogado ou cancelado
            - Na planilha Excel, esses PCLs são identificados pelo campo técnico **"Ativo em Coletas" = False**
            - Este é um **dado técnico** da planilha que sinaliza que o PCL não deve mais ser considerado
            
            **Como funciona:**
            - Se o arquivo Excel contém a coluna "Ativo em Coletas", o sistema verifica este campo técnico
            - Apenas PCLs com valor **True** são processados e aparecem nas análises
            - PCLs com valor **False** (descredenciados) são **automaticamente excluídos** antes de qualquer processamento
            - Esses PCLs não aparecem em nenhuma análise, listagem, métrica ou download
            
            **Por que isso acontece?**
            - É uma regra de negócio: PCLs descredenciados não devem ser considerados em análises operacionais
            - O campo "Ativo em Coletas" = False é o indicador técnico usado para identificar esses PCLs na planilha
            
            **Solução:** Se um PCL descredenciado precisa aparecer novamente, é necessário atualizar o campo "Ativo em Coletas" para **True** na planilha fonte no SharePoint.
            """)

        with st.expander("Uma cidade não aparece nos resultados"):
            st.markdown("""
            - Verifique se o nome está escrito corretamente
            - Cidades podem ter nomes diferentes (ex: "SAO PAULO" vs "SÃO PAULO")
            - O sistema normaliza os nomes, mas diferenças significativas podem impedir o match
            """)

        st.markdown("---")
        st.markdown("### Glossário")
        st.markdown("""
        | Termo | Definição |
        |-------|-----------|
        | **PCL** | Ponto de Coleta/Laboratório credenciado |
        | **Voucher** | Crédito para pagamento de coletas |
        | **Credenciamento** | Cadastro e autorização no sistema |
        | **Status Ativo** | Entidade com atividade recente |
        | **Status Inativo** | Entidade sem atividade por período prolongado |
        """)

    with subtab6:
        st.markdown("### Fontes de Dados")
        st.markdown("""
        O dashboard carrega dados de diferentes fontes para compor as análises.
        """)

        st.markdown("#### Dados de PCLs e Empresas")

        # Mostrar arquivos carregados
        col_pcl, col_emp = st.columns(2)

        with col_pcl:
            st.markdown("**PCLs (Laboratórios)**")
            if file_info.get('labs_file'):
                icon = "☁️" if file_info.get('labs_source') == 'sharepoint' else "💻"
                st.success(f"{icon} {file_info['labs_file']}")
            else:
                st.warning("Nenhum arquivo carregado")

        with col_emp:
            st.markdown("**Empresas**")
            if file_info.get('empresas_file'):
                icon = "☁️" if file_info.get('empresas_source') == 'sharepoint' else "💻"
                st.success(f"{icon} {file_info['empresas_file']}")
            else:
                st.warning("Nenhum arquivo carregado")

        st.markdown("#### Matriz Logística (Transportadora/Frequência)")
        # Exibir status do carregamento da matriz logística
        if matriz_status.get('loaded'):
            # Buscar nome do arquivo das configurações
            sp_logistica = st.secrets.get("sharepoint_logistica", {})
            matriz_file_path = sp_logistica.get("file_path", "")
            matriz_file_name = matriz_file_path.split("/")[-1] if matriz_file_path else "CONSULTA MATRIZ LOGISTICA.1.xlsx"
            st.success(f"☁️ {matriz_file_name}")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Registros:** {matriz_status.get('records', 0):,}")
            with col2:
                loaded_at = matriz_status.get('loaded_at')
                if loaded_at:
                    st.markdown(f"**Carregada em:** {loaded_at.strftime('%d/%m/%Y %H:%M')}")
        else:
            st.error("Não foi possível carregar a matriz logística")
            if matriz_status.get('error'):
                st.markdown(f"**Erro:** {matriz_status.get('error')}")
            st.markdown("""
            **Possíveis causas:**
            - Arquivo não encontrado no SharePoint
            - Problemas de conexão
            - Credenciais inválidas ou expiradas
            """)

        st.caption("☁️ = SharePoint/OneDrive | 💻 = Arquivo local")

        st.markdown("---")
        st.markdown("#### Atualização dos dados")
        st.markdown("""
        Os dados são atualizados de acordo com a disponibilidade nas planilhas no SharePoint corporativo. 
        As planilhas estão localizadas no **SharePoint corporativo** na pasta **"Data Analysis"**, dentro das subpastas:
        - **"Acumulado de Coletas - Empresas"** (para dados de empresas)
        - **"Acumulado de Coletas - Labs"** (para dados de PCLs)
        
        Para forçar uma nova leitura:
        1. Pressione **F5** para recarregar a página
        2. O cache é automaticamente invalidado após 1 hora
        """)

# ============================================
# RODAPÉ
# ============================================

st.markdown("---")
st.caption(f"📅 Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M')} | CTOX Analytics v2.0")
