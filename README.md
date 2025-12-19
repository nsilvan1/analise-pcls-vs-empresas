# 📊 Painel de Análise CTOX - PCLs vs Empresas

Aplicação Streamlit para análise da base CTOX, incluindo listagens de PCLs e Empresas com gráficos comparativos.

## 📋 Requisitos

- Python 3.8 ou superior
- Arquivos Excel nas pastas:
  - `Acumulado de Coletas - Empresas/`
  - `Acumulado de Coletas - Labs/`

## 🚀 Instalação

1. Instale as dependências:
```bash
pip install -r requirements.txt
```

## ▶️ Execução

Execute a aplicação com:
```bash
streamlit run app.py
```

A aplicação será aberta automaticamente no navegador em `http://localhost:8501`

## 📁 Estrutura de Pastas

```
.
├── app.py                                    # Aplicação principal
├── requirements.txt                          # Dependências
├── README.md                                 # Este arquivo
├── Acumulado de Coletas - Empresas/         # Arquivos Excel de empresas
│   └── empresas_data_*.xlsx
└── Acumulado de Coletas - Labs/              # Arquivos Excel de labs (PCLs)
    └── laboratories_data_*.xlsx
```

## 🎯 Funcionalidades

### 1. Visão Geral
- Métricas gerais (Total de PCLs, PCLs Ativos, Total de Empresas, Empresas Ativas)
- Gráficos por estado:
  - Quantidade de PCLs cadastrados
  - PCLs ativos vs inativos
  - Quantidade de empresas cadastradas
  - Empresas ativas vs inativas

### 2. Listagem de PCLs
Exibe todos os campos solicitados:
- CNPJ
- Razão Social
- Nome Fantasia
- Data Credenciamento
- Aceita Voucher (sim/não)
- Valor do Exame
- Acumulado de Coletas
- Acumulado de Coletas neste Ano
- Ativo/Inativo
- Data da Última Coleta
- Cidade
- UF
- Representante
- Quantidade de empresas na cidade do PCL
- Quantidade de empresas que utilizaram voucher na cidade do PCL
- Quantidade de empresas que nunca utilizaram voucher na cidade

### 3. Listagem de Empresas
Exibe todos os campos solicitados:
- CNPJ
- Razão Social
- Nome Fantasia
- Data Credenciamento
- Valor Negociado (preço exclusivo)
- Acumulado de Vouchers
- Acumulado de Vouchers neste Ano
- Ativo/Inativo
- Data da Última Utilização de Voucher
- Cidade
- UF
- Representante
- Quantidade de PCLs credenciados na cidade da empresa
- Quantidade de PCLs ativos na cidade da empresa
- Quantidade de PCLs inativos na cidade

### 4. Análises Específicas
Implementa as 4 regras de análise:
1. **PCLs sem Empresas credenciadas na cidade**: Lista PCLs em determinada cidade que não têm empresas credenciadas
2. **PCLs com Empresas credenciadas inativas (365 dias)**: Lista PCLs em determinada cidade que têm empresas credenciadas, mas todas estão inativas
3. **Empresas sem PCL credenciado na cidade**: Lista empresas em determinada cidade que não têm PCL credenciado
4. **Empresas com PCL credenciado inativo (90 dias)**: Lista empresas em determinada cidade que têm PCL credenciado, mas ele não está ativo

## 📊 Regras de Negócio

- **PCL ativo/inativo**: 90 dias (≤90 dias sem coleta é ativo, >91 dias é inativo)
- **Empresa ativa/inativa**: 365 dias (≤365 dias sem utilização de voucher é ativo, >366 dias é inativo)
- **Atualização mensal**: Último dia do mês
- **Não separar por representação interna/externa**

## 📥 Download

Todas as listagens podem ser baixadas em formato Excel através dos botões de download disponíveis em cada seção.

## 🔍 Filtros

A aplicação permite filtrar por:
- Estado (UF)
- Cidade

## 📝 Notas

- A aplicação lê automaticamente todos os arquivos `.xlsx` das pastas especificadas
- Os dados são combinados automaticamente quando há múltiplos arquivos
- A aplicação tenta normalizar automaticamente os nomes das colunas para diferentes variações

