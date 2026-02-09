# Changelog - Implementacao das Novas Colunas de PCLs

**Data:** 2026-02-09
**Arquivo modificado:** `app.py`

---

## Resumo Geral

O arquivo Excel de PCLs foi retroalimentado com 29 novas colunas. A dashboard foi atualizada para mapear, processar e exibir esses novos dados em todas as abas relevantes.

---

## Fase 1: Mapeamento de Colunas

### ANTES
A funcao `normalize_column_names()` mapeava apenas 21 colunas do arquivo de PCLs:
- Identificacao: cnpj, razao_social, nome_fantasia
- Datas: data_credenciamento, data_ultima_coleta
- Coletas: acumulado_coletas, coletas_2024, coletas_2025
- Localizacao: endereco, cidade, uf, cep, representante
- Dias: dias_sem_coleta, dias_sem_coleta_voucher, dias_sem_coleta_nao_voucher

### DEPOIS
Adicionadas 29 novas colunas ao mapeamento (com variantes acentuadas e nao-acentuadas):

| Categoria | Colunas Adicionadas |
|-----------|-------------------|
| Contato | email, email_financeiro, telefone |
| Temporal | ano_mes_credenciamento, ano_mes_ultima_coleta |
| Historico | coletas_2023, dias_desde_credenciamento |
| Capacidade | volume_maximo_coletas_2023, volume_maximo_coletas_2024, volume_maximo_coletas_2025 |
| Comercial | status_crm, status_voucher |
| Finalidade | finalidade_cnh, finalidade_clt, finalidade_outro, finalidade_concurso_publico |
| Finalidade Detalhado | cnh_fe, cnh_fc, clt_fe, clt_fc |
| Financeiro | valor_total_coleta |
| Fiscal | simples_nacional, cnpj_paulinia |
| Formas de Pagamento | pag_dinheiro, pag_credito, pag_debito, pag_boleto, pag_faturamento, pag_faturamento_empresa, pag_faturamento_laboratorio, pag_credito_online |
| Dados Mensais | coletas_nov_2025, coletas_dez_2025, coletas_jan_2026 |

---

## Fase 2: Processamento de Dados

### ANTES
Apenas `acumulado_coletas` recebia conversao numerica em `process_labs()`:
```python
if 'acumulado_coletas' in df.columns:
    df['acumulado_coletas'] = pd.to_numeric(df['acumulado_coletas'], errors='coerce').fillna(0)
```

### DEPOIS
17 novas colunas numericas recebem conversao automatica:
- coletas_2023
- volume_maximo_coletas_2023, 2024, 2025
- finalidade_cnh, finalidade_clt, finalidade_outro, finalidade_concurso_publico
- cnh_fe, cnh_fc, clt_fe, clt_fc
- valor_total_coleta
- coletas_nov_2025, coletas_dez_2025, coletas_jan_2026
- dias_desde_credenciamento

---

## Fase 3: Aba PCLs - Seletor de Colunas

### ANTES
A tabela de PCLs exibia 18 colunas fixas:
- CNPJ, Razao Social, Nome Fantasia, Endereco, Bairro, Cidade, UF, CEP
- Data Credenciamento, Representante
- Coletas Total, Coletas 2025, Ultima Coleta, Dias s/ Coleta, Status
- Transportadora, Frequencia, Empresas na Cidade

O download Excel tambem incluia apenas essas colunas + Empresas Ativas/Inativas/c/ Voucher.

### DEPOIS
- **Novo seletor `st.multiselect`** acima da tabela com 8 grupos de colunas extras:
  - Contato (3 colunas)
  - Comercial (2 colunas)
  - Capacidade (3 colunas)
  - Finalidade (4 colunas)
  - Finalidade Detalhado (4 colunas)
  - Financeiro/Fiscal (3 colunas)
  - Formas de Pagamento (8 colunas)
  - Historico/Mensal (4 colunas)
- Tabela expande dinamicamente com as colunas selecionadas
- **Download Excel inclui TODAS as colunas** (originais + 29 novas), independente do seletor

---

## Fase 4: Visao Geral - Perfil Comercial

### ANTES
A Visao Geral exibia apenas:
- Linha 1: Total PCLs, PCLs Inativos, Total Empresas, Empresas Inativas
- Linha 2: Total Coletas, Total Vouchers, Cobertura, Media de Coletas
- Secao Status: PCLs Ativos, Empresas Ativas, Top 5 UFs PCLs, Top 5 UFs Empresas
- Graficos por Estado (barras simples + agrupadas)

### DEPOIS
Nova secao **"Perfil Comercial"** adicionada entre Volumes e Status:
- **Aceita Voucher** - Progress card (quantidade e % de PCLs que aceitam voucher)
- **Simples Nacional** - Progress card (quantidade e % de PCLs no Simples Nacional)
- **Valor Medio Coleta** - Metric card em formato R$ (media do valor total da coleta)
- **Capacidade 2025** - Metric card (soma do volume maximo de coletas 2025)
- **Grafico "Finalidade das Coletas"** - Barras horizontais com CNH, CLT, Outro, Concurso

---

## Fase 5: Aba Coletas - Tendencia e Capacidade

### ANTES
A aba Coletas exibia:
- Metricas: Total de Coletas, Media por PCL, Maximo, PCLs com Coleta
- Graficos: Coletas por Estado (total e media)

### DEPOIS
Duas novas secoes entre as metricas e os graficos por estado:

**Secao "Tendencia Mensal":**
- Coletas 2023, Nov/2025, Dez/2025, Jan/2026 (se disponiveis no arquivo)
- Permite ver evolucao meses a mes

**Secao "Capacidade vs Realizado 2025":**
- Capacidade Total (volume maximo 2025)
- Realizado (coletas 2025)
- Utilizacao (progress card com %)

---

## Fase 6: Aba Analises - 3 Novas Analises

### ANTES
7 analises disponiveis:
1. PCLs em cidades SEM Empresas credenciadas
2. PCLs em cidades COM Empresas INATIVAS (365 dias)
3. Empresas em cidades SEM PCL credenciado
4. Empresas em cidades COM PCL INATIVO (90 dias)
5. Top PCLs por volume de coletas
6. Estados com menor cobertura
7. Acompanhamento por Representante

### DEPOIS
10 analises disponiveis (+3 novas):

8. **PCLs sem coleta recente (3 meses)**
   - Filtra PCLs com Nov/2025 = 0, Dez/2025 = 0, Jan/2026 = 0
   - Exibe dados de contato (e-mail, telefone) para follow-up
   - Tabela + download Excel

9. **PCLs com capacidade ociosa**
   - Filtra PCLs onde coletas_2025 < 50% do volume_maximo_coletas_2025
   - Exibe % de utilizacao para priorizar acoes
   - Tabela ordenada por menor utilizacao + download Excel

10. **Distribuicao por Finalidade**
    - Metricas: Total CNH, Total CLT, Total Outro, Total Concurso
    - Tabela por UF com breakdown por finalidade + coluna Total
    - Download Excel

---

## Fase 7: Aba Qualidade - 3 Novas Validacoes

### ANTES
Validacoes de PCLs:
- CNPJ invalido (Critico)
- Razao Social vazia (Critico)
- Cidade vazia (Critico)
- UF invalida (Critico)
- Endereco vazio (Alto)
- CEP invalido (Alto)
- Representante vazio (Medio)
- Nome Fantasia vazio (Baixo)
- Transportadora nao cadastrada (Medio)

### DEPOIS
+3 novas validacoes (apenas para PCLs ativos):
- **E-mail vazio (PCL ativo)** - Severidade: Medio
- **Telefone vazio (PCL ativo)** - Severidade: Medio
- **Valor total da coleta zerado (PCL ativo)** - Severidade: Alto

---

## Fase 8: Aba Ajuda - Documentacao

### ANTES
Documentacao cobria apenas as colunas originais da tabela de PCLs (18 colunas).
Analises documentadas: 7.

### DEPOIS
- Nova secao **"Colunas Adicionais de PCLs (Seletor)"** com tabela explicando os 8 grupos de colunas extras
- 3 novas analises documentadas (8, 9, 10) com descricao e utilidade

---

## Arquivos Impactados

| Arquivo | Alteracao |
|---------|-----------|
| `app.py` | Todas as 8 fases implementadas |

## Funcoes Modificadas

| Funcao | Alteracao |
|--------|-----------|
| `normalize_column_names()` | +60 entradas no column_mapping |
| `process_labs()` | +17 colunas com coercao numerica |
| `identificar_ofensores_pcls()` | +3 validacoes (e-mail, telefone, valor coleta) |
| `_listagem_pcls_fragment()` | Seletor de colunas extras + download expandido |
| `_analise_coletas_fragment()` | Tendencia mensal + capacidade vs realizado |
| `_analises_especificas_fragment()` | +3 analises no selectbox |
| Tab Visao Geral (inline) | Secao Perfil Comercial + grafico finalidade |
| Tab Ajuda (inline) | Documentacao das novas colunas e analises |
