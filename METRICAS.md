# Documentação de Métricas - CTOX Analytics

## Visão Geral

Este documento descreve como cada métrica é calculada no sistema de análise de PCLs (Pontos de Coleta/Laboratórios) e Empresas.

---

## 📊 Fontes de Dados

### Arquivo de Empresas (`Acumulado de Coletas - Empresas/*.xlsx`)

| Coluna Original | Descrição |
|-----------------|-----------|
| CNPJ da Empresa | Identificador único da empresa |
| Nome da Empresa | Razão social |
| Cidade / Estado | Localização |
| Data de Credenciamento | Data em que a empresa foi credenciada |
| Representante | Nome do representante comercial |
| Acumulado Coletas Voucher | Total histórico de coletas com voucher |
| Acumulado Coletas Não-Voucher | Total histórico de coletas sem voucher |
| Total Coletas Voucher 2025 | Coletas voucher no ano de 2025 |
| Total Coletas Não-Voucher 2025 | Coletas não-voucher no ano de 2025 |
| Última Coleta (Voucher) | Data da última coleta com voucher |
| Última Coleta (Não-Voucher) | Data da última coleta sem voucher |
| Dias Sem Coleta (Voucher) | Dias desde a última coleta voucher |
| Dias Sem Coleta (Não-Voucher) | Dias desde a última coleta não-voucher |

### Arquivo de PCLs (`Acumulado de Coletas - Labs/*.xlsx`)

| Coluna Original | Descrição |
|-----------------|-----------|
| CNPJ | Identificador único do laboratório |
| Razão Social | Nome oficial |
| Nome Fantasia | Nome comercial |
| Cidade / Estado | Localização |
| Data de credenciamento | Data em que o PCL foi credenciado |
| Representante | Nome do representante comercial |
| Acumulado de Coletas | Total histórico de coletas realizadas |
| Total de Coletas 2025 | Coletas realizadas em 2025 |
| Data da Última Coleta | Data da última coleta |
| Dias sem coleta | Dias desde a última coleta |
| Ativo em Coletas | Status de atividade (True/False) |

---

## 🏢 Métricas de Empresas

### Status de Atividade

Uma empresa é considerada **ATIVA** se atender a **qualquer um** dos critérios:

```
ATIVO = (Dias Sem Coleta Voucher <= 365) OU (Dias Sem Coleta Não-Voucher <= 365)
```

**Explicação:**
- Se a empresa fez qualquer tipo de coleta (voucher OU não-voucher) nos últimos 365 dias, ela é considerada ativa
- Empresas sem nenhuma coleta há mais de 365 dias são marcadas como **INATIVAS**

### Métricas Calculadas

| Métrica | Fórmula | Descrição |
|---------|---------|-----------|
| **Total Empresas** | `COUNT(*)` | Quantidade total de empresas no sistema |
| **Empresas Ativas** | `COUNT(status == 'Ativo')` | Empresas com atividade nos últimos 365 dias |
| **Empresas Inativas** | `Total - Ativas` | Empresas sem atividade há mais de 365 dias |
| **% Ativas** | `(Ativas / Total) * 100` | Percentual de empresas ativas |
| **Total Coletas** | `Voucher + Não-Voucher` | Soma de todas as coletas (ambos os tipos) |
| **Coletas Voucher** | `Acumulado Coletas Voucher` | Total histórico de coletas com voucher |
| **Coletas Não-Voucher** | `Acumulado Coletas Não-Voucher` | Total histórico de coletas sem voucher |
| **Coletas 2025** | `Voucher 2025 + Não-Voucher 2025` | Total de coletas realizadas em 2025 |
| **Última Coleta** | `MAX(Última Voucher, Última Não-Voucher)` | Data mais recente entre os dois tipos |

### Métricas por Cidade

| Métrica | Fórmula | Descrição |
|---------|---------|-----------|
| **PCLs na Cidade** | `COUNT(PCLs na mesma cidade)` | Quantidade de laboratórios credenciados na cidade |
| **PCLs Ativos na Cidade** | `COUNT(PCLs ativos na cidade)` | PCLs com coletas recentes na cidade |
| **PCLs Inativos na Cidade** | `PCLs Total - PCLs Ativos` | PCLs sem atividade recente na cidade |

---

## 🏥 Métricas de PCLs (Laboratórios)

### Status de Atividade

Um PCL é considerado **ATIVO** se atender a **qualquer um** dos critérios:

```
ATIVO = (Dias sem coleta <= 90) OU (Ativo em Coletas == True)
```

**Explicação:**
- PCLs com coleta nos últimos 90 dias são considerados ativos
- O arquivo Excel também contém uma coluna "Ativo em Coletas" que é usada como referência
- PCLs sem coletas há mais de 90 dias são marcados como **INATIVOS**

### Métricas Calculadas

| Métrica | Fórmula | Descrição |
|---------|---------|-----------|
| **Total PCLs** | `COUNT(*)` | Quantidade total de laboratórios no sistema |
| **PCLs Ativos** | `COUNT(status == 'Ativo')` | PCLs com atividade nos últimos 90 dias |
| **PCLs Inativos** | `Total - Ativos` | PCLs sem atividade há mais de 90 dias |
| **% Ativos** | `(Ativos / Total) * 100` | Percentual de PCLs ativos |
| **Coletas Total** | `Acumulado de Coletas` | Total histórico de coletas realizadas |
| **Coletas 2025** | `Total de Coletas 2025` | Coletas realizadas no ano de 2025 |
| **Última Coleta** | `Data da Última Coleta` | Data da última coleta realizada |

### Métricas por Cidade

| Métrica | Fórmula | Descrição |
|---------|---------|-----------|
| **Empresas na Cidade** | `COUNT(Empresas na mesma cidade)` | Quantidade de empresas credenciadas na cidade |
| **Empresas Ativas na Cidade** | `COUNT(Empresas ativas na cidade)` | Empresas com coletas recentes na cidade |
| **Empresas que Utilizaram** | `COUNT(Empresas com voucher > 0)` | Empresas que já usaram voucher |

---

## 📈 Visão Geral - Métricas do Dashboard

### Cards Principais

| Card | Cálculo |
|------|---------|
| **Total PCLs** | Contagem total de registros no arquivo de Labs |
| **PCLs Ativos** | PCLs com `Dias sem coleta <= 90` |
| **Total Empresas** | Contagem total de registros no arquivo de Empresas |
| **Empresas Inativas** | Empresas com ambos `Dias Sem Coleta > 365` |
| **Total Coletas** | Soma de `Acumulado de Coletas` de todos os PCLs |
| **Total Vouchers** | Soma de `Acumulado Coletas Voucher` de todas as empresas |

### Barras de Progresso

| Barra | Cálculo |
|-------|---------|
| **PCLs Ativos** | `(PCLs Ativos / Total PCLs) * 100%` |
| **Empresas Ativas** | `(Empresas Ativas / Total Empresas) * 100%` |

---

## 🔍 Análises Específicas

### 1. PCLs em cidades SEM Empresas credenciadas

```sql
SELECT PCLs WHERE cidade NOT IN (SELECT DISTINCT cidade FROM Empresas)
```

Lista PCLs que estão em cidades onde não há nenhuma empresa credenciada.

### 2. PCLs em cidades COM Empresas INATIVAS

```sql
SELECT PCLs WHERE cidade IN (
  SELECT cidade FROM Empresas 
  GROUP BY cidade 
  HAVING COUNT(status='Ativo') = 0
)
```

Lista PCLs em cidades onde existem empresas, mas todas estão inativas.

### 3. Empresas em cidades SEM PCL credenciado

```sql
SELECT Empresas WHERE cidade NOT IN (SELECT DISTINCT cidade FROM PCLs)
```

Lista empresas que estão em cidades onde não há nenhum laboratório credenciado.

### 4. Empresas em cidades COM PCL INATIVO

```sql
SELECT Empresas WHERE cidade IN (
  SELECT cidade FROM PCLs 
  GROUP BY cidade 
  HAVING COUNT(status='Ativo') = 0
)
```

Lista empresas em cidades onde existem PCLs, mas todos estão inativos.

---

## 📅 Critérios de Tempo

| Entidade | Threshold Atividade | Descrição |
|----------|---------------------|-----------|
| **PCL** | 90 dias | Considerado inativo se última coleta > 90 dias |
| **Empresa** | 365 dias | Considerado inativo se última coleta > 365 dias |

---

## 🔄 Fluxo de Processamento

```
1. Carregar arquivo Excel mais recente de cada pasta
   ├── Acumulado de Coletas - Empresas/*.xlsx
   └── Acumulado de Coletas - Labs/*.xlsx

2. Normalizar nomes das colunas
   └── Converter para formato padronizado (snake_case)

3. Processar Empresas
   ├── Calcular total de coletas (Voucher + Não-Voucher)
   ├── Determinar última coleta (mais recente entre os dois tipos)
   ├── Calcular dias sem atividade
   └── Definir status (Ativo/Inativo)

4. Processar PCLs
   ├── Usar coluna "Ativo em Coletas" se disponível
   ├── OU calcular baseado em "Dias sem coleta"
   └── Definir status (Ativo/Inativo)

5. Aplicar filtros (Estado/Cidade)

6. Exibir métricas e tabelas
```

---

## 📝 Notas Importantes

1. **Voucher vs Não-Voucher**: Empresas podem fazer dois tipos de coleta. Ambos são considerados para determinar atividade.

2. **Dados mais recentes**: O sistema sempre carrega o arquivo Excel mais recente de cada pasta, baseado na data de modificação do arquivo.

3. **Valores ausentes**: Colunas numéricas com valores ausentes (NaN) são tratadas como 0.

4. **Datas inválidas**: Datas que não podem ser convertidas são tratadas como "sem atividade" (9999 dias).

---

*Última atualização: Dezembro 2025*
