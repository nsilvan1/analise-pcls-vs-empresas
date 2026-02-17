# Correção da Regra de 90 Dias — Plano de Implementação

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Corrigir a classificação Ativo/Inativo dos PCLs para usar exclusivamente a coluna "Dias sem coleta" da planilha com a regra de 90 dias, removendo o override da coluna "Ativo em Coletas".

**Architecture:** Modificação cirúrgica na função `process_labs()` do `app.py` — remover o recálculo de `dias_sem_coleta` e o override por "Ativo em Coletas", substituindo por leitura direta da coluna da planilha. Atualizar a seção Ajuda com a documentação correta.

**Tech Stack:** Python, Pandas, Streamlit

---

### Task 1: Remover lógica de `_inativo_em_coletas`

**Files:**
- Modify: `app.py:1872-1879` (bloco PASSO 2)
- Modify: `app.py:1943-1945` (override)
- Modify: `app.py:1950-1952` (LOG)

**Step 1: Remover bloco PASSO 2 (linhas 1872-1879)**

Substituir:
```python
    # PASSO 2: Identificar PCLs com "Ativo em Coletas" = False para marcar como Inativo
    df['_inativo_em_coletas'] = False
    if 'ativo em coletas' in df.columns:
        df['_inativo_em_coletas'] = df['ativo em coletas'].apply(
            lambda x: not (x == True or str(x).lower() in ['true', '1', 'sim', 'yes', 's', 'y'])
        )
        inativos_coletas = df['_inativo_em_coletas'].sum()
        print(f"[LOG PCLs] PCLs com 'Ativo em Coletas' = False: {inativos_coletas}")
```

Por (remover completamente — não substituir por nada, manter linha em branco):
```python
```

**Step 2: Remover override (linhas 1943-1945)**

Substituir:
```python
    # Marcar PCLs com "Ativo em Coletas" = False como Inativo (sobrescreve o status calculado acima)
    if '_inativo_em_coletas' in df.columns:
        df.loc[df['_inativo_em_coletas'] == True, 'status'] = 'Inativo'
```

Por (remover completamente):
```python
```

**Step 3: Simplificar LOG (linhas 1947-1952)**

Substituir:
```python
    # LOG: Contagem de PCLs por status
    ativos = (df['status'] == 'Ativo').sum()
    inativos = (df['status'] == 'Inativo').sum()
    inativos_em_coletas = df['_inativo_em_coletas'].sum() if '_inativo_em_coletas' in df.columns else 0
    inativos_por_dias = inativos - inativos_em_coletas
    print(f"[LOG PCLs] Total: {len(df)} | Ativos: {ativos} | Inativos: {inativos} (por 'Ativo em Coletas'=False: {inativos_em_coletas}, por >90 dias: {inativos_por_dias})")
```

Por:
```python
    # LOG: Contagem de PCLs por status
    ativos = (df['status'] == 'Ativo').sum()
    inativos = (df['status'] == 'Inativo').sum()
    print(f"[LOG PCLs] Total: {len(df)} | Ativos: {ativos} | Inativos: {inativos} (regra: dias_sem_coleta <= 90)")
```

---

### Task 2: Usar coluna `dias_sem_coleta` da planilha em vez de recalcular

**Files:**
- Modify: `app.py:1926-1941` (bloco de cálculo de dias e status)

**Step 1: Substituir recálculo por leitura da planilha (linhas 1926-1941)**

Substituir:
```python
    # Calcular dias_sem_coleta a partir de data_ultima_coleta (ignora valor da planilha)
    hoje = pd.Timestamp.now().normalize()
    if 'data_ultima_coleta' in df.columns:
        df['dias_sem_coleta'] = (hoje - df['data_ultima_coleta']).dt.days
    else:
        df['dias_sem_coleta'] = None

    # Calcular status baseado em dias sem coleta (regra: <=90 dias = Ativo, >90 dias = Inativo)
    # PCLs com "Ativo em Coletas" = False são sempre marcados como Inativo
    if df['dias_sem_coleta'].notna().any():
        dias = df['dias_sem_coleta'].fillna(9999)
        df['status'] = dias.apply(lambda x: 'Ativo' if x <= 90 else 'Inativo')
    else:
        # Fallback: usar acumulado (se não tem data_ultima_coleta, considera ativo se tem coletas)
        df['status'] = df['acumulado_coletas'].apply(lambda x: 'Ativo' if x > 0 else 'Inativo')
        print(f"[LOG PCLs] AVISO: Coluna 'data_ultima_coleta' não encontrada. Usando fallback por acumulado.")
```

Por:
```python
    # Usar dias_sem_coleta da planilha (coluna já normalizada pelo normalize_column_names)
    if 'dias_sem_coleta' in df.columns:
        df['dias_sem_coleta'] = pd.to_numeric(df['dias_sem_coleta'], errors='coerce')
    else:
        df['dias_sem_coleta'] = None
        print(f"[LOG PCLs] AVISO: Coluna 'dias_sem_coleta' não encontrada na planilha.")

    # Calcular status baseado em dias sem coleta (regra: <=90 dias = Ativo, >90 dias = Inativo)
    if df['dias_sem_coleta'].notna().any():
        dias = df['dias_sem_coleta'].fillna(9999)
        df['status'] = dias.apply(lambda x: 'Ativo' if x <= 90 else 'Inativo')
    else:
        df['status'] = df['acumulado_coletas'].apply(lambda x: 'Ativo' if x > 0 else 'Inativo')
        print(f"[LOG PCLs] AVISO: Coluna 'dias_sem_coleta' sem dados. Usando fallback por acumulado.")
```

**Nota:** A coluna `data_ultima_coleta` continua sendo convertida para datetime na linha 1922 — isso é mantido pois é usada para exibição na listagem.

---

### Task 3: Atualizar docstring da função `process_labs()`

**Files:**
- Modify: `app.py:1817-1833` (docstring)

**Step 1: Substituir docstring**

Substituir:
```python
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
```

Por:
```python
    """
    Processa dados de labs (PCLs).

    FILTRO DE ENTRADA:
    - PCLs com "Ativo (credenciado)" = False são EXCLUÍDOS (descredenciados)

    CRITÉRIO DE ATIVIDADE (para PCLs credenciados):
    - Dias sem coleta <= 90 → ATIVO
    - Dias sem coleta > 90  → INATIVO

    Fonte: coluna "Dias sem coleta" da planilha Excel (não recalculada).
    Fallback: se a coluna não existir, usa Acumulado de Coletas > 0.
    """
```

---

### Task 4: Atualizar seção Ajuda

**Files:**
- Modify: `app.py:4487-4506` (seção "Qual a diferença entre PCL Ativo e Inativo?")

**Step 1: Substituir bloco de ajuda sobre descredenciados/ativo-inativo**

Substituir (linhas 4487-4506):
```python
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
        - **PCLs descredenciados** (coluna "Ativo (credenciado)" = Falso) são **excluídos** antes de qualquer análise
        - Para os PCLs restantes, usa-se "Dias sem coleta": ≤ 90 dias = Ativo, > 90 dias = Inativo
        """)
```

Por:
```python
        st.warning("""
        ⚠️ **IMPORTANTE:** PCLs **descredenciados** são excluídos completamente do sistema e não aparecem em nenhuma análise.

        **O que são PCLs descredenciados?**
        - São PCLs que tiveram seu credenciamento revogado ou cancelado
        - Na planilha Excel, esses PCLs aparecem com a coluna **"Ativo (credenciado)" = Falso**
        - Esses PCLs são automaticamente filtrados e não entram em nenhuma contagem, análise ou listagem
        """)
        col1, col2 = st.columns(2)
        with col1:
            st.success("**ATIVO**: Até **90 dias** sem coleta")
        with col2:
            st.error("**INATIVO**: Mais de **90 dias** sem coleta")

        st.markdown("""
        **Critérios de classificação:**

        **1. Filtro de entrada — Credenciamento:**
        - PCLs com coluna **"Ativo (credenciado)" = Falso** são **excluídos** antes de qualquer análise
        - Apenas PCLs **credenciados** aparecem no dashboard

        **2. Regra de atividade — 90 dias:**
        - O status Ativo/Inativo é definido **exclusivamente** pela coluna **"Dias sem coleta"** da planilha
        - **Ativo:** Dias sem coleta **≤ 90**
        - **Inativo:** Dias sem coleta **> 90**
        - A coluna "Ativo em Coletas" da planilha **não** é utilizada para essa classificação
        """)
```

---

### Task 5: Commit

**Step 1: Commit das alterações**

```bash
git add app.py
git commit -m "fix: usa coluna 'dias sem coleta' da planilha e remove override 'Ativo em Coletas'

- Remove recálculo de dias_sem_coleta a partir de data_ultima_coleta
- Usa valor original da planilha Excel para classificação
- Remove override pela coluna 'Ativo em Coletas' (~335 PCLs divergentes)
- Regra única: dias_sem_coleta <= 90 = Ativo, > 90 = Inativo
- Atualiza documentação na aba Ajuda

Reportado por: Lilian Nascimento"
```
