# Design: Correção da Regra de 90 Dias para Status PCL

**Data:** 2026-02-17
**Contexto:** Reportado por Lilian Nascimento — divergência de ~335 PCLs entre o sistema e conferência manual.

## Problema

O painel classifica PCLs como Ativo/Inativo de forma inconsistente com a regra de negócio (90 dias sem coleta). Duas causas identificadas:

1. **Override "Ativo em Coletas"** (`app.py:1944-1945`): A coluna "Ativo em Coletas" sobrescreve o status para "Inativo" mesmo que o PCL tenha ≤90 dias sem coleta.
2. **Recálculo de dias_sem_coleta** (`app.py:1926-1929`): O sistema ignora o valor da planilha e recalcula a partir de `data_ultima_coleta`, podendo gerar divergências por parsing de datas.

### Evidência

| Fonte               | Ativos | Inativos | Total |
|----------------------|--------|----------|-------|
| Sistema (código)     | 1.858  | 3.536    | 5.394 |
| Conferência manual   | 2.193  | 3.201    | 5.394 |
| **Divergência**      | **335 PCLs** | | |

## Decisão

- **Regra única:** `dias_sem_coleta ≤ 90 = Ativo`, `> 90 = Inativo`
- **Fonte de dados:** Usar a coluna "Dias sem coleta" da planilha Excel (não recalcular)
- **Remover** o override pela coluna "Ativo em Coletas"
- **Manter** o filtro de "Ativo (credenciado)" para excluir PCLs descredenciados

## Mudanças

### 1. `process_labs()` em `app.py` (linhas ~1872-1952)

- Remover bloco de `_inativo_em_coletas` (linhas 1872-1879 e 1944-1945)
- Remover recálculo de `dias_sem_coleta` a partir de `data_ultima_coleta` (linhas 1926-1931)
- Usar a coluna `dias_sem_coleta` normalizada da planilha
- Converter para numérico com fallback (PCLs sem valor = 9999 = Inativo)
- Aplicar regra: `dias_sem_coleta ≤ 90 = Ativo`, `> 90 = Inativo`

### 2. Seção "Ajuda" em `app.py`

- Atualizar explicação dos critérios na aba de Ajuda
- Deixar explícito que apenas "Dias sem coleta" define o status
- Explicar que "Ativo (credenciado)" é filtro de entrada, não critério de status
- Remover menções a "Ativo em Coletas" como critério de classificação

### 3. Limpeza

- Remover variável `_inativo_em_coletas` e toda lógica associada
- Manter `data_ultima_coleta` apenas para exibição (não para cálculo)
