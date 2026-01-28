# Plano de Testes - Ordenação Alfabética

## 📋 Alterações Implementadas

1. Gráfico "PCLs por Cidade" - ordenação alfabética
2. Gráfico "Empresas por Cidade" - ordenação alfabética
3. Tabela "Resumo por Estado" - ordenação alfabética por UF

---

## ✅ Checklist de Testes

### Testes Básicos

#### Gráficos de Cidades
- [ ] Selecionar um estado e verificar que as cidades estão em ordem alfabética (A-Z) no gráfico de PCLs
- [ ] Verificar que as cidades estão em ordem alfabética no gráfico de Empresas
- [ ] Confirmar que a ordem das cidades é **idêntica** nos dois gráficos (facilita cruzamento)
- [ ] Testar com estado com muitas cidades (ex: SP, MG)
- [ ] Testar com estado com poucas cidades (ex: AC, RR)

#### Tabela Resumo
- [ ] Verificar que os estados estão ordenados alfabeticamente por UF (AC, AL, AP, AM, BA, ...)
- [ ] Confirmar que é fácil encontrar um estado específico na tabela
- [ ] Verificar que os dados numéricos estão corretos após a ordenação

### Testes de Regressão

- [ ] Verificar que outros gráficos da aplicação ainda funcionam normalmente
- [ ] Verificar que não há erros no console do navegador
- [ ] Verificar que não há erros nos logs do Streamlit

---

## 🔍 Cenários de Teste

### Cenário 1: Estado com Muitas Cidades
1. Acessar "Visão por Estado"
2. Selecionar "São Paulo" (ou MG, RS)
3. Verificar ordenação alfabética nos gráficos
4. Confirmar que todas as cidades são exibidas

### Cenário 2: Cruzamento de Dados
1. Selecionar um estado
2. Identificar uma cidade no gráfico de PCLs (ex: "Belo Horizonte")
3. Verificar que a mesma cidade está na **mesma posição** no gráfico de Empresas
4. Comparar valores entre os dois gráficos

### Cenário 3: Navegação entre Estados
1. Testar vários estados diferentes
2. Verificar que a ordenação funciona consistentemente em todos

---

## ✅ Critérios de Aceitação

- [x] Gráficos de cidades ordenados alfabeticamente
- [x] Tabela resumo ordenada alfabeticamente por UF
- [x] Ordem idêntica entre gráficos de PCLs e Empresas
- [x] Sem erros no console/logs
- [x] Outros gráficos não foram afetados

---

**Data**: 26/01/2026
