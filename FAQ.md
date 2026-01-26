# FAQ - Perguntas Frequentes | CTOX Analytics

## Sumário

1. [Acesso e Login](#acesso-e-login)
2. [Navegação no Dashboard](#navegação-no-dashboard)
3. [Entendendo os Dados](#entendendo-os-dados)
4. [Filtros e Buscas](#filtros-e-buscas)
5. [Colunas e Métricas](#colunas-e-métricas)
6. [Análises Específicas](#análises-específicas)
7. [Exportação de Dados](#exportação-de-dados)
8. [Problemas Comuns](#problemas-comuns)

---

## Acesso e Login

### Como acesso o dashboard?
O acesso é feito através da URL do sistema. Você será redirecionado para a tela de login da Microsoft para autenticação.

### Preciso de permissões especiais?
Sim, você precisa ter uma conta Microsoft corporativa autorizada. Caso não consiga acessar, entre em contato com o administrador do sistema.

### O que fazer se meu login expirar?
O sistema renova automaticamente sua sessão. Se for desconectado, basta clicar em "Continuar com Microsoft" novamente.

---

## Navegação no Dashboard

### Quais são as seções disponíveis?
O dashboard possui as seguintes seções principais:

| Seção | Descrição |
|-------|-----------|
| **Visão Geral** | Métricas gerais e gráficos comparativos por estado |
| **Visão por Estado** | Análise detalhada por UF |
| **Análise de Coletas** | Estatísticas de coletas realizadas |
| **Listagem de PCLs** | Tabela completa de laboratórios/pontos de coleta |
| **Listagem de Empresas** | Tabela completa de empresas credenciadas |
| **Análises Específicas** | Consultas customizadas para cenários específicos |

### Como navego entre as seções?
Use o menu lateral (sidebar) à esquerda para selecionar a seção desejada no campo "Tipo de Análise".

### O que significam os ícones na sidebar?
- ☁️ = Dados carregados do SharePoint (nuvem)
- 💻 = Dados carregados de arquivos locais

---

## Entendendo os Dados

### O que é um PCL?
**PCL** (Ponto de Coleta/Laboratório) é um estabelecimento credenciado para realizar coletas de exames toxicológicos.

### Qual a diferença entre PCL Ativo e Inativo?
| Status | Critério |
|--------|----------|
| **Ativo** | Realizou coleta nos últimos **90 dias** |
| **Inativo** | Sem coletas há mais de **90 dias** |

### Qual a diferença entre Empresa Ativa e Inativa?
| Status | Critério |
|--------|----------|
| **Ativa** | Utilizou voucher ou fez coleta nos últimos **365 dias** |
| **Inativa** | Sem atividade há mais de **365 dias** |

### O que são Vouchers?
Vouchers são créditos que empresas utilizam para pagar coletas de exames toxicológicos de seus funcionários.

### Os dados são atualizados em tempo real?
Não. Os dados são atualizados de acordo com a disponibilidade nas planilhas no SharePoint corporativo. As planilhas estão localizadas no SharePoint na pasta **"Data Analysis"**, dentro das subpastas:
- **"Acumulado de Coletas - Empresas"** (para dados de empresas)
- **"Acumulado de Coletas - Labs"** (para dados de PCLs)

O sistema carrega automaticamente o arquivo mais recente de cada pasta sempre que a página é recarregada.

---

## Filtros e Buscas

### Como filtro por Estado ou Cidade?
Use os filtros na sidebar (menu lateral):
1. Selecione o **Estado (UF)** desejado
2. Opcionalmente, selecione uma **Cidade** específica
3. Os dados serão filtrados automaticamente

### Posso filtrar por múltiplos estados?
Não. O filtro permite selecionar apenas um estado por vez. Para ver todos os estados, selecione a opção "Todos".

### Como busco um PCL ou Empresa específica?
Na tabela de listagem, você pode usar a busca nativa do navegador (Ctrl+F) ou ordenar as colunas clicando no cabeçalho.

---

## Colunas e Métricas

### O que significa cada coluna na Listagem de PCLs?

| Coluna | Descrição |
|--------|-----------|
| **CNPJ** | Número de identificação fiscal do laboratório |
| **Razão Social** | Nome oficial registrado |
| **Nome Fantasia** | Nome comercial do estabelecimento |
| **Cidade / UF** | Localização do PCL |
| **Data Credenciamento** | Data em que o PCL foi credenciado no sistema |
| **Representante** | Nome do representante comercial responsável |
| **Transportadora** | Empresas de transporte disponíveis para coleta na cidade |
| **Frequência** | Frequência de coleta disponível (DIARIO, SEMANAL, etc.) |
| **Coletas Total** | Total histórico de coletas realizadas |
| **Coletas 2025** | Coletas realizadas no ano de 2025 |
| **Última Coleta** | Data da última coleta realizada |
| **Status** | Ativo ou Inativo (baseado nos últimos 90 dias) |
| **Empresas na Cidade** | Quantidade de empresas credenciadas na mesma cidade |

### O que significa quando a coluna Transportadora mostra múltiplos valores?
Quando uma cidade possui mais de uma transportadora disponível, elas são exibidas separadas por " | ". Exemplo:
```
AIRLAB | BIOMED LOG | CORREIOS AP
```
Isso significa que existem 3 opções de transporte para coletas naquela cidade.

### O que significa cada valor de Frequência?

| Valor | Significado |
|-------|-------------|
| **DIARIO** | Coleta todos os dias úteis |
| **SEMANAL** | Coleta uma vez por semana |
| **2ª, 4ª E 6ª** | Coleta nas segundas, quartas e sextas-feiras |
| **3ª E 5ª** | Coleta nas terças e quintas-feiras |
| **ALTERNADO** | Coleta em dias alternados |
| **CONSULTAR** | Frequência variável, necessário consultar |

### O que significa cada coluna na Listagem de Empresas?

| Coluna | Descrição |
|--------|-----------|
| **CNPJ** | Número de identificação fiscal da empresa |
| **Razão Social** | Nome oficial registrado |
| **Nome Fantasia** | Nome comercial |
| **Cidade / UF** | Localização da empresa |
| **Data Credenciamento** | Data em que a empresa foi credenciada |
| **Vouchers** | Total de vouchers utilizados (histórico) |
| **Vouchers 2025** | Vouchers utilizados em 2025 |
| **Última Utilização** | Data da última utilização de voucher |
| **Status** | Ativo ou Inativo (baseado nos últimos 365 dias) |
| **PCLs na Cidade** | Quantidade de laboratórios na mesma cidade |

---

## Análises Específicas

### Quais análises específicas estão disponíveis?

1. **PCLs em cidades SEM Empresas credenciadas**
   - Lista laboratórios em cidades onde não há nenhuma empresa cliente
   - Útil para identificar PCLs que podem precisar de prospecção comercial

2. **PCLs em cidades COM Empresas INATIVAS (365 dias)**
   - Lista laboratórios em cidades onde todas as empresas estão inativas
   - Indica oportunidades de reativação de clientes

3. **Empresas em cidades SEM PCL credenciado**
   - Lista empresas que não têm laboratório disponível em sua cidade
   - Indica necessidade de credenciar novos PCLs

4. **Empresas em cidades COM PCL INATIVO (90 dias)**
   - Lista empresas em cidades onde todos os PCLs estão inativos
   - Indica risco de perda de clientes por falta de atendimento

5. **Top PCLs por volume de coletas**
   - Ranking dos 50 PCLs com maior volume de coletas
   - Útil para identificar os principais parceiros

6. **Estados com menor cobertura**
   - Lista estados ordenados por quantidade de cidades atendidas
   - Indica oportunidades de expansão

### Como interpreto os resultados das análises?
Cada análise mostra uma tabela com os registros encontrados. O número total de resultados aparece no topo da tabela. Você pode exportar os resultados para Excel para análise mais detalhada.

---

## Exportação de Dados

### Como exporto os dados para Excel?
Em cada seção de listagem ou análise, há um botão **"📥 Download Excel"**. Clique nele para baixar os dados filtrados em formato `.xlsx`.

### Os filtros aplicados afetam a exportação?
Sim. O arquivo Excel exportado conterá apenas os dados visíveis na tela, respeitando os filtros de Estado e Cidade aplicados.

### Posso exportar todas as análises de uma vez?
Não. Cada análise deve ser exportada individualmente.

---

## Problemas Comuns

### O dashboard está lento. O que fazer?
- Verifique sua conexão com a internet
- Tente recarregar a página (F5)
- Se o problema persistir, limpe o cache do navegador

### Os dados parecem desatualizados. O que fazer?
1. Verifique a data de atualização no rodapé do dashboard
2. Se necessário, clique no botão de recarregar (F5) para forçar uma nova leitura dos dados
3. Se o problema persistir, verifique se os arquivos fonte foram atualizados

### Aparece erro "Arquivo está aberto em outro programa"
Isso ocorre quando o arquivo Excel fonte está aberto em outro programa. Feche o Excel e recarregue a página.

### Não consigo ver os dados de Transportadora/Frequência
Verifique se o arquivo `CONSULTA MATRIZ LOGISTICA.1.xlsx` está presente na pasta do sistema. Se a coluna aparecer vazia, pode ser que a cidade do PCL não esteja cadastrada na matriz logística.

### Uma cidade não aparece nos resultados
- Verifique se o nome da cidade está escrito corretamente nos dados fonte
- Algumas cidades podem ter nomes ligeiramente diferentes (ex: "SAO PAULO" vs "SÃO PAULO")
- O sistema normaliza os nomes, mas diferenças significativas podem impedir o match

### Como reporto um problema ou solicito uma melhoria?
Entre em contato com a equipe de TI ou com o administrador do sistema informando:
- Descrição detalhada do problema
- Prints de tela (se aplicável)
- Passos para reproduzir o erro

---

## Glossário

| Termo | Definição |
|-------|-----------|
| **PCL** | Ponto de Coleta/Laboratório credenciado |
| **Voucher** | Crédito utilizado por empresas para pagar coletas |
| **Credenciamento** | Processo de cadastro e autorização no sistema |
| **Coleta** | Procedimento de coleta de material para exame toxicológico |
| **SharePoint** | Plataforma Microsoft onde os dados ficam armazenados |
| **Status Ativo** | Entidade com atividade recente no sistema |
| **Status Inativo** | Entidade sem atividade por período prolongado |

---

*Última atualização: Janeiro 2026*
