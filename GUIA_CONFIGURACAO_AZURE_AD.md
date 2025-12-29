# 🔐 Guia Completo: Configurar Login Microsoft com Novo Aplicativo Azure AD

Este guia explica como configurar um novo aplicativo Azure AD para autenticação no sistema CTOX Analytics (Empresa vs PCL).

---

## 📋 Pré-requisitos

- Acesso ao [Azure Portal](https://portal.azure.com)
- Permissões para criar aplicativos no Azure AD (geralmente requer Admin)
- Tenant ID do Azure AD da sua organização

---

## 🚀 Passo 1: Criar o Aplicativo no Azure AD

### 1.1 Acessar o Azure Portal

1. Acesse [https://portal.azure.com](https://portal.azure.com)
2. Faça login com uma conta que tenha permissões de administrador
3. No menu, procure por **"Azure Active Directory"** ou **"Microsoft Entra ID"**

### 1.2 Registrar Novo Aplicativo

1. No menu lateral, clique em **"App registrations"** (Registros de aplicativo)
2. Clique no botão **"+ New registration"** (+ Novo registro)
3. Preencha os campos:
   - **Name**: Nome do aplicativo (ex: "CTOX Analytics Authentication")
   - **Supported account types**: 
     - Se for apenas para sua organização: **"Accounts in this organizational directory only"**
     - Se for multi-tenant: **"Accounts in any organizational directory"**
   - **Redirect URI**: 
     - **Platform**: Selecione **"Web"**
     - **URI**: Adicione:
       - `http://localhost:8501` (para desenvolvimento local)
       - `https://empresavspcl.streamlit.app` (para produção - ajuste conforme sua URL)
4. Clique em **"Register"**

### 1.3 Obter as Credenciais

Após criar o aplicativo, você verá a página **"Overview"**:

- **Application (client) ID**: Este é o `client_id` que você precisará
- **Directory (tenant) ID**: Este é o `tenant_id` que você precisará

**⚠️ IMPORTANTE**: Anote esses valores, você precisará deles!

---

## 🔑 Passo 2: Criar Client Secret

### 2.1 Gerar o Secret

1. No menu lateral do aplicativo, clique em **"Certificates & secrets"** (Certificados e segredos)
2. Na aba **"Client secrets"**, clique em **"+ New client secret"**
3. Preencha:
   - **Description**: Uma descrição (ex: "CTOX Analytics Secret")
   - **Expires**: Escolha a validade (recomendado: 24 meses)
4. Clique em **"Add"**

### 2.2 Copiar o Secret

⚠️ **ATENÇÃO**: O valor do secret só aparece UMA VEZ! Copie imediatamente.

- O campo **"Value"** mostrará o secret (começa com algo como `abc123~...`)
- **Copie este valor** - você precisará dele como `client_secret`

---

## 🔗 Passo 3: Configurar Redirect URIs

### 3.1 Adicionar Redirect URIs

1. No menu lateral, clique em **"Authentication"** (Autenticação)
2. Na seção **"Redirect URIs"**, adicione:
   - `http://localhost:8501` (desenvolvimento)
   - `https://empresavspcl.streamlit.app` (produção - ajuste conforme necessário)
3. Na seção **"Implicit grant and hybrid flows"**, marque:
   - ✅ **ID tokens** (opcional, mas recomendado)
4. Clique em **"Save"**

---

## 🔐 Passo 4: Configurar Permissões (API Permissions)

### 4.1 Adicionar Permissões do Microsoft Graph

1. No menu lateral, clique em **"API permissions"** (Permissões de API)
2. Clique em **"+ Add a permission"**
3. Selecione **"Microsoft Graph"**
4. Selecione **"Delegated permissions"**
5. Adicione as seguintes permissões:
   - ✅ **User.Read** (Ler perfil do usuário)
   - ✅ **offline_access** (Manter acesso aos dados que você concedeu - necessário para refresh token)
6. Clique em **"Add permissions"**

### 4.2 Conceder Consentimento do Administrador (se necessário)

- Se você for administrador, clique em **"Grant admin consent for [sua organização]"**
- Isso permite que todos os usuários usem o aplicativo sem precisar consentir individualmente

---

## ⚙️ Passo 5: Configurar no Código

### 5.1 Para Desenvolvimento Local

Crie ou edite o arquivo `.streamlit/secrets.toml` na raiz do projeto:

```toml
[auth]
client_id = "SEU_CLIENT_ID_AQUI"
client_secret = "SEU_CLIENT_SECRET_AQUI"
tenant_id = "SEU_TENANT_ID_AQUI"
redirect_uri_local = "http://localhost:8501"
redirect_uri_prod = "https://empresavspcl.streamlit.app"
authority = "https://login.microsoftonline.com/SEU_TENANT_ID_AQUI"
scope = ["https://graph.microsoft.com/User.Read", "offline_access"]
```

**Exemplo real:**
```toml
[auth]
client_id = "7c19a480-bc01-45b6-8b74-6a3e6154e876"
client_secret = "abc123~xyz789_secret_value"
tenant_id = "fee1b506-24b6-444a-919e-83df9442dc5d"
redirect_uri_local = "http://localhost:8501"
redirect_uri_prod = "https://empresavspcl.streamlit.app"
authority = "https://login.microsoftonline.com/fee1b506-24b6-444a-919e-83df9442dc5d"
scope = ["https://graph.microsoft.com/User.Read", "offline_access"]
```

### 5.2 Para Produção (Streamlit Cloud)

1. Acesse seu app no [Streamlit Cloud](https://share.streamlit.io)
2. Vá em **Settings** → **Secrets**
3. Cole o mesmo conteúdo do `secrets.toml`:

```toml
[auth]
client_id = "SEU_CLIENT_ID_AQUI"
client_secret = "SEU_CLIENT_SECRET_AQUI"
tenant_id = "SEU_TENANT_ID_AQUI"
redirect_uri_local = "http://localhost:8501"
redirect_uri_prod = "https://empresavspcl.streamlit.app"
authority = "https://login.microsoftonline.com/SEU_TENANT_ID_AQUI"
scope = ["https://graph.microsoft.com/User.Read", "offline_access"]
```

### 5.3 Alternativa: Variáveis de Ambiente

Se preferir usar variáveis de ambiente (útil para Docker, CI/CD, etc.):

```bash
# Windows PowerShell
$env:AZURE_CLIENT_ID = "seu-client-id"
$env:AZURE_CLIENT_SECRET = "seu-client-secret"
$env:AZURE_TENANT_ID = "seu-tenant-id"

# Linux/Mac
export AZURE_CLIENT_ID="seu-client-id"
export AZURE_CLIENT_SECRET="seu-client-secret"
export AZURE_TENANT_ID="seu-tenant-id"
```

O código em `auth_microsoft.py` já está configurado para ler essas variáveis como fallback.

---

## ✅ Passo 6: Testar a Configuração

### 6.1 Testar Localmente

1. Certifique-se de que o arquivo `.streamlit/secrets.toml` está configurado
2. Execute o app:
   ```bash
   streamlit run app_streamlit_churn.py
   ```
3. Acesse `http://localhost:8501`
4. Você deve ver a tela de login
5. Clique em **"Entrar com Microsoft"**
6. Faça login com uma conta Microsoft válida
7. Você deve ser redirecionado de volta e autenticado

### 6.2 Verificar Logs

Se houver erros, verifique:
- Os logs no console do Streamlit
- O console do navegador (F12)
- Se os redirect URIs estão corretos no Azure AD
- Se o client_secret está correto e não expirou

---

## 🔍 Troubleshooting (Solução de Problemas)

### Erro: "Configurações de autenticação Microsoft incompletas"

**Causa**: Faltam credenciais (client_id, client_secret ou tenant_id)

**Solução**: 
- Verifique se o arquivo `.streamlit/secrets.toml` existe
- Verifique se a seção `[auth]` está correta
- Verifique se não há espaços extras ou aspas incorretas

### Erro: "AADSTS50011: The redirect URI specified in the request does not match"

**Causa**: O redirect URI no código não corresponde ao configurado no Azure AD

**Solução**:
1. Verifique o redirect URI no Azure AD (Authentication → Redirect URIs)
2. Verifique o `redirect_uri_local` ou `redirect_uri_prod` no `secrets.toml`
3. Certifique-se de que são EXATAMENTE iguais (incluindo http vs https, porta, etc.)

### Erro: "AADSTS7000215: Invalid client secret is provided"

**Causa**: O client_secret está incorreto ou expirou

**Solução**:
1. Gere um novo client secret no Azure AD
2. Atualize o `client_secret` no `secrets.toml`
3. Se estiver em produção, atualize também no Streamlit Cloud

### Erro: "AADSTS65001: The user or administrator has not consented"

**Causa**: Permissões não foram concedidas

**Solução**:
1. No Azure AD, vá em **API permissions**
2. Clique em **"Grant admin consent"**
3. Ou faça login e conceda consentimento quando solicitado

### Token não renova automaticamente

**Causa**: Falta a permissão `offline_access`

**Solução**:
1. No Azure AD, adicione a permissão `offline_access` em **API permissions**
2. Adicione `"offline_access"` ao array `scope` no `secrets.toml`

---

## 📝 Resumo das Configurações Necessárias

### No Azure AD:
- ✅ Aplicativo registrado
- ✅ Client ID anotado
- ✅ Client Secret criado e copiado
- ✅ Redirect URIs configurados (localhost e produção)
- ✅ Permissões adicionadas (User.Read, offline_access)
- ✅ Consentimento do administrador concedido (se aplicável)

### No Código:
- ✅ Arquivo `.streamlit/secrets.toml` criado com `[auth]`
- ✅ Todas as credenciais preenchidas
- ✅ Redirect URIs correspondem aos do Azure AD

---

## 🔒 Segurança

⚠️ **IMPORTANTE**:
- **NUNCA** commite o arquivo `.streamlit/secrets.toml` no Git
- Adicione `.streamlit/secrets.toml` ao `.gitignore`
- Use variáveis de ambiente em ambientes de CI/CD
- Rotacione os client secrets periodicamente
- Use secrets diferentes para desenvolvimento e produção quando possível

---

## 📚 Referências

- [Documentação MSAL Python](https://github.com/AzureAD/microsoft-authentication-library-for-python)
- [Azure AD App Registration](https://docs.microsoft.com/azure/active-directory/develop/quickstart-register-app)
- [OAuth 2.0 Authorization Code Flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-auth-code-flow)

---

## 💡 Dicas Extras

1. **Múltiplos Ambientes**: Considere criar aplicativos separados no Azure AD para dev, staging e produção
2. **Logs**: Habilite logging detalhado no Azure AD para debug
3. **Validação de Domínio**: Se quiser restringir a apenas contas `@synvia.com`, configure isso no Azure AD
4. **Refresh Token**: O código já implementa renovação automática de tokens - não é necessário fazer nada adicional

---

**Pronto!** Seu novo aplicativo Azure AD está configurado e pronto para uso! 🎉

