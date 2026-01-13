"""
Pagina de debug para verificar secrets no Streamlit Cloud
REMOVER APOS RESOLVER O PROBLEMA!
"""
import streamlit as st

st.set_page_config(page_title="Debug Secrets", page_icon="🔧")

st.title("Debug Secrets")
st.warning("REMOVA ESTA PAGINA APOS RESOLVER O PROBLEMA!")

st.header("1. Verificando secoes disponiveis")
try:
    secoes = list(st.secrets.keys())
    st.success(f"Secoes encontradas: {secoes}")
except Exception as e:
    st.error(f"Erro ao ler secrets: {e}")

st.header("2. Secao [graph]")
try:
    graph = st.secrets.get("graph", {})
    if graph:
        st.write("- tenant_id:", graph.get("tenant_id", "NAO ENCONTRADO")[:10] + "...")
        st.write("- client_id:", graph.get("client_id", "NAO ENCONTRADO")[:10] + "...")

        # Verificar client_secret (mostrar apenas primeiros e ultimos caracteres)
        cs = graph.get("client_secret", "")
        if cs:
            st.write(f"- client_secret: {cs[:5]}...{cs[-5:]} (tamanho: {len(cs)})")
            # Verificar caracteres especiais
            if "~" in cs:
                st.info("client_secret contem '~'")
        else:
            st.error("client_secret NAO ENCONTRADO ou VAZIO")

        st.write("- hostname:", graph.get("hostname", "NAO ENCONTRADO"))
        st.write("- site_path:", graph.get("site_path", "NAO ENCONTRADO"))
    else:
        st.error("Secao [graph] NAO ENCONTRADA!")
except Exception as e:
    st.error(f"Erro: {e}")

st.header("3. Secao [auth]")
try:
    auth = st.secrets.get("auth", {})
    if auth:
        st.write("- tenant_id:", auth.get("tenant_id", "NAO ENCONTRADO")[:10] + "...")
        st.write("- client_id:", auth.get("client_id", "NAO ENCONTRADO")[:10] + "...")

        cs = auth.get("client_secret", "")
        if cs:
            st.write(f"- client_secret: {cs[:5]}...{cs[-5:]} (tamanho: {len(cs)})")
        else:
            st.error("client_secret NAO ENCONTRADO ou VAZIO")
    else:
        st.error("Secao [auth] NAO ENCONTRADA!")
except Exception as e:
    st.error(f"Erro: {e}")

st.header("4. Secao [sharepoint_logistica]")
try:
    sp = st.secrets.get("sharepoint_logistica", {})
    if sp:
        st.write("- hostname:", sp.get("hostname", "NAO ENCONTRADO"))
        st.write("- site_path:", sp.get("site_path", "NAO ENCONTRADO"))
        st.write("- file_path:", sp.get("file_path", "NAO ENCONTRADO")[:30] + "...")
        st.success("Secao [sharepoint_logistica] encontrada!")
    else:
        st.error("Secao [sharepoint_logistica] NAO ENCONTRADA!")
except Exception as e:
    st.error(f"Erro: {e}")

st.header("5. Teste de conexao MSAL")
try:
    import msal

    graph = st.secrets.get("graph", {})
    if graph:
        app = msal.ConfidentialClientApplication(
            client_id=graph.get("client_id"),
            authority=f"https://login.microsoftonline.com/{graph.get('tenant_id')}",
            client_credential=graph.get("client_secret"),
        )

        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

        if "access_token" in result:
            st.success("Token obtido com sucesso!")
        else:
            st.error(f"Erro ao obter token: {result.get('error')}")
            st.error(f"Descricao: {result.get('error_description')}")
except Exception as e:
    st.error(f"Erro MSAL: {e}")
