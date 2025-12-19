"""
Script de teste para verificar conexão com SharePoint/OneDrive
"""
import streamlit as st
from sp_connector import get_sp_connector

# Configurar página - DEVE SER A PRIMEIRA LINHA DO STREAMLIT
st.set_page_config(page_title="Teste de Conexão SharePoint", layout="wide")

st.title("🔍 Teste de Conexão SharePoint/OneDrive")

# Tentar criar conexão
st.subheader("1. Criando Conexão")
try:
    sp_connector = get_sp_connector()
    if sp_connector:
        st.success("✅ Conexão criada com sucesso!")
        st.json({
            "Modo": "OneDrive" if sp_connector.is_onedrive else "SharePoint Site",
            "User UPN": sp_connector.user_upn if sp_connector.is_onedrive else "N/A",
            "Hostname": sp_connector.hostname if not sp_connector.is_onedrive else "N/A",
            "Site Path": sp_connector.site_path if not sp_connector.is_onedrive else "N/A"
        })
    else:
        st.error("❌ Falha ao criar conexão: retornou None")
        st.stop()
except Exception as e:
    st.error(f"❌ Erro ao criar conexão: {e}")
    st.stop()

# Testar listagem de arquivos na raiz do Documents
st.subheader("2. Listando Arquivos na Raiz de Documents")
try:
    files_root = sp_connector.list_files("")
    st.success(f"✅ Encontrados {len(files_root)} itens na raiz")
    
    # Mostrar primeiros 20 itens
    if files_root:
        st.write("**Itens encontrados:**")
        for item in files_root[:20]:
            item_type = "📁 Pasta" if item.get("folder") else "📄 Arquivo"
            st.write(f"- {item_type}: {item.get('name', 'N/A')}")
    else:
        st.warning("⚠️ Nenhum item encontrado na raiz")
except Exception as e:
    st.error(f"❌ Erro ao listar arquivos na raiz: {e}")

# Testar listagem na pasta Data Analysis
st.subheader("3. Listando Arquivos em 'Data Analysis'")
try:
    files_data_analysis = sp_connector.list_files("Data Analysis")
    st.success(f"✅ Encontrados {len(files_data_analysis)} itens em 'Data Analysis'")
    
    if files_data_analysis:
        st.write("**Itens encontrados:**")
        for item in files_data_analysis:
            item_type = "📁 Pasta" if item.get("folder") else "📄 Arquivo"
            name = item.get('name', 'N/A')
            last_modified = item.get('lastModifiedDateTime', 'N/A')
            st.write(f"- {item_type}: **{name}** (Modificado: {last_modified})")
    else:
        st.warning("⚠️ Nenhum item encontrado em 'Data Analysis'")
except Exception as e:
    st.error(f"❌ Erro ao listar arquivos em 'Data Analysis': {e}")
    st.code(str(e))

# Testar listagem na pasta de Empresas
st.subheader("4. Listando Arquivos em 'Data Analysis/Acumulado de Coletas - Empresas'")
try:
    folder_path_empresas = "Data Analysis/Acumulado de Coletas - Empresas"
    files_empresas = sp_connector.list_files(folder_path_empresas)
    st.success(f"✅ Encontrados {len(files_empresas)} itens em '{folder_path_empresas}'")
    
    if files_empresas:
        st.write("**Itens encontrados:**")
        excel_files = []
        for item in files_empresas:
            item_type = "📁 Pasta" if item.get("folder") else "📄 Arquivo"
            name = item.get('name', 'N/A')
            last_modified = item.get('lastModifiedDateTime', 'N/A')
            is_excel = name.endswith('.xlsx') if name else False
            
            if is_excel:
                excel_files.append(item)
                st.write(f"- ✅ {item_type}: **{name}** (Modificado: {last_modified})")
            else:
                st.write(f"- {item_type}: {name} (Modificado: {last_modified})")
        
        if excel_files:
            st.success(f"✅ Encontrados {len(excel_files)} arquivo(s) Excel (.xlsx)")
        else:
            st.warning("⚠️ Nenhum arquivo Excel (.xlsx) encontrado nesta pasta")
    else:
        st.warning("⚠️ Nenhum item encontrado nesta pasta")
except Exception as e:
    st.error(f"❌ Erro ao listar arquivos: {e}")
    st.code(str(e))

# Testar listagem na pasta de Labs
st.subheader("5. Listando Arquivos em 'Data Analysis/Acumulado de Coletas - Labs'")
try:
    folder_path_labs = "Data Analysis/Acumulado de Coletas - Labs"
    files_labs = sp_connector.list_files(folder_path_labs)
    st.success(f"✅ Encontrados {len(files_labs)} itens em '{folder_path_labs}'")
    
    if files_labs:
        st.write("**Itens encontrados:**")
        excel_files = []
        for item in files_labs:
            item_type = "📁 Pasta" if item.get("folder") else "📄 Arquivo"
            name = item.get('name', 'N/A')
            last_modified = item.get('lastModifiedDateTime', 'N/A')
            is_excel = name.endswith('.xlsx') if name else False
            
            if is_excel:
                excel_files.append(item)
                st.write(f"- ✅ {item_type}: **{name}** (Modificado: {last_modified})")
            else:
                st.write(f"- {item_type}: {name} (Modificado: {last_modified})")
        
        if excel_files:
            st.success(f"✅ Encontrados {len(excel_files)} arquivo(s) Excel (.xlsx)")
        else:
            st.warning("⚠️ Nenhum arquivo Excel (.xlsx) encontrado nesta pasta")
    else:
        st.warning("⚠️ Nenhum item encontrado nesta pasta")
except Exception as e:
    st.error(f"❌ Erro ao listar arquivos: {e}")
    st.code(str(e))

# Testar função get_latest_file_from_sp (implementada localmente para evitar conflito de imports)
st.subheader("6. Testando Busca de Arquivo Mais Recente")

def get_latest_file(sp, folder_path):
    """Busca o arquivo Excel mais recente em uma pasta"""
    files = sp.list_files(folder_path)
    excel_files = [f for f in files if "file" in f and "folder" not in f and f.get("name", "").lower().endswith(".xlsx")]
    if not excel_files:
        return None, None, None
    latest = max(excel_files, key=lambda x: x.get("lastModifiedDateTime", ""))
    return f"{folder_path}/{latest['name']}", latest['name'], latest.get("lastModifiedDateTime", "")

try:
    # Testar para Empresas
    st.write("**Testando para Empresas:**")
    path_emp, name_emp, date_emp = get_latest_file(
        sp_connector, 
        "Data Analysis/Acumulado de Coletas - Empresas"
    )
    if path_emp:
        st.success(f"✅ Arquivo mais recente: **{name_emp}**")
        st.write(f"- Caminho: `{path_emp}`")
        st.write(f"- Data modificação: {date_emp}")
    else:
        st.warning("⚠️ Nenhum arquivo Excel encontrado")
    
    # Testar para Labs
    st.write("**Testando para Labs:**")
    path_labs, name_labs, date_labs = get_latest_file(
        sp_connector,
        "Data Analysis/Acumulado de Coletas - Labs"
    )
    if path_labs:
        st.success(f"✅ Arquivo mais recente: **{name_labs}**")
        st.write(f"- Caminho: `{path_labs}`")
        st.write(f"- Data modificação: {date_labs}")
    else:
        st.warning("⚠️ Nenhum arquivo Excel encontrado")
        
except Exception as e:
    st.error(f"❌ Erro ao testar função: {e}")
    st.code(str(e))

st.markdown("---")
st.caption("Teste concluído. Verifique os resultados acima para identificar problemas na conexão ou caminhos.")
