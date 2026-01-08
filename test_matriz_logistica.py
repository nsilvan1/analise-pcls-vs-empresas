# test_matriz_logistica.py
# Script para testar acesso à planilha CONSULTA MATRIZ LOGISTICA no SharePoint

import streamlit as st
from sp_connector import get_sp_connector, SPConnector
import pandas as pd

def test_sharepoint_access():
    """Testa diferentes caminhos para acessar a planilha no SharePoint"""

    print("=" * 60)
    print("TESTE DE ACESSO AO SHAREPOINT - MATRIZ LOGISTICA")
    print("=" * 60)

    # Carregar configurações do secrets
    try:
        # Simular secrets do Streamlit
        import toml
        secrets_path = ".streamlit/secrets.toml"
        with open(secrets_path, 'r', encoding='utf-8') as f:
            secrets = toml.load(f)

        graph_cfg = secrets.get("graph", {})
        tenant_id = graph_cfg.get("tenant_id")
        client_id = graph_cfg.get("client_id")
        client_secret = graph_cfg.get("client_secret")
        hostname = graph_cfg.get("hostname")

        print(f"\n[INFO] Tenant ID: {tenant_id[:8]}..." if tenant_id else "[ERRO] Tenant ID não encontrado")
        print(f"[INFO] Client ID: {client_id[:8]}..." if client_id else "[ERRO] Client ID não encontrado")
        print(f"[INFO] Hostname: {hostname}" if hostname else "[ERRO] Hostname não encontrado")

    except Exception as e:
        print(f"[ERRO] Não foi possível carregar secrets: {e}")
        return

    # Caminhos possíveis baseados na URL fornecida
    # URL: https://synviagroup.sharepoint.com/sites/comercialtoxicologico/Administrativo/...
    # Pasta: /sites/comercialtoxicologico/Administrativo/TLMK ATIVO/TLMK ATIVO - TOXICOLOGICO/A Planilha Logistica

    # O hostname nos secrets é synviagroup-my.sharepoint.com (OneDrive)
    # Mas o site SharePoint é synviagroup.sharepoint.com (sem -my)
    # Vamos testar ambos

    sharepoint_hostname = "synviagroup.sharepoint.com"  # SharePoint real
    onedrive_hostname = hostname  # Do secrets (synviagroup-my.sharepoint.com)

    print(f"\n[INFO] SharePoint hostname: {sharepoint_hostname}")
    print(f"[INFO] OneDrive hostname (secrets): {onedrive_hostname}")

    test_configs = [
        {
            "name": "SharePoint Site - comercialtoxicologico (Administrativo)",
            "hostname": sharepoint_hostname,
            "site_path": "sites/comercialtoxicologico",
            "library_name": "Administrativo",
            "file_path": "TLMK ATIVO/TLMK ATIVO - TOXICOLOGICO/A Planilha Logistica/CONSULTA MATRIZ LOGISTICA.1.xlsx"
        },
        {
            "name": "SharePoint Site - comercialtoxicologico (Documents)",
            "hostname": sharepoint_hostname,
            "site_path": "sites/comercialtoxicologico",
            "library_name": "Documents",
            "file_path": "CONSULTA MATRIZ LOGISTICA.1.xlsx"
        },
        {
            "name": "SharePoint Site - comercialtoxicologico (Shared Documents)",
            "hostname": sharepoint_hostname,
            "site_path": "sites/comercialtoxicologico",
            "library_name": "Shared Documents",
            "file_path": "CONSULTA MATRIZ LOGISTICA.1.xlsx"
        },
    ]

    print("\n" + "-" * 60)
    print("TESTANDO DIFERENTES CONFIGURAÇÕES...")
    print("-" * 60)

    for i, config in enumerate(test_configs, 1):
        print(f"\n[TESTE {i}] {config['name']}")
        print(f"  Site Path: {config['site_path']}")
        print(f"  Library: {config['library_name']}")
        print(f"  File Path: {config['file_path']}")

        try:
            connector = SPConnector(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret,
                hostname=config.get('hostname', hostname),
                site_path=config['site_path'],
                library_name=config['library_name']
            )

            # Tentar listar arquivos na raiz primeiro
            print(f"  > Listando arquivos na raiz da biblioteca...")
            try:
                files = connector.list_files("")
                print(f"  > Encontrados {len(files)} itens na raiz:")
                for f in files[:10]:  # Mostrar apenas os 10 primeiros
                    tipo = "📁" if f.get("folder") else "📄"
                    print(f"    {tipo} {f.get('name')}")
                if len(files) > 10:
                    print(f"    ... e mais {len(files) - 10} itens")
            except Exception as e:
                print(f"  [ERRO] Não foi possível listar: {e}")

            # Tentar baixar o arquivo
            print(f"  > Tentando baixar arquivo...")
            try:
                df = connector.read_excel(config['file_path'])
                print(f"  [SUCESSO] Arquivo carregado! {len(df)} linhas, {len(df.columns)} colunas")
                print(f"  Colunas: {df.columns.tolist()}")
                return connector, config  # Retorna a configuração que funcionou
            except FileNotFoundError:
                print(f"  [ERRO] Arquivo não encontrado")
            except Exception as e:
                print(f"  [ERRO] {type(e).__name__}: {e}")

        except Exception as e:
            print(f"  [ERRO] Falha ao criar connector: {e}")

    print("\n" + "=" * 60)
    print("TENTANDO DESCOBRIR BIBLIOTECAS DISPONÍVEIS...")
    print("=" * 60)

    try:
        connector = SPConnector(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            hostname="synviagroup.sharepoint.com",  # SharePoint real (sem -my)
            site_path="sites/comercialtoxicologico",
            library_name="Documents"  # Qualquer uma para inicializar
        )

        # Listar drives/bibliotecas do site
        import requests
        GRAPH = "https://graph.microsoft.com/v1.0"
        site_id = connector._site_id()
        print(f"\n[INFO] Site ID: {site_id}")

        url = f"{GRAPH}/sites/{site_id}/drives"
        r = requests.get(url, headers=connector._headers(), timeout=30)
        r.raise_for_status()
        drives = r.json().get("value", [])

        print(f"\n[INFO] Bibliotecas encontradas no site:")
        for d in drives:
            print(f"  - {d.get('name')} (ID: {d.get('id')[:20]}...)")
            print(f"    Tipo: {d.get('driveType')}")
            print(f"    Web URL: {d.get('webUrl')}")

    except Exception as e:
        print(f"[ERRO] Não foi possível listar bibliotecas: {e}")

    return None, None


def test_folder_navigation():
    """Testa navegação em pastas específicas"""

    print("\n" + "=" * 60)
    print("TESTE DE NAVEGAÇÃO EM PASTAS")
    print("=" * 60)

    import toml
    secrets_path = ".streamlit/secrets.toml"
    with open(secrets_path, 'r', encoding='utf-8') as f:
        secrets = toml.load(f)

    graph_cfg = secrets.get("graph", {})

    # Testar com biblioteca "Administrativo"
    connector = SPConnector(
        tenant_id=graph_cfg.get("tenant_id"),
        client_id=graph_cfg.get("client_id"),
        client_secret=graph_cfg.get("client_secret"),
        hostname="synviagroup.sharepoint.com",  # SharePoint real (sem -my)
        site_path="sites/comercialtoxicologico",
        library_name="Administrativo"
    )

    # Navegar pela estrutura de pastas
    folders_to_check = [
        "",
        "TLMK ATIVO",
        "TLMK ATIVO/TLMK ATIVO - TOXICOLOGICO",
        "TLMK ATIVO/TLMK ATIVO - TOXICOLOGICO/A Planilha Logistica",
    ]

    for folder in folders_to_check:
        print(f"\n[PASTA] '{folder or '(raiz)'}'")
        try:
            files = connector.list_files(folder)
            print(f"  Encontrados {len(files)} itens:")
            for f in files[:15]:
                tipo = "📁" if f.get("folder") else "📄"
                nome = f.get('name')
                print(f"    {tipo} {nome}")
                if "MATRIZ" in nome.upper() or "LOGISTICA" in nome.upper():
                    print(f"       ^^^ POSSÍVEL ARQUIVO ALVO!")
        except Exception as e:
            print(f"  [ERRO] {e}")


if __name__ == "__main__":
    # Executar testes
    connector, config = test_sharepoint_access()

    if connector is None:
        print("\n[INFO] Tentando navegação em pastas...")
        test_folder_navigation()
    else:
        print(f"\n[SUCESSO] Configuração funcional encontrada!")
        print(f"  Use: site_path='{config['site_path']}', library_name='{config['library_name']}'")
        print(f"  Caminho do arquivo: '{config['file_path']}'")
