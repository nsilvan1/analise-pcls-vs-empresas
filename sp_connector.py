# sp_connector.py
import io, time, requests, msal, pandas as pd
from urllib.parse import quote
import streamlit as st

GRAPH = "https://graph.microsoft.com/v1.0"

class SPConnector:
    """
    Conecta no SharePoint/OneDrive via Microsoft Graph (app-only).
    Suporta:
      - SharePoint Site: hostname + site_path + library_name
      - OneDrive do usuário: user_upn (ex: "washington.gouvea@synvia.com")
    Caminho do arquivo:
      - OneDrive: RELATIVO a Documents (ex: "Pasta/arquivo.xlsx")
        (aceita tb /personal/<upn>/Documents/... que será normalizado)
      - SharePoint: RELATIVO à biblioteca (ex: "Pasta/arquivo.xlsx")
        (aceita tb server-relative /sites/<site>/<lib>/... que será normalizado)
    """

    def __init__(self, tenant_id, client_id, client_secret,
                 hostname=None, site_path=None, library_name=None, user_upn=None,
                 access_token=None):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret

        self.hostname = hostname or ""
        self.site_path = site_path or ""
        self.library_name = library_name or ""
        self.user_upn = user_upn or ""          # se presente, opera em OneDrive
        self._delegated_token = access_token     # se presente, usa token do usuário (delegated)

        self._app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            client_credential=self.client_secret,
        )
        self._tok = None
        self._exp = 0
        self._site_id_cache = None
        self._drive_id_cache = None

    # -------- Auth --------
    def _token(self):
        # Se houver token delegado (do usuário logado), usa-o diretamente
        if self._delegated_token:
            return self._delegated_token
        now = time.time()
        if self._tok and now < self._exp:
            return self._tok
        res = self._app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in res:
            raise RuntimeError(res.get("error_description") or res)
        self._tok = res["access_token"]
        self._exp = now + int(res.get("expires_in", 3600)) - 60
        return self._tok

    def _headers(self):
        return {"Authorization": f"Bearer {self._token()}"}

    # -------- Modo --------
    @property
    def is_onedrive(self) -> bool:
        return bool(self.user_upn)

    # -------- Descoberta (apenas p/ SharePoint Site) --------
    def _site_id(self):
        if self.is_onedrive:
            return None
        if self._site_id_cache:
            return self._site_id_cache
        url = f"{GRAPH}/sites/{self.hostname}:/{self.site_path}"
        r = requests.get(url, headers=self._headers(), timeout=30)
        r.raise_for_status()
        self._site_id_cache = r.json()["id"]
        return self._site_id_cache

    def _drive_id(self):
        if self.is_onedrive:
            return None
        if self._drive_id_cache:
            return self._drive_id_cache
        url = f"{GRAPH}/sites/{self._site_id()}/drives"
        r = requests.get(url, headers=self._headers(), timeout=30)
        r.raise_for_status()
        drives = r.json().get("value", [])
        for d in drives:
            if d.get("name", "").lower() == self.library_name.lower():
                self._drive_id_cache = d["id"]
                return self._drive_id_cache
        for d in drives:
            if d.get("driveType") == "documentLibrary":
                self._drive_id_cache = d["id"]
                return self._drive_id_cache
        raise RuntimeError(f"Biblioteca '{self.library_name}' não encontrada em {self.site_path}")

    # -------- Normalização de caminho --------
    def normalize_path(self, path: str) -> str:
        """
        Retorna o caminho relativo correto ao "root" usado em cada modo:
          - OneDrive: relativo a Documents/
          - SharePoint: relativo à biblioteca
        Aceita caminhos server-relative e normaliza.
        """
        if not path:
            raise ValueError("Caminho vazio.")
        path = path.strip()

        if self.is_onedrive:
            # Aceita: "Pasta/arquivo.xlsx" OU "/personal/<upn>/Documents/Pasta/arquivo.xlsx"
            p = path.strip()
            if p.startswith("/"):
                # Se vier server-relative, tenta cortar a partir de /Documents/
                marker = "/documents/"
                idx = p.lower().find(marker)
                if idx != -1:
                    rel = p[idx + len(marker):]
                else:
                    rel = p.lstrip("/")
            else:
                rel = p
            # Garanta que está debaixo de Documents/
            if not rel.lower().startswith("documents/"):
                rel = f"Documents/{rel}"
            return rel
        else:
            # SharePoint site
            if path.startswith("/"):
                prefix = f"/{self.site_path}/{self.library_name}/"
                if not path.startswith(prefix):
                    raise ValueError(
                        "file_path server-relative não bate com site/biblioteca dos secrets.\n"
                        f"Esperado prefixo: {prefix}\nRecebido: {path}"
                    )
                return path[len(prefix):]
            return path

    # -------- Download / Upload --------
    def download(self, path: str) -> bytes:
        rel = quote(self.normalize_path(path), safe="/")
        if self.is_onedrive:
            base = f"{GRAPH}/me/drive" if self._delegated_token else f"{GRAPH}/users/{self.user_upn}/drive"
            url = f"{base}/root:/{rel}:/content"
        else:
            url = f"{GRAPH}/drives/{self._drive_id()}/root:/{rel}:/content"
        r = requests.get(url, headers=self._headers(), timeout=180)
        if r.status_code == 404:
            raise FileNotFoundError(path)
        r.raise_for_status()
        return r.content

    def upload_small(self, path: str, content: bytes, overwrite: bool = True):
        rel = quote(self.normalize_path(path), safe="/")
        params = {"@microsoft.graph.conflictBehavior": "replace" if overwrite else "fail"}
        if self.is_onedrive:
            base = f"{GRAPH}/me/drive" if self._delegated_token else f"{GRAPH}/users/{self.user_upn}/drive"
            url = f"{base}/root:/{rel}:/content"
        else:
            url = f"{GRAPH}/drives/{self._drive_id()}/root:/{rel}:/content"
        r = requests.put(url, headers=self._headers(), params=params, data=content, timeout=300)
        r.raise_for_status()
        return r.json()

    def delete_file(self, path: str) -> bool:
        """Exclui um arquivo ou pasta (envia para a lixeira do OneDrive/SharePoint)."""
        rel = quote(self.normalize_path(path), safe="/")
        if self.is_onedrive:
            base = f"{GRAPH}/me/drive" if self._delegated_token else f"{GRAPH}/users/{self.user_upn}/drive"
            url = f"{base}/root:/{rel}"
        else:
            url = f"{GRAPH}/drives/{self._drive_id()}/root:/{rel}"
        r = requests.delete(url, headers=self._headers(), timeout=60)
        if r.status_code in (200, 204):
            return True
        if r.status_code == 404:
            return False
        r.raise_for_status()
        return True

    # -------- Conveniências DataFrame --------
    def read_excel(self, path: str, **kw) -> pd.DataFrame:
        return pd.read_excel(io.BytesIO(self.download(path)), **kw)

    def read_csv(self, path: str, **kw) -> pd.DataFrame:
        return pd.read_csv(io.BytesIO(self.download(path)), **kw)

    def write_excel(self, df: pd.DataFrame, path: str, overwrite: bool = True):
        bio = io.BytesIO()
        df.to_excel(bio, index=False)
        return self.upload_small(path, bio.getvalue(), overwrite=overwrite)

    def write_csv(self, df: pd.DataFrame, path: str, overwrite: bool = True):
        bio = io.StringIO()
        df.to_csv(bio, index=False, encoding='utf-8-sig')
        return self.upload_small(path, bio.getvalue().encode('utf-8'), overwrite=overwrite)

    # -------- Métodos específicos para nosso projeto --------
    def upload_file(self, path: str, content: bytes, overwrite: bool = True):
        """Upload de arquivo genérico"""
        return self.upload_small(path, content, overwrite)

    def file_exists(self, path: str) -> bool:
        """Verifica se um arquivo existe"""
        try:
            self.download(path)
            return True
        except FileNotFoundError:
            return False
        except Exception:
            return False

    def list_files(self, folder_path: str = ""):
        """Lista arquivos em uma pasta"""
        rel = quote(self.normalize_path(folder_path), safe="/") if folder_path else ""
        if self.is_onedrive:
            base = f"{GRAPH}/me/drive" if self._delegated_token else f"{GRAPH}/users/{self.user_upn}/drive"
            url = f"{base}/root:/{rel}:/children"
        else:
            url = f"{GRAPH}/drives/{self._drive_id()}/root:/{rel}:/children"
        
        r = requests.get(url, headers=self._headers(), timeout=30)
        r.raise_for_status()
        return r.json().get("value", [])

    def create_folder(self, folder_path: str):
        """Cria uma pasta"""
        rel = quote(self.normalize_path(folder_path), safe="/")
        folder_data = {
            "name": folder_path.split("/")[-1],
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }
        
        parent_path = "/".join(folder_path.split("/")[:-1]) if "/" in folder_path else ""
        parent_rel = quote(self.normalize_path(parent_path), safe="/") if parent_path else ""
        if self.is_onedrive:
            base = f"{GRAPH}/me/drive" if self._delegated_token else f"{GRAPH}/users/{self.user_upn}/drive"
            url = f"{base}/root:/{parent_rel}:/children" if parent_rel else f"{base}/root/children"
        else:
            url = f"{GRAPH}/drives/{self._drive_id()}/root:/{parent_rel}:/children" if parent_rel else f"{GRAPH}/drives/{self._drive_id()}/root/children"
        
        r = requests.post(url, headers=self._headers(), json=folder_data, timeout=30)
        r.raise_for_status()
        return r.json()


def get_sp_connector():
    """Cria uma instância do SPConnector usando as configurações do secrets.
    Preferência: SharePoint Site (Sites.ReadWrite.All). Fallback: OneDrive por UPN."""
    try:
        graph_cfg = st.secrets.get("graph", {})
        tenant_id = graph_cfg.get("tenant_id")
        client_id = graph_cfg.get("client_id")
        client_secret = graph_cfg.get("client_secret")
        hostname = graph_cfg.get("hostname")
        site_path = graph_cfg.get("site_path")
        library_name = graph_cfg.get("library_name", "Documents")
        user_upn_secret = st.secrets.get("onedrive", {}).get("user_upn")
        # Se o usuário estiver autenticado, preferir token delegado
        access_token = st.session_state.get("access_token")
        session_user = st.session_state.get("user_info", {})
        user_upn_session = session_user.get("mail") or session_user.get("userPrincipalName")

        if not tenant_id or not client_id or not client_secret:
            raise RuntimeError("Configurações do Microsoft Graph ausentes em secrets.")

        # 1) Se houver token delegado, operar no OneDrive do usuário logado via /me/drive
        if access_token and user_upn_session:
            return SPConnector(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret,
                user_upn=user_upn_session,
                access_token=access_token
            )

        # 2) Caso contrário, se houver hostname + site_path, usar modo SharePoint Site (app-only)
        if hostname and site_path:
            # Se site_path vier como URL completa, extrair apenas o path após o hostname
            if site_path.startswith("http://") or site_path.startswith("https://"):
                try:
                    # Remover protocolo e hostname
                    prefix = f"https://{hostname}/"
                    if site_path.lower().startswith(prefix.lower()):
                        site_path = site_path[len(prefix):]
                except Exception:
                    pass
            return SPConnector(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret,
                hostname=hostname,
                site_path=site_path.strip("/"),
                library_name=library_name
            )

        # Fallback para OneDrive do usuário
        if user_upn_secret:
            return SPConnector(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret,
                user_upn=user_upn_secret
            )

        # Se nada disponível
        raise RuntimeError("Secrets sem dados suficientes: informe hostname/site_path ou onedrive.user_upn.")
    except Exception as e:
        st.error(f"Erro ao criar SPConnector: {e}")
        return None
