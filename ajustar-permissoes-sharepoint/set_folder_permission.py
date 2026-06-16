"""
set_folder_permission.py
------------------------
Atribui permissões de pasta no SharePoint a partir de um arquivo CSV,
usando somente o Microsoft Graph API.

Formato do CSV (separador vírgula, com cabeçalho):
  email,pasta,permissao
  qualidade@neomot.com,INDUSTRIAL,write
  pcp@neomot.com,INDUSTRIAL,write

Colunas:
  email      – e-mail do grupo ou usuário M365
  pasta      – caminho relativo dentro da biblioteca (ex.: INDUSTRIAL ou A/B)
  permissao  – 'write' (Editor) ou 'read' (Leitor)

Estratégia por linha:
  1. Concede a permissão na pasta alvo
  2. Concede 'read' nas pastas pai para que o grupo navegue até ela
  3. Remove a permissão do grupo dos irmãos de cada pai, restringindo
     a visibilidade apenas ao caminho alvo

Permissões necessárias no App Registration (Azure Portal):
  Microsoft Graph → Application → Sites.ReadWrite.All

Uso:
  python set_folder_permission.py [arquivo.csv]
  (padrão: permissions.csv no mesmo diretório)

Variáveis de ambiente (.env):
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
  SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH, SHAREPOINT_LIBRARY_NAME
"""

import csv
import os
import sys

import msal
import requests
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

TENANT_ID     = os.getenv("AZURE_TENANT_ID", "")
CLIENT_ID     = os.getenv("AZURE_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "")

SHAREPOINT_HOSTNAME = os.getenv("SHAREPOINT_HOSTNAME", "neomotelevadores.sharepoint.com")
SITE_PATH           = os.getenv("SHAREPOINT_SITE_PATH", "/sites/ArquivosNeomot")
LIBRARY_NAME        = os.getenv("SHAREPOINT_LIBRARY_NAME", "Documentos Compartilhados")

DEFAULT_CSV = os.path.join(os.path.dirname(__file__), "permissions.csv")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


# ---------------------------------------------------------------------------
# Autenticação
# ---------------------------------------------------------------------------

def get_access_token() -> str:
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(
            f"Falha ao obter token: {result.get('error_description', result)}"
        )
    return result["access_token"]


# ---------------------------------------------------------------------------
# Helpers – site / drive
# ---------------------------------------------------------------------------

def get_site_id(token: str) -> str:
    url = f"{GRAPH_BASE}/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}"
    resp = requests.get(url, headers=_auth(token))
    resp.raise_for_status()
    site_id = resp.json()["id"]
    print(f"  Site ID : {site_id}")
    return site_id


def get_site_url(token: str, site_id: str) -> str:
    """Retorna a URL absoluta do site (ex.: https://host/sites/xxx)."""
    url = f"{GRAPH_BASE}/sites/{site_id}?$select=webUrl"
    resp = requests.get(url, headers=_auth(token))
    resp.raise_for_status()
    web_url = resp.json()["webUrl"]
    print(f"  Site URL: {web_url}")
    return web_url


def get_drive_id(token: str, site_id: str) -> str:
    url = f"{GRAPH_BASE}/sites/{site_id}/drives"
    resp = requests.get(url, headers=_auth(token))
    resp.raise_for_status()
    drives = resp.json().get("value", [])
    for drive in drives:
        if drive.get("name") == LIBRARY_NAME:
            print(f"  Drive ID: {drive['id']}  ({drive['name']})")
            return drive["id"]
    if drives:
        print(
            f"  AVISO: biblioteca '{LIBRARY_NAME}' não encontrada; "
            f"usando primeiro drive: {drives[0]['name']}"
        )
        return drives[0]["id"]
    raise RuntimeError("Nenhum drive encontrado no site.")


# ---------------------------------------------------------------------------
# Helpers – pasta alvo
# ---------------------------------------------------------------------------

def get_folder_item_id(token: str, drive_id: str, folder_path: str) -> str:
    """Retorna o item ID da pasta no drive."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{folder_path}"
    resp = requests.get(url, headers=_auth(token))
    if resp.status_code == 404:
        raise RuntimeError(
            f"Pasta '{folder_path}' não encontrada na biblioteca '{LIBRARY_NAME}'."
        )
    resp.raise_for_status()
    item = resp.json()
    item_id = item["id"]
    print(f"  Item ID da pasta '{folder_path}': {item_id}")
    return item_id


# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------

def load_csv(path: str) -> list[dict]:
    """
    Lê o CSV e retorna lista de dicts com chaves: email, pasta, permissao.
    Ignora linhas vazias e com '#'.
    """
    rows = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for i, row in enumerate(reader, start=2):
            email     = row.get("email", "").strip()
            pasta     = row.get("pasta", "").strip().strip("/")
            permissao = row.get("permissao", "write").strip().lower() or "write"
            if not email or email.startswith("#") or not pasta:
                continue
            if permissao not in ("read", "write"):
                print(f"  AVISO linha {i}: permissão '{permissao}' inválida, usando 'write'.")
                permissao = "write"
            rows.append({"email": email, "pasta": pasta, "permissao": permissao})
    return rows


# ---------------------------------------------------------------------------
# Permissões via Graph API
# ---------------------------------------------------------------------------

def list_current_permissions(token: str, drive_id: str, item_id: str) -> list:
    """Lista permissões do item via Graph API."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/permissions"
    resp = requests.get(url, headers=_auth(token))
    resp.raise_for_status()
    return resp.json().get("value", [])


def _invite(
    token: str, drive_id: str, item_id: str,
    roles: list, label: str, email: str
) -> None:
    """Concede roles ao email no item via /invite."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/invite"
    payload = {
        "requireSignIn": True,
        "sendInvitation": False,
        "roles": roles,
        "recipients": [{"email": email}],
        "message": "",
    }
    resp = requests.post(url, headers=_auth(token), json=payload)
    if not resp.ok:
        print(f"  ERRO /invite HTTP {resp.status_code} em '{label}': {resp.text}", file=sys.stderr)
        resp.raise_for_status()
    granted = resp.json().get("value", [])
    for p in granted:
        r = p.get("roles", [])
        grantee = p.get("grantedToV2", {}) or p.get("grantedTo", {})
        name = (grantee.get("group") or grantee.get("user") or {}).get("displayName", email)
        print(f"    ✓ [{label}] {name} → {r}")


def _find_group_permission_id(
    token: str, drive_id: str, item_id: str, email: str
) -> str | None:
    """
    Encontra o ID da permissão do email no item.
    Retorna None se não encontrar permissão direta.
    """
    perms = list_current_permissions(token, drive_id, item_id)
    for p in perms:
        grantee = p.get("grantedToV2", {}) or p.get("grantedTo", {})
        for key in ("group", "user", "siteGroup"):
            principal = grantee.get(key, {})
            if principal.get("email", "").lower() == email.lower():
                return p["id"]
    return None


def _remove_permission(
    token: str, drive_id: str, item_id: str, perm_id: str, label: str
) -> None:
    """Remove uma permissão específica do item."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/permissions/{perm_id}"
    resp = requests.delete(url, headers=_auth(token))
    if resp.status_code == 204:
        print(f"    ✓ Permissão removida de '{label}'")
    elif resp.status_code == 403:
        # Permissão herdada — ignorar (o item não tem permissão direta)
        print(f"    ~ Sem permissão direta em '{label}' (herdada do pai, ignorando)")
    else:
        print(f"    ! ERRO ao remover de '{label}': HTTP {resp.status_code} {resp.text}")


def _list_children(token: str, drive_id: str, item_id: str) -> list:
    """Lista todos os filhos diretos de um item (pastas e arquivos)."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/children"
    children = []
    while url:
        resp = requests.get(url, headers=_auth(token))
        resp.raise_for_status()
        data = resp.json()
        children.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return children


def grant_editor_permission(
    token: str, drive_id: str, item_id: str,
    email: str, target_folder: str, permissao: str
) -> None:
    """
    Para um par (email, pasta):
      1. Concede a permissão (write ou read) na pasta alvo
      2. Para cada pasta pai: concede 'read' e remove o email dos irmãos
    """
    parts = target_folder.strip("/").split("/")

    print(f"  Concedendo '{permissao}' em '{target_folder}' …")
    _invite(token, drive_id, item_id, [permissao], target_folder, email)

    for depth in range(len(parts) - 1, 0, -1):
        parent_path = "/".join(parts[:depth])
        keep_child  = parts[depth]

        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{parent_path}"
        resp = requests.get(url, headers=_auth(token))
        if resp.status_code == 404:
            print(f"  AVISO: pasta pai '{parent_path}' não encontrada.")
            continue
        resp.raise_for_status()
        parent_id = resp.json()["id"]

        print(f"  Concedendo 'read' em '{parent_path}' …")
        _invite(token, drive_id, parent_id, ["read"], parent_path, email)

        print(f"  Removendo acesso dos irmãos de '{keep_child}' em '{parent_path}' …")
        children = _list_children(token, drive_id, parent_id)
        for child in children:
            child_name = child.get("name", "")
            child_id   = child["id"]
            if child_name.lower() == keep_child.lower():
                continue
            perm_id = _find_group_permission_id(token, drive_id, child_id, email)
            if perm_id:
                _remove_permission(token, drive_id, child_id, perm_id, child_name)
            else:
                print(f"    ~ '{child_name}' já sem permissão do grupo")


def _auth(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


# ---------------------------------------------------------------------------
# Relatório de permissões atuais
# ---------------------------------------------------------------------------

def print_permissions(perms: list, label: str = "") -> None:
    if label:
        print(f"\n  {label}")
    if not perms:
        print("    (nenhuma permissão listada)")
        return
    for p in perms:
        roles     = p.get("roles", [])
        inherited = p.get("inheritedFrom", {})
        grantee   = (
            p.get("grantedToV2", {})
            or p.get("grantedTo", {})
        )
        principal = (
            grantee.get("group")
            or grantee.get("user")
            or grantee.get("siteGroup")
            or {}
        )
        name  = principal.get("displayName", "—")
        email = principal.get("email", "")
        inh   = f"  [herdado de: {inherited.get('path', '?')}]" if inherited else ""
        print(f"    • {name} ({email})  →  {roles}{inh}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    csv_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_CSV

    # Validação
    missing = [
        var for var, val in [
            ("AZURE_TENANT_ID",     TENANT_ID),
            ("AZURE_CLIENT_ID",     CLIENT_ID),
            ("AZURE_CLIENT_SECRET", CLIENT_SECRET),
        ] if not val
    ]
    if missing:
        print(f"ERRO: Variáveis não definidas: {', '.join(missing)}", file=sys.stderr)
        sys.exit(1)

    if not os.path.exists(csv_path):
        print(f"ERRO: Arquivo CSV não encontrado: {csv_path}", file=sys.stderr)
        sys.exit(1)

    rows = load_csv(csv_path)
    if not rows:
        print("ERRO: CSV vazio ou sem linhas válidas.", file=sys.stderr)
        sys.exit(1)

    print("=" * 60)
    print(f"  CSV        : {csv_path}  ({len(rows)} regras)")
    print(f"  Biblioteca : {LIBRARY_NAME}")
    print(f"  Site       : {SHAREPOINT_HOSTNAME}{SITE_PATH}")
    print("=" * 60)

    print("\n[1/3] Obtendo token de acesso …")
    token = get_access_token()
    print("      ✓ Token obtido.")

    print("\n[2/3] Resolvendo site e drive …")
    site_id  = get_site_id(token)
    _        = get_site_url(token, site_id)
    drive_id = get_drive_id(token, site_id)

    print(f"\n[3/3] Processando {len(rows)} regra(s) do CSV …")
    errors = []
    for i, row in enumerate(rows, start=1):
        email     = row["email"]
        pasta     = row["pasta"]
        permissao = row["permissao"]
        print(f"\n  [{i}/{len(rows)}] {email}  →  {pasta}  ({permissao})")
        try:
            item_id = get_folder_item_id(token, drive_id, pasta)
            grant_editor_permission(token, drive_id, item_id, email, pasta, permissao)
            print(f"        ✓ Concluído.")
        except Exception as exc:
            msg = f"ERRO em '{pasta}' para '{email}': {exc}"
            print(f"        ✗ {msg}", file=sys.stderr)
            errors.append(msg)

    print("\n" + "=" * 60)
    if errors:
        print(f"  Concluído com {len(errors)} erro(s):")
        for e in errors:
            print(f"    ✗ {e}")
    else:
        print(f"  Todas as {len(rows)} regras aplicadas com sucesso.")
    print("=" * 60)


if __name__ == "__main__":
    main()
