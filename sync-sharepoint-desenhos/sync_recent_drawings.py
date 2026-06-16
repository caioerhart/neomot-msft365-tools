"""
sync_recent_drawings.py
-----------------------
Copia arquivos .pdf e .dxf modificados nos ultimos N dias de uma pasta
SharePoint (origem) para uma pasta de sincronizacao em outro caminho (destino)
dentro da mesma biblioteca.

Uso:
  python sync_recent_drawings.py --dry-run
  python sync_recent_drawings.py
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from datetime import datetime, timedelta, timezone
from typing import Iterable

import msal
import requests
from dotenv import load_dotenv

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "sim", "y"}


def headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Falha ao obter token: {result.get('error_description', result)}")
    return result["access_token"]


def get_site_id(token: str, hostname: str, site_path: str) -> str:
    url = f"{GRAPH_BASE}/sites/{hostname}:{site_path}"
    response = requests.get(url, headers=headers(token), timeout=30)
    response.raise_for_status()
    return response.json()["id"]


def get_drive_id(token: str, site_id: str, library_name: str) -> str:
    url = f"{GRAPH_BASE}/sites/{site_id}/drives"
    response = requests.get(url, headers=headers(token), timeout=30)
    response.raise_for_status()
    drives = response.json().get("value", [])
    for drive in drives:
        if drive.get("name") == library_name:
            return drive["id"]
    if drives:
        print(
            f"AVISO: biblioteca '{library_name}' nao encontrada. Usando '{drives[0].get('name')}'.",
            file=sys.stderr,
        )
        return drives[0]["id"]
    raise RuntimeError("Nenhum drive encontrado no site.")


def get_item_by_path(token: str, drive_id: str, folder_path: str) -> dict:
    safe_path = folder_path.strip("/")
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{safe_path}"
    response = requests.get(url, headers=headers(token), timeout=30)
    if response.status_code == 404:
        raise RuntimeError(f"Pasta nao encontrada: {folder_path}")
    response.raise_for_status()
    return response.json()


def create_child_folder(token: str, drive_id: str, parent_item_id: str, folder_name: str) -> dict:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent_item_id}/children"
    payload = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "fail",
    }
    response = requests.post(url, headers=headers(token), json=payload, timeout=30)
    if response.status_code in (200, 201):
        return response.json()
    response.raise_for_status()
    return response.json()


def ensure_child_folder(token: str, drive_id: str, parent_item_id: str, folder_name: str) -> dict:
    for child in list_children(token, drive_id, parent_item_id):
        if child.get("name", "").strip().lower() == folder_name.strip().lower() and "folder" in child:
            return child
    return create_child_folder(token, drive_id, parent_item_id, folder_name)


def list_children(token: str, drive_id: str, item_id: str) -> list[dict]:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/children?$top=200"
    items: list[dict] = []
    while url:
        response = requests.get(url, headers=headers(token), timeout=60)
        response.raise_for_status()
        data = response.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def walk_files(token: str, drive_id: str, root_item_id: str, recursive: bool) -> Iterable[dict]:
    stack = [root_item_id]
    while stack:
        current_id = stack.pop()
        for item in list_children(token, drive_id, current_id):
            if "folder" in item:
                if recursive:
                    stack.append(item["id"])
                continue
            if "file" in item:
                yield item


def parse_utc(ts: str) -> datetime:
    return datetime.fromisoformat(ts.replace("Z", "+00:00")).astimezone(timezone.utc)


def should_sync(item: dict, valid_extensions: set[str], min_modified_utc: datetime) -> bool:
    name = item.get("name", "")
    ext = os.path.splitext(name)[1].lower()
    if ext not in valid_extensions:
        return False

    modified_raw = item.get("lastModifiedDateTime")
    if not modified_raw:
        return False

    modified_at = parse_utc(modified_raw)
    return modified_at >= min_modified_utc


def find_destination_file(token: str, drive_id: str, dest_folder_id: str, file_name: str) -> dict | None:
    for child in list_children(token, drive_id, dest_folder_id):
        if "file" in child and child.get("name", "") == file_name:
            return child
    return None


def delete_item(token: str, drive_id: str, item_id: str) -> None:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
    response = requests.delete(url, headers=headers(token), timeout=30)
    if response.status_code not in (204, 200):
        response.raise_for_status()


def copy_item(token: str, drive_id: str, source_item_id: str, destination_folder_id: str, file_name: str) -> None:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{source_item_id}/copy"
    payload = {
        "parentReference": {"driveId": drive_id, "id": destination_folder_id},
        "name": file_name,
    }
    response = requests.post(url, headers=headers(token), json=payload, timeout=30)

    if response.status_code == 202:
        monitor_url = response.headers.get("location") or response.headers.get("Location")
        if monitor_url:
            wait_copy_completion(token, monitor_url)
        return

    if response.status_code not in (200, 201):
        response.raise_for_status()


def wait_copy_completion(token: str, monitor_url: str, timeout_seconds: int = 180) -> None:
    start = time.time()
    while True:
        response = requests.get(
            monitor_url,
            headers={"Authorization": f"Bearer {token}"},
            timeout=30,
        )
        if response.status_code in (200, 201):
            return

        if response.status_code == 202:
            if (time.time() - start) > timeout_seconds:
                raise RuntimeError("Timeout aguardando conclusao da copia no Graph.")
            time.sleep(2)
            continue

        response.raise_for_status()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sincroniza .pdf e .dxf recentes de uma pasta SharePoint para pasta de teste/destino."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Simula o processo sem copiar arquivos.",
    )
    return parser.parse_args()


def main() -> int:
    env_path = os.path.join(os.path.dirname(__file__), ".env")
    load_dotenv(env_path)

    args = parse_args()

    tenant_id = os.getenv("AZURE_TENANT_ID", "").strip()
    client_id = os.getenv("AZURE_CLIENT_ID", "").strip()
    client_secret = os.getenv("AZURE_CLIENT_SECRET", "").strip()

    hostname = os.getenv("SHAREPOINT_HOSTNAME", "neomotelevadores.sharepoint.com").strip()
    site_path = os.getenv("SHAREPOINT_SITE_PATH", "/sites/ArquivosNeomot").strip()
    library_name = os.getenv("SHAREPOINT_LIBRARY_NAME", "Documentos Compartilhados").strip()

    source_folder_path = os.getenv("SOURCE_FOLDER_PATH", "").strip().strip("/")
    dest_parent_folder_path = os.getenv("DEST_PARENT_FOLDER_PATH", "").strip().strip("/")
    dest_sync_folder_name = os.getenv("DEST_SYNC_FOLDER_NAME", "").strip()

    days_back_raw = os.getenv("DAYS_BACK", "7").strip()
    file_extensions_raw = os.getenv("FILE_EXTENSIONS", ".pdf,.dxf")
    recursive = env_bool("RECURSIVE", default=True)
    replace_if_exists = env_bool("REPLACE_IF_EXISTS", default=True)

    if not all([tenant_id, client_id, client_secret]):
        raise RuntimeError("Defina AZURE_TENANT_ID, AZURE_CLIENT_ID e AZURE_CLIENT_SECRET no .env")
    if not source_folder_path:
        raise RuntimeError("Defina SOURCE_FOLDER_PATH no .env")
    if not dest_parent_folder_path:
        raise RuntimeError("Defina DEST_PARENT_FOLDER_PATH no .env")

    try:
        days_back = int(days_back_raw)
    except ValueError as exc:
        raise RuntimeError(f"DAYS_BACK invalido: {days_back_raw}") from exc

    valid_extensions = {
        ext.strip().lower() if ext.strip().startswith(".") else f".{ext.strip().lower()}"
        for ext in file_extensions_raw.split(",")
        if ext.strip()
    }
    if not valid_extensions:
        raise RuntimeError("FILE_EXTENSIONS esta vazio no .env")

    min_modified_utc = datetime.now(timezone.utc) - timedelta(days=days_back)

    print("== Configuracao ==")
    print(f"Origem: {source_folder_path}")
    print(f"Destino pai: {dest_parent_folder_path}")
    if dest_sync_folder_name:
        print(f"Subpasta destino: {dest_sync_folder_name}")
    else:
        print("Subpasta destino: (nao usada; copia direto na raiz do destino)")
    print(f"Extensoes: {sorted(valid_extensions)}")
    print(f"Ultimos dias: {days_back}")
    print(f"Recursivo: {recursive}")
    print(f"Substituir se existir: {replace_if_exists}")
    print(f"Dry run: {args.dry_run}")

    token = get_access_token(tenant_id, client_id, client_secret)
    site_id = get_site_id(token, hostname, site_path)
    drive_id = get_drive_id(token, site_id, library_name)

    source_item = get_item_by_path(token, drive_id, source_folder_path)
    dest_parent_item = get_item_by_path(token, drive_id, dest_parent_folder_path)
    if dest_sync_folder_name:
        dest_target_item = ensure_child_folder(token, drive_id, dest_parent_item["id"], dest_sync_folder_name)
    else:
        dest_target_item = dest_parent_item

    print("\n== Filtro de arquivos ==")
    candidates = []
    for item in walk_files(token, drive_id, source_item["id"], recursive=recursive):
        if should_sync(item, valid_extensions, min_modified_utc):
            candidates.append(item)

    print(f"Arquivos elegiveis: {len(candidates)}")
    if not candidates:
        print("Nenhum arquivo para copiar.")
        return 0

    print("\n== Copia ==")
    copied = 0
    skipped = 0
    failed = 0

    for idx, item in enumerate(candidates, start=1):
        name = item.get("name", "")
        modified = item.get("lastModifiedDateTime", "")
        print(f"[{idx}/{len(candidates)}] {name} (modificado: {modified})")

        try:
            existing = find_destination_file(token, drive_id, dest_target_item["id"], name)
            if existing and not replace_if_exists:
                print("  - ja existe no destino, pulando")
                skipped += 1
                continue

            if existing and replace_if_exists:
                if args.dry_run:
                    print("  - dry-run: removeria versao existente")
                else:
                    delete_item(token, drive_id, existing["id"])
                    print("  - versao anterior removida")

            if args.dry_run:
                print("  - dry-run: copiaria arquivo")
                copied += 1
                continue

            copy_item(token, drive_id, item["id"], dest_target_item["id"], name)
            print("  - copiado")
            copied += 1

        except Exception as exc:  # noqa: BLE001
            print(f"  - erro: {exc}", file=sys.stderr)
            failed += 1

    print("\n== Resumo ==")
    print(f"Copiados/simulados: {copied}")
    print(f"Pulados: {skipped}")
    print(f"Falhas: {failed}")

    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
