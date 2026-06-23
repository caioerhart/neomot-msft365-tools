"""
Extrai dados do Oracle e grava na tabela do Excel 'db_material_rv' no SharePoint.

Uso (append):
  python oracle_to_db_material_rv.py

Uso (replace):
  python oracle_to_db_material_rv.py --data-inicial 01/06/2026 --data-final 23/06/2026 --mode replace

Variaveis de ambiente esperadas no .env deste diretorio:
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
  SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH, SHAREPOINT_LIBRARY_NAME

Variaveis opcionais:
  ORACLE_HOST, ORACLE_PORT, ORACLE_SID, ORACLE_USER, ORACLE_PASSWORD
  ORACLE_CLIENT_LIB_DIR, ORACLE_CLIENT_CONFIG_FILE
  EXCEL_FILE_PATH, EXCEL_TABLE_NAME
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from datetime import datetime, timedelta
from decimal import Decimal
from typing import Any
from urllib.parse import quote

import msal
import oracledb
import requests
from dotenv import load_dotenv

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
EXCEL_EPOCH = datetime(1899, 12, 30)
SOURCE_COLUMN_COUNT = 13
USER_FORMULA_SOURCE_COLUMN = "L"
FORMULA_UPDATE_BATCH_SIZE = 500

SQL = """
SELECT
    TEMPRESAS.COD_EMP AS "divisao_venda",
    TCLIENTES.COD_CLI || ' - ' || TCLIENTES.DESCRICAO AS "cliente",
    TNFS_SAIDA.NUM_NF AS "num_nf",
    TNFS_SAIDA.DT_EMIS AS "dt_emissao_nf",
    TNFS_SAIDA.SIT_NF AS "sit_nf",
    TITENS_COMERCIAL.COD_ITEM || ' - ' || TITENS_NFS.DESCRICAO AS "item",
    TITENS_NFS.QTDE AS "qtd",
    TITENS_NFS.PRECO_UNIT AS "vlr_un",
    TITENS_NFS.PRECO_UNIT * TITENS_NFS.QTDE AS "vlr_total_item",
    TITENS_NFS.VLR_CONTABIL AS "vlt_total_nf",
    TTIPOS_NF.COD_TP_NF AS "cod_tp_nf",
    TPEDIDOS_VENDA.USUARIO AS "user_sistema",
    TDIVISOES_VENDAS.COD_DIVD || ' - ' || TDIVISOES_VENDAS.DESCRICAO AS "divisao_vendas",
    NULL AS "user",
    NULL AS "status_bi"
FROM TEMPRESAS
JOIN TNFS_SAIDA ON TEMPRESAS.ID = TNFS_SAIDA.EMPR_ID
JOIN TCLIENTES ON TCLIENTES.ID = TNFS_SAIDA.CLI_ID
JOIN TITENS_NFS ON TNFS_SAIDA.ID = TITENS_NFS.NFS_ID
JOIN TITENS_COMERCIAL ON TITENS_COMERCIAL.ID = TITENS_NFS.ITCM_ID
JOIN TTIPOS_NF ON TTIPOS_NF.ID = TITENS_NFS.TPNF_ID
JOIN TNATUREZAS_OPERACAO ON TNATUREZAS_OPERACAO.ID = TTIPOS_NF.NAOP_ID
LEFT JOIN THIST_MOV_ITE_PDV ON TITENS_NFS.ID = THIST_MOV_ITE_PDV.ITNFS_ID
LEFT JOIN TITENS_PDV ON TITENS_PDV.ID = THIST_MOV_ITE_PDV.ITPDV_ID
LEFT JOIN TPEDIDOS_VENDA ON TPEDIDOS_VENDA.ID = TITENS_PDV.PDV_ID
LEFT JOIN TDIVISOES_VENDAS ON TDIVISOES_VENDAS.ID = TITENS_NFS.DIVD_ID
WHERE TEMPRESAS.COD_EMP IN (1, 2)
  AND TNFS_SAIDA.DT_EMIS BETWEEN TO_DATE(:DATA_INICIAL,'DD/MM/YYYY')
                             AND TO_DATE(:DATA_FINAL,'DD/MM/YYYY')
ORDER BY TEMPRESAS.COD_EMP, TNFS_SAIDA.DT_EMIS, TNFS_SAIDA.NUM_NF
"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Oracle -> Excel Table (SharePoint)")
    default_date = yesterday()
    parser.add_argument("--data-inicial", default=default_date, help="DD/MM/YYYY; padrao: ontem")
    parser.add_argument("--data-final", default=default_date, help="DD/MM/YYYY; padrao: ontem")
    parser.add_argument("--mode", choices=["append", "replace"], default="append")
    parser.add_argument("--batch-size", type=int, default=500)
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Consulta o Oracle e encerra sem gravar no SharePoint",
    )

    parser.add_argument("--oracle-host", default=os.getenv("ORACLE_HOST", "152.70.223.229"))
    parser.add_argument("--oracle-port", type=int, default=int(os.getenv("ORACLE_PORT", "1521")))
    parser.add_argument("--oracle-sid", default=os.getenv("ORACLE_SID", "f3ipro"))
    parser.add_argument("--oracle-user", default=os.getenv("ORACLE_USER", "FOCCO3I"))
    parser.add_argument("--oracle-password", default=os.getenv("ORACLE_PASSWORD", ""))
    oracle_client_lib_dir = os.getenv("ORACLE_CLIENT_LIB_DIR") or os.getenv("ORACLE_CLIENT_CONFIG_FILE", "")
    parser.add_argument("--oracle-client-lib-dir", default=oracle_client_lib_dir)

    parser.add_argument(
        "--excel-file-path",
        default=os.getenv(
            "EXCEL_FILE_PATH",
            "ASSISTEC/1 - POWERBI/2 - BASES DE DADOS/dev-db_material_rv.xlsx",
        ),
    )
    parser.add_argument(
        "--excel-table-name",
        default=os.getenv("EXCEL_TABLE_NAME", "db_material_rv"),
    )
    args = parser.parse_args()
    try:
        validate_date_range(args.data_inicial, args.data_final)
    except ValueError:
        parser.error("datas devem estar no formato DD/MM/YYYY")
    except argparse.ArgumentTypeError as exc:
        parser.error(str(exc))
    if args.batch_size <= 0:
        parser.error("--batch-size deve ser maior que zero")
    return args


def yesterday() -> str:
    return (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")


def validate_date(value: str) -> str:
    datetime.strptime(value, "%d/%m/%Y")
    return value


def validate_date_range(data_inicial: str, data_final: str) -> None:
    start = datetime.strptime(data_inicial, "%d/%m/%Y")
    end = datetime.strptime(data_final, "%d/%m/%Y")
    if start > end:
        raise argparse.ArgumentTypeError("--data-inicial nao pode ser maior que --data-final")


def graph_headers(token: str, session_id: str | None = None) -> dict[str, str]:
    h = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    if session_id:
        h["workbook-session-id"] = session_id
    return h


def get_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Falha ao obter token Graph: {result}")
    return result["access_token"]


def req(method: str, url: str, headers: dict[str, str], **kwargs) -> dict[str, Any]:
    r = requests.request(method, url, headers=headers, timeout=120, **kwargs)
    if not r.ok:
        raise RuntimeError(f"Erro Graph {r.status_code} em {url}: {r.text}")
    if r.text:
        return r.json()
    return {}


def get_site_id(token: str, hostname: str, site_path: str) -> str:
    data = req("GET", f"{GRAPH_BASE}/sites/{hostname}:{site_path}", graph_headers(token))
    return data["id"]


def get_drive_id(token: str, site_id: str, library_name: str) -> str:
    data = req("GET", f"{GRAPH_BASE}/sites/{site_id}/drives", graph_headers(token))
    requested = library_name.lower()
    shared_documents_aliases = {"documentos compartilhados", "shared documents"}
    documents_aliases = {"documentos", "documents"}
    for drive in data.get("value", []):
        drive_name = drive.get("name", "").lower()
        if drive_name == requested:
            return drive["id"]
        if requested in shared_documents_aliases and drive_name in documents_aliases:
            return drive["id"]
    names = [d.get("name") for d in data.get("value", [])]
    raise RuntimeError(f"Biblioteca '{library_name}' nao encontrada. Disponiveis: {names}")


def get_item_id_by_path(token: str, drive_id: str, item_path: str) -> str:
    normalized = item_path.lstrip("/")
    encoded_path = quote(normalized, safe="/")
    data = req("GET", f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded_path}:", graph_headers(token))
    return data["id"]


def create_workbook_session(token: str, drive_id: str, item_id: str) -> str:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/createSession"
    data = req("POST", url, graph_headers(token), json={"persistChanges": True})
    return data["id"]


def close_workbook_session(token: str, drive_id: str, item_id: str, session_id: str) -> None:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/closeSession"
    req("POST", url, graph_headers(token, session_id), json={})


def get_table_info(token: str, drive_id: str, item_id: str, session_id: str, table_name: str) -> tuple[str, str]:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/tables"
    data = req("GET", url, graph_headers(token, session_id))
    for table in data.get("value", []):
        if table.get("name", "").lower() == table_name.lower():
            return table["id"], table.get("name", table_name)
    names = [t.get("name") for t in data.get("value", [])]
    if len(data.get("value", [])) == 1:
        only_table = data["value"][0]
        print(
            f"Aviso: tabela '{table_name}' nao encontrada; usando a unica tabela disponivel "
            f"'{only_table.get('name')}'.",
            file=sys.stderr,
        )
        return only_table["id"], only_table.get("name", table_name)
    raise RuntimeError(f"Tabela '{table_name}' nao encontrada. Disponiveis: {names}")


def get_table_row_count(token: str, drive_id: str, item_id: str, session_id: str, table_id: str) -> int:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/tables/{table_id}/rows"
    total = 0
    while url:
        data = req("GET", url, graph_headers(token, session_id))
        total += len(data.get("value", []))
        url = data.get("@odata.nextLink", "")
    return total


def get_table_rows(token: str, drive_id: str, item_id: str, session_id: str, table_id: str) -> list[list[Any]]:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/tables/{table_id}/rows"
    rows: list[list[Any]] = []
    while url:
        data = req("GET", url, graph_headers(token, session_id))
        for row in data.get("value", []):
            values = row.get("values", [])
            if values:
                rows.append(values[0])
        url = data.get("@odata.nextLink", "")
    return rows


def get_table_range_address(token: str, drive_id: str, item_id: str, session_id: str, table_id: str) -> str:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/tables/{table_id}/range"
    data = req("GET", url, graph_headers(token, session_id))
    return data["address"]


def parse_range_start_row(address: str) -> int:
    _, _, cells = address.rpartition("!")
    first_cell = cells.split(":", 1)[0]
    match = re.search(r"(\d+)$", first_cell.replace("$", ""))
    if not match:
        raise RuntimeError(f"Nao foi possivel identificar a linha inicial do intervalo: {address}")
    return int(match.group(1))


def parse_range_worksheet_name(address: str) -> str:
    worksheet, sep, _ = address.rpartition("!")
    if not sep:
        raise RuntimeError(f"Nao foi possivel identificar a planilha do intervalo: {address}")
    if worksheet.startswith("'") and worksheet.endswith("'"):
        return worksheet[1:-1].replace("''", "'")
    return worksheet


def normalize_key_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        normalized = f"{value:.8f}".rstrip("0").rstrip(".")
        return normalized or "0"
    if isinstance(value, Decimal):
        if value == value.to_integral_value():
            return str(int(value))
        normalized = format(value.normalize(), "f").rstrip("0").rstrip(".")
        return normalized or "0"
    return str(value).strip()


def row_key(row: list[Any]) -> tuple[str, ...]:
    return tuple(normalize_key_value(value) for value in row[:SOURCE_COLUMN_COUNT])


def filter_duplicate_rows(rows: list[list[Any]], existing_rows: list[list[Any]]) -> list[list[Any]]:
    existing_keys = {row_key(row) for row in existing_rows}
    unique_rows: list[list[Any]] = []
    for row in rows:
        key = row_key(row)
        if key in existing_keys:
            continue
        existing_keys.add(key)
        unique_rows.append(row)
    return unique_rows


def clear_table_rows(token: str, drive_id: str, item_id: str, session_id: str, table_id: str, row_count: int) -> None:
    if row_count <= 0:
        return
    for idx in range(row_count - 1, -1, -1):
        url = (
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/tables/"
            f"{table_id}/rows/itemAt(index={idx})"
        )
        req("DELETE", url, graph_headers(token, session_id))


def add_rows(
    token: str,
    drive_id: str,
    item_id: str,
    session_id: str,
    table_id: str,
    rows: list[list[Any]],
    batch_size: int,
) -> None:
    if not rows:
        return
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/tables/{table_id}/rows/add"
    for i in range(0, len(rows), batch_size):
        chunk = rows[i : i + batch_size]
        req("POST", url, graph_headers(token, session_id), json={"values": chunk})


def user_formula_for_row(worksheet_row: int) -> str:
    return f'=TEXTOANTES(TEXTODEPOIS({USER_FORMULA_SOURCE_COLUMN}{worksheet_row};"/");"/")'


def update_user_formulas(
    token: str,
    drive_id: str,
    item_id: str,
    session_id: str,
    table_range_address: str,
    first_worksheet_row: int,
    row_count: int,
) -> None:
    if row_count <= 0:
        return
    worksheet_name = parse_range_worksheet_name(table_range_address)
    last_data_row = first_worksheet_row + row_count - 1
    worksheet_ref = quote(worksheet_name, safe="")
    for start_row in range(first_worksheet_row, last_data_row + 1, FORMULA_UPDATE_BATCH_SIZE):
        end_row = min(start_row + FORMULA_UPDATE_BATCH_SIZE - 1, last_data_row)
        user_range = f"N{start_row}:N{end_row}"
        formulas = [[user_formula_for_row(row_num)] for row_num in range(start_row, end_row + 1)]
        url = (
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/worksheets/"
            f"{worksheet_ref}/range(address='{user_range}')"
        )
        req("PATCH", url, graph_headers(token, session_id), json={"formulasLocal": formulas})


def excel_serial_date(value: datetime) -> float | int:
    delta = value - EXCEL_EPOCH
    serial = delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86400
    if float(serial).is_integer():
        return int(serial)
    return serial


def normalize_cell(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return excel_serial_date(value)
    if isinstance(value, Decimal):
        return float(value)
    return value


def init_oracle_client(lib_dir: str) -> None:
    if not lib_dir:
        return
    try:
        oracledb.init_oracle_client(lib_dir=lib_dir)
    except Exception as exc:
        details = str(exc)
        if "libaio.so.1" in details:
            raise RuntimeError(
                "Falha ao inicializar Oracle Client: dependencia Linux 'libaio.so.1' ausente. "
                "Instale 'libaio1' no host ou execute via Docker."
            ) from exc
        if "libnnz.so" in details:
            raise RuntimeError(
                "Falha ao inicializar Oracle Client: bibliotecas do Instant Client nao foram "
                "encontradas pelo loader. No Linux, exporte LD_LIBRARY_PATH para o diretorio "
                "do Instant Client antes de iniciar o Python, ou use run_with_instantclient.sh."
            ) from exc
        raise RuntimeError(
            f"Falha ao inicializar Oracle Client em '{lib_dir}'. "
            "Verifique ORACLE_CLIENT_LIB_DIR."
        ) from exc


def fetch_oracle_rows(args: argparse.Namespace) -> list[list[Any]]:
    if not args.oracle_password:
        raise RuntimeError("Informe ORACLE_PASSWORD no .env ou via --oracle-password")

    init_oracle_client(args.oracle_client_lib_dir)
    dsn = oracledb.makedsn(args.oracle_host, args.oracle_port, sid=args.oracle_sid)
    try:
        with oracledb.connect(user=args.oracle_user, password=args.oracle_password, dsn=dsn) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    SQL,
                    DATA_INICIAL=validate_date(args.data_inicial),
                    DATA_FINAL=validate_date(args.data_final),
                )
                rows = cur.fetchall()
    except oracledb.NotSupportedError as exc:
        if "DPY-3015" in str(exc):
            raise RuntimeError(
                "O banco Oracle usa um verificador de senha nao suportado pelo modo thin. "
                "Instale o Oracle Instant Client e informe ORACLE_CLIENT_LIB_DIR no .env "
                "ou via --oracle-client-lib-dir."
            ) from exc
        raise

    return [[normalize_cell(col) for col in row] for row in rows]


def main() -> None:
    base_dir = os.path.dirname(__file__)
    load_dotenv(os.path.join(base_dir, ".env"))
    args = parse_args()

    tenant_id = os.getenv("AZURE_TENANT_ID", "")
    client_id = os.getenv("AZURE_CLIENT_ID", "")
    client_secret = os.getenv("AZURE_CLIENT_SECRET", "")
    hostname = os.getenv("SHAREPOINT_HOSTNAME", "neomotelevadores.sharepoint.com")
    site_path = os.getenv("SHAREPOINT_SITE_PATH", "/sites/ArquivosNeomot")
    library_name = os.getenv("SHAREPOINT_LIBRARY_NAME", "Documentos")

    print(f"Periodo: {args.data_inicial} a {args.data_final}.")
    rows = fetch_oracle_rows(args)
    print(f"Oracle: {len(rows)} linhas retornadas.")
    if args.dry_run:
        print("Dry-run ativo: nenhuma alteracao foi enviada ao SharePoint.")
        return

    if not tenant_id or not client_id or not client_secret:
        raise RuntimeError("Credenciais Azure ausentes no .env")

    token = get_token(tenant_id, client_id, client_secret)
    site_id = get_site_id(token, hostname, site_path)
    drive_id = get_drive_id(token, site_id, library_name)
    item_id = get_item_id_by_path(token, drive_id, args.excel_file_path)

    session_id = create_workbook_session(token, drive_id, item_id)
    try:
        table_id, actual_table_name = get_table_info(
            token,
            drive_id,
            item_id,
            session_id,
            args.excel_table_name,
        )

        if args.mode == "replace":
            row_count = get_table_row_count(token, drive_id, item_id, session_id, table_id)
            print(f"Removendo {row_count} linhas atuais da tabela {actual_table_name}...")
            clear_table_rows(token, drive_id, item_id, session_id, table_id, row_count)
            current_row_count = 0
        else:
            existing_rows = get_table_rows(token, drive_id, item_id, session_id, table_id)
            original_count = len(rows)
            rows = filter_duplicate_rows(rows, existing_rows)
            skipped_count = original_count - len(rows)
            if skipped_count:
                print(f"Linhas ignoradas por ja existirem na tabela {actual_table_name}: {skipped_count}.")
            current_row_count = len(existing_rows)

        range_address = get_table_range_address(token, drive_id, item_id, session_id, table_id)
        header_row = parse_range_start_row(range_address)
        first_new_worksheet_row = header_row + 1 + current_row_count

        add_rows(
            token,
            drive_id,
            item_id,
            session_id,
            table_id,
            rows,
            args.batch_size,
        )
        if rows:
            update_user_formulas(
                token,
                drive_id,
                item_id,
                session_id,
                range_address,
                first_new_worksheet_row,
                len(rows),
            )
        print(f"Carga concluida na tabela {actual_table_name}: {len(rows)} linhas gravadas.")
    finally:
        try:
            close_workbook_session(token, drive_id, item_id, session_id)
        except Exception as exc:
            print(f"Aviso: falha ao fechar sessao do Excel: {exc}", file=sys.stderr)


if __name__ == "__main__":
    try:
        main()
    except RuntimeError as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        raise SystemExit(1) from None
