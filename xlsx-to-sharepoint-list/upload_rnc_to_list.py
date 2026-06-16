"""
upload_rnc_to_list.py
---------------------
Lê a planilha 'Dados Consolidados' do arquivo RNC📄(1).xlsx e cria (ou
atualiza) os itens na lista SharePoint  sites/ArquivosNeomot/Lists/RNC,
usando a Microsoft Graph API com autenticação de aplicativo (app-only).

Comportamento:
  - Por padrão faz UPSERT usando o campo "Identificação da tarefa" como chave:
      • se o item já existir → atualiza (PATCH)
      • se não existir → cria (POST)
  - Use --dry-run para simular sem gravar nada.
  - Use --force-create para ignorar a checagem e sempre criar.

Uso:
  python upload_rnc_to_list.py
  python upload_rnc_to_list.py --xlsx caminho/para/RNC.xlsx
  python upload_rnc_to_list.py --dry-run
  python upload_rnc_to_list.py --force-create

Variáveis de ambiente (.env no mesmo diretório):
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
  SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH

Permissões necessárias no App Registration:
  Microsoft Graph → Application → Sites.ReadWrite.All
"""

import argparse
import os
import re
import sys
from datetime import datetime

import msal
import pandas as pd
import requests
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

TENANT_ID           = os.getenv("AZURE_TENANT_ID", "")
CLIENT_ID           = os.getenv("AZURE_CLIENT_ID", "")
CLIENT_SECRET       = os.getenv("AZURE_CLIENT_SECRET", "")
SHAREPOINT_HOSTNAME = os.getenv("SHAREPOINT_HOSTNAME", "neomotelevadores.sharepoint.com")
SITE_PATH           = os.getenv("SHAREPOINT_SITE_PATH", "/sites/ArquivosNeomot")
LIST_NAME           = "RNC"

GRAPH_BASE  = "https://graph.microsoft.com/v1.0"
SHEET_NAME  = "Dados Consolidados"

DEFAULT_XLSX = os.path.join(os.path.dirname(__file__), "RNC📄(1).xlsx")

# Mapeamento:  coluna Excel  →  nome interno do campo na lista SharePoint
# Ajuste os valores à direita para corresponder ao InternalName de cada campo
# da lista RNC no SharePoint (consulte /_api/web/lists/getbytitle('RNC')/fields).
# Mapeamento: coluna Excel → InternalName do campo na lista SharePoint RNC
COLUMN_MAP: dict[str, str] = {
    "Identificação da tarefa":  "_IdentificacaoDaTarefa",  # chave de upsert (não enviado à API)
    "Nome da tarefa":           "Title",                  # Texto → campo Código
    "Status":                   "Status",                 # Opção
    "Prioridade":               "Priority",               # Opção
    "Criado em":                "_CriadoEm",              # readOnly, apenas para referência
    "Data de conclusão":        "_DataDeConclusao",
    "Data de início":           "_DataDeInicio",
    "Concluído em":             "_ConcluidoEm",
}

# Campos extras extraídos do texto livre da coluna "Notas"
# Chave  = label no texto das Notas
# Valor  = InternalName do campo na lista SharePoint
# Prefixo "__" indica campo intermediário que será fundido em outro campo
NOTES_FIELD_MAP: dict[str, str] = {
    "Data":                      "Data",                              # Data e Hora
    "Tipo de não conformidade":   "Tipoden_x00e3_oconformidade",       # Opção
    "Cliente / Empreendimento":   "Cliente_x002f_Empreendimento",      # Texto linha única
    "Código da peça":             "C_x00f3_digodape_x00e7_a",           # Texto linha única
    "Equipamento":                "Equipamento",                       # Opção
    "Descrição do problema":      "Description",                       # Texto multilinha
    "Motivo":                     "__motivo__",                        # → fundido em Description
    "Tempo de resposta":          "_TempoResposta",                    # sem campo na lista; ignorado
    "Resolução":                  "__resolucao__",                     # → fundido em Description
    "Impacto":                    "_Impacto",                          # sem campo na lista; ignorado
    "ID do Equipamento":          "IDdoequipamento",                   # Texto linha única
}

# URL base para links de pastas RNC no SharePoint
RNC_FOLDER_BASE = (
    "https://neomotelevadores.sharepoint.com/:f:/r/sites/ArquivosNeomot"
    "/Documentos%20Compartilhados/INDUSTRIAL/QUALIDADE"
    "/FORMUL%C3%81RIO%20RNC"
)
# Sufixo fixo do link compartilhável
RNC_LINK_SUFFIX = "?csf=1&web=1&e=jBJnY8"

# Campo usado como chave de upsert
KEY_EXCEL_COL   = "Identificação da tarefa"
KEY_SP_FIELD    = "Title"   # campo Title = Código na lista; usado para upsert

# Normalização dos valores de Status: Excel → choice exata da lista SharePoint
STATUS_MAP: dict[str, str] = {
    "em andamento":   "Em andamento",
    "não iniciado":   "Não iniciado",
    "nao iniciado":   "Não iniciado",
    "concluída":      "Concluído",
    "concluido":      "Concluído",
    "concluído":      "Concluído",
    "em análise":     "Em análise",
    "em analise":     "Em análise",
}


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


def headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


# ---------------------------------------------------------------------------
# SharePoint helpers
# ---------------------------------------------------------------------------

def get_site_id(token: str) -> str:
    url = f"{GRAPH_BASE}/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}"
    r = requests.get(url, headers=headers(token))
    r.raise_for_status()
    return r.json()["id"]


def get_list_id(token: str, site_id: str) -> str:
    url = f"{GRAPH_BASE}/sites/{site_id}/lists"
    r = requests.get(url, headers=headers(token))
    r.raise_for_status()
    for lst in r.json().get("value", []):
        if lst["name"].lower() == LIST_NAME.lower() or lst["displayName"].lower() == LIST_NAME.lower():
            return lst["id"]
    raise RuntimeError(
        f"Lista '{LIST_NAME}' não encontrada. Listas disponíveis: "
        + str([l["displayName"] for l in r.json().get("value", [])])
    )


def list_existing_items(token: str, site_id: str, list_id: str) -> dict[str, int]:
    """Retorna {chave: item_id} para todos os itens existentes na lista."""
    existing = {}
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"
        f"?$expand=fields($select={KEY_SP_FIELD},id)&$top=999"
    )
    while url:
        r = requests.get(url, headers=headers(token))
        r.raise_for_status()
        data = r.json()
        for item in data.get("value", []):
            key = item.get("fields", {}).get(KEY_SP_FIELD, "")
            if key:
                existing[key] = item["id"]
        url = data.get("@odata.nextLink")
    return existing


def create_item(token: str, site_id: str, list_id: str, fields: dict) -> str:
    url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"
    payload = {"fields": fields}
    r = requests.post(url, headers=headers(token), json=payload)
    r.raise_for_status()
    return r.json()["id"]


def update_item(token: str, site_id: str, list_id: str, item_id: str, fields: dict):
    url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
    r = requests.patch(url, headers=headers(token), json=fields)
    r.raise_for_status()


# ---------------------------------------------------------------------------
# Transformação de dados
# ---------------------------------------------------------------------------

def _safe_str(val) -> str | None:
    """Converte valor para string, retorna None se vazio/NaN."""
    if val is None:
        return None
    if isinstance(val, float) and (val != val):  # NaN
        return None
    s = str(val).strip()
    return s if s else None


def _safe_bool(val) -> bool | None:
    if val is None:
        return None
    if isinstance(val, float) and (val != val):
        return None
    if isinstance(val, bool):
        return val
    s = str(val).strip().lower()
    if s in ("true", "sim", "yes", "1"):
        return True
    if s in ("false", "não", "nao", "no", "0"):
        return False
    return None


# Padrão para código RNC  ex: RNC-20260424-4
_RNC_CODE_RE = re.compile(r"RNC-\d{8}-\d+", re.IGNORECASE)


def _rnc_link(code: str) -> str:
    """Monta URL para a pasta do código RNC no SharePoint."""
    return f"{RNC_FOLDER_BASE}/{code}{RNC_LINK_SUFFIX}"


def parse_notas(text: str) -> dict:
    """
    Fragmenta o texto livre da coluna "Notas" nos campos estruturados.
    Formato esperado (gerado pelo Planner):
        ChaveA:\nValorA;\nChaveB:\nValorB;\n...

    Retorna um dict com:
      - campos de NOTES_FIELD_MAP quando encontrados
      - "CodigosRNC"  → lista de códigos  ex: ["RNC-20260424-4"]
      - "LinksRNC"    → HTML multi-link ou string vazia
    """
    result: dict = {}
    if not text:
        return result

    # ── Detectar códigos RNC e gerar links ──────────────────────────────────
    rnc_codes = _RNC_CODE_RE.findall(text)
    if rnc_codes:
        # Title = "Código" na lista SharePoint — chave de upsert
        result["Title"] = rnc_codes[0]
        # Linkdosanexos = "Link dos anexos" — texto com URL para a pasta no SharePoint
        result["Linkdosanexos"] = _rnc_link(rnc_codes[0])

    # ── Parsear pares chave: valor ───────────────────────────────────────────
    # Normaliza separadores de linha e ponto-e-vírgula final
    clean = text.replace("\r\n", "\n").replace("\r", "\n")

    # Divide em blocos "Chave:\nValor;" usando lookahead para próxima chave
    known_keys = list(NOTES_FIELD_MAP.keys())
    # Monta regex que captura cada par até a próxima chave conhecida ou fim
    escaped = "|".join(re.escape(k) for k in known_keys)
    pattern = re.compile(
        rf"({escaped}):\s*\n?(.*?)(?=(?:{escaped}):|\Z)",
        re.DOTALL | re.IGNORECASE,
    )

    for m in pattern.finditer(clean):
        key_raw = m.group(1).strip()
        val_raw = m.group(2).strip().rstrip(";,").strip()
        if not val_raw:
            continue
        # Remove lixo de linha separadora (---, ===, etc.)
        val_raw = re.sub(r"[-=]{3,}.*", "", val_raw).strip().rstrip(";,").strip()
        # Pega somente até o primeiro ';' ou fim da primeira linha não-vazia
        # para evitar capturar sub-chaves que ainda não foram parseadas
        first_line = val_raw.split(";")[0].split("\n")[0].strip()
        val_final = first_line if first_line else val_raw
        if not val_final:
            continue
        # Localizar chave canônica (case-insensitive)
        for canonical in known_keys:
            if canonical.lower() == key_raw.lower():
                sp_field = NOTES_FIELD_MAP[canonical]
                result[sp_field] = val_final
                break

    # ── Fundir Motivo e Resolução em Descrição do problema ──────────────────
    desc_parts = []
    if result.get("Description"):
        desc_parts.append(result["Description"])
    motivo = result.pop("__motivo__", None)
    resolucao = result.pop("__resolucao__", None)
    if motivo:
        desc_parts.append(f"Motivo: {motivo}")
    if resolucao:
        desc_parts.append(f"Resolução: {resolucao}")
    if desc_parts:
        result["Description"] = "\n\n".join(desc_parts)

    return result


def row_to_fields(row: pd.Series) -> dict:
    """Converte uma linha do DataFrame nos campos do item SharePoint."""
    fields: dict = {}
    for excel_col, sp_field in COLUMN_MAP.items():
        if excel_col not in row.index:
            continue
        val = row[excel_col]

        # Booleanos (campos removidos do COLUMN_MAP; mantido por precaução)
        if excel_col in ("É Recorrente", "Atrasados"):
            b = _safe_bool(val)
            if b is not None:
                fields[sp_field] = b
            continue

        # Status → normaliza para choice exata da lista
        if excel_col == "Status":
            s = _safe_str(val)
            if s:
                fields[sp_field] = STATUS_MAP.get(s.lower(), s)
            continue

        # Datas
        if excel_col in ("Criado em", "Data de conclusão", "Data de início",
                         "Concluído em"):
            s = _safe_str(val)
            if s:
                # Normaliza para ISO 8601 (aceita DD/MM/YYYY ou YYYY-MM-DD)
                for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                    try:
                        dt = datetime.strptime(s[:10], fmt)
                        fields[sp_field] = dt.strftime("%Y-%m-%dT00:00:00Z")
                        break
                    except ValueError:
                        continue
                else:
                    fields[sp_field] = s  # mantém como string se não parsear
            continue

        # Numéricos (itens de lista de verificação)
        if excel_col in ("Itens concluídos da lista de verificação",
                         "Itens da lista de verificação"):
            s = _safe_str(val)
            if s:
                try:
                    fields[sp_field] = int(float(s))
                except ValueError:
                    fields[sp_field] = s
            continue

        # Demais → string
        s = _safe_str(val)
        if s is not None:
            fields[sp_field] = s

    # ── Fragmentar coluna Notas ──────────────────────────────────────────────
    notas_raw = _safe_str(row.get("Notas"))
    if notas_raw:
        parsed = parse_notas(notas_raw)
        # LinkRNC é um hiperlink → formato especial para Graph API
        # Campos de Url no SharePoint precisam ser enviados como string URL simples
        # quando o tipo do campo for URL; ajuste abaixo se for campo de texto simples.
        for k, v in parsed.items():
            fields[k] = v

    # Aprovação da RNC → sempre "Aprovado" (itens migrados do Planner já foram aprovados)
    fields["Aprova_x00e7__x00e3_odaRNC"] = "Aprovado"

    # Remover campos somente-leitura/sem mapeamento (prefixo _) e campos intermediários (__)
    return {k: v for k, v in fields.items() if not k.startswith("_")}


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args():
    p = argparse.ArgumentParser(description="Upload RNC Excel → SharePoint List")
    p.add_argument("--xlsx",         default=DEFAULT_XLSX,
                   help="Caminho para o arquivo Excel RNC (padrão: RNC📄(1).xlsx no mesmo diretório)")
    p.add_argument("--sheet",        default=SHEET_NAME,
                   help=f"Nome da aba (padrão: '{SHEET_NAME}')")
    p.add_argument("--dry-run",      action="store_true",
                   help="Simula sem gravar nada no SharePoint")
    p.add_argument("--force-create", action="store_true",
                   help="Sempre cria novos itens (ignora upsert)")
    p.add_argument("--debug",        action="store_true",
                   help="Imprime os fields de cada linha antes de enviar")
    p.add_argument("--row",          type=int, default=None,
                   help="Processar apenas a linha N do Excel (número da linha, começando em 2)")
    return p.parse_args()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    args = parse_args()

    # --- Ler Excel ---
    print(f"Lendo '{args.xlsx}' aba '{args.sheet}'...")
    df = pd.read_excel(args.xlsx, sheet_name=args.sheet, dtype=str)
    # Substituir NaN por None para facilitar conversão
    df = df.where(df.notna(), None)
    print(f"  {len(df)} linha(s) encontrada(s).")

    if KEY_EXCEL_COL not in df.columns:
        print(f"[AVISO] Coluna '{KEY_EXCEL_COL}' não encontrada — chave virá do campo Title (Notas).")

    if args.dry_run:
        print("\n[DRY-RUN] Nenhuma alteração será feita no SharePoint.\n")
    if args.debug:
        print("[DEBUG] Modo debug ativo — fields serão impressos antes do envio.\n")

    # --- Autenticar e resolver IDs ---
    print("Autenticando no Microsoft Graph...")
    token   = get_access_token()
    site_id = get_site_id(token)
    list_id = get_list_id(token, site_id)
    print(f"  Site  : {site_id}")
    print(f"  Lista : {list_id} ('{LIST_NAME}')")

    # --- Carregar itens existentes ---
    existing: dict[str, str] = {}
    if not args.force_create:
        print("Carregando itens existentes para upsert...")
        existing = list_existing_items(token, site_id, list_id)
        print(f"  {len(existing)} item(s) existente(s).")

    # --- Processar linhas ---
    created = updated = skipped = errors = 0
    for idx, row in df.iterrows():
        # --row filtra para uma linha específica (número da linha no Excel, base 2)
        if args.row is not None and (idx + 2) != args.row:
            continue

        fields = row_to_fields(row)
        if not fields:
            print(f"  [SKIP] Linha {idx+2}: nenhum campo mapeado.")
            skipped += 1
            continue

        # Chave de upsert = Title (código RNC extraído das Notas ou "Nome da tarefa")
        key = fields.get(KEY_SP_FIELD) or _safe_str(row.get(KEY_EXCEL_COL))
        if not key:
            print(f"  [SKIP] Linha {idx+2}: campo chave '{KEY_SP_FIELD}' vazio.")
            skipped += 1
            continue

        if args.debug:
            import json
            print(f"  [DEBUG] Linha {idx+2} | chave: {key[:60]}")
            for k, v in sorted(fields.items()):
                print(f"    {k}: {repr(str(v)[:120])}")
            print()

        try:
            if not args.dry_run:
                token = get_access_token()  # renova token a cada iteração (longa execução)

            if not args.force_create and key in existing:
                # UPDATE
                item_id = existing[key]
                if not args.dry_run:
                    update_item(token, site_id, list_id, item_id, fields)
                print(f"  [UPDATE] Linha {idx+2}: '{key[:60]}' → item {item_id}")
                updated += 1
            else:
                # CREATE
                if not args.dry_run:
                    new_id = create_item(token, site_id, list_id, fields)
                    print(f"  [CREATE] Linha {idx+2}: '{key[:60]}' → item {new_id}")
                else:
                    print(f"  [CREATE-DRY] Linha {idx+2}: '{key[:60]}'")
                created += 1

        except requests.HTTPError as e:
            body = e.response.text
            if args.debug:
                print(f"  [ERRO] Linha {idx+2}: {e}\n  Resposta: {body}", file=sys.stderr)
            else:
                print(f"  [ERRO] Linha {idx+2}: {e} – {body[:200]}", file=sys.stderr)
            errors += 1

    # --- Resumo ---
    print(f"\nConcluído: {created} criado(s), {updated} atualizado(s), "
          f"{skipped} ignorado(s), {errors} erro(s).")
    if args.dry_run:
        print("[DRY-RUN] Nenhum dado foi gravado.")


if __name__ == "__main__":
    main()
