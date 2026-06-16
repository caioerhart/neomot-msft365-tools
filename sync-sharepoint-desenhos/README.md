# Sync SharePoint Desenhos ENG

Script para copiar arquivos `.pdf` e `.dxf` modificados nos últimos dias da pasta de origem para o destino no SharePoint.

## O que ele faz

1. Lê origem e destino via `.env`.
2. Procura arquivos modificados nos últimos `DAYS_BACK` dias.
3. Filtra por extensões definidas em `FILE_EXTENSIONS`.
4. Se `DEST_SYNC_FOLDER_NAME` estiver preenchido, cria (se não existir) essa subpasta dentro de `DEST_PARENT_FOLDER_PATH`.
5. Se `DEST_SYNC_FOLDER_NAME` estiver vazio, copia direto na raiz de `DEST_PARENT_FOLDER_PATH`.

## Exemplo de cenário (definitivo)

- Origem: `INDUSTRIAL/ENGENHARIA MECANICA/MODELOS - PEÇAS`
- Destino: `INDUSTRIAL/PRODUCAO/BIBLIOTECA DESENHOS`
- Subpasta: vazia (cópia direto na raiz)

## Configuração

1. Copie `example.env` para `.env`.
2. Preencha as credenciais do App Registration e ajuste os caminhos.
3. Garanta a permissão Graph Application `Sites.ReadWrite.All`.

## Execução

```bash
cd sync-sharepoint-desenhos
python sync_recent_drawings.py --dry-run
python sync_recent_drawings.py
```

## Opções

- `--dry-run`: simula sem copiar.

## Dependências

Usa pacotes já existentes no repositório:

- `msal`
- `requests`
- `python-dotenv`
