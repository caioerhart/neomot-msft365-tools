# Oracle -> db_material_rv (Excel no SharePoint)

Script isolado para:
1. Consultar dados no Oracle.
2. Inserir na tabela db_material_rv do arquivo dev-db_material_rv.xlsx via Microsoft Graph.

## Arquivos

- oracle_to_db_material_rv.py
- requirements.txt
- example.env
- Dockerfile

## Configuracao

1. Copie example.env para .env.
2. Preencha credenciais do Azure App e Oracle.
3. Instale dependencias:

   pip install -r requirements.txt

## Execucao

Teste sem gravar no SharePoint:

python oracle_to_db_material_rv.py --dry-run

Modo append (adiciona linhas):

python oracle_to_db_material_rv.py --mode append

Modo replace (limpa a tabela e recarrega):

python oracle_to_db_material_rv.py --data-inicial 01/06/2026 --data-final 23/06/2026 --mode replace

Por padrao, quando `--data-inicial` e `--data-final` nao sao informadas, o script usa
o dia anterior ao dia da execucao para as duas datas. Informe as datas manualmente
somente quando precisar reprocessar outro periodo.

## Oracle Instant Client

Se a conexao Oracle retornar `DPY-3015`, instale o Oracle Instant Client e informe o
caminho em `ORACLE_CLIENT_CONFIG_FILE` ou `ORACLE_CLIENT_LIB_DIR` no `.env`.

Exemplo Windows:

```env
ORACLE_CLIENT_CONFIG_FILE=C:\Users\modelo.bi\Documents\Modelo\integra_BI\instantclient\instantclient_23_0
```

No Docker/Linux o caminho configurado e:

```env
ORACLE_CLIENT_CONFIG_FILE=/opt/oracle/instantclient
```

Para usar o Instant Client baixado localmente neste repositorio:

```env
ORACLE_CLIENT_CONFIG_FILE=./instantclient/current
```

Em Linux fora do Docker, instale tambem a dependencia do sistema:

```bash
sudo apt-get install libaio1
```

Depois execute com o wrapper, que configura `LD_LIBRARY_PATH` antes de iniciar o Python:

```bash
./run_with_instantclient.sh --dry-run
```

## Docker

Build da imagem usando o Instant Client Linux ja baixado em `instantclient/instantclient_23_26`:

```bash
docker build -t oracle-db-material-rv-sync .
```

Execucao em dry-run usando o `.env` local:

```bash
docker run --rm --env-file .env oracle-db-material-rv-sync --dry-run
```

Execucao gravando no SharePoint:

```bash
docker run --rm --env-file .env oracle-db-material-rv-sync --mode append
```

Para reprocessar um periodo especifico:

```bash
docker run --rm --env-file .env oracle-db-material-rv-sync --data-inicial 01/06/2026 --data-final 23/06/2026 --mode append
```
