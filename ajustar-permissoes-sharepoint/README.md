# Ajustar Permissões de Pastas — SharePoint

Atribui permissões de pasta no SharePoint a grupos/usuários M365 a partir
de um CSV, usando somente o **Microsoft Graph API**.

---

## Como funciona

Para cada linha do CSV o script:

1. Concede a permissão (`write` ou `read`) na pasta alvo
2. Concede `read` nas pastas pai para que o grupo consiga navegar até ela
3. Remove a permissão do grupo das pastas **irmãs** de cada nível pai,
   garantindo que o grupo veja apenas o caminho autorizado

```
Biblioteca/
├── INDUSTRIAL/          ← grupo vê (read automático no pai)
│   └── PRODUÇÃO/        ← grupo vê (read automático no pai)
│       └── CLIENTES/    ← grupo tem write aqui
│       └── OUTRO/       ← grupo NÃO vê (permissão removida)
├── JURÍDICO/            ← grupo NÃO vê
└── MARKETING/           ← grupo NÃO vê
```

---

## Uso

```bash
# CSV padrão (permissions.csv no mesmo diretório)
python set_folder_permission.py

# CSV personalizado
python set_folder_permission.py outro_arquivo.csv
```

---

## Formato do CSV

```csv
email,pasta,permissao
qualidade@neomot.com,INDUSTRIAL,write
pcp@neomot.com,INDUSTRIAL/PRODUÇÃO/CLIENTES,write
financeiro@neomot.com,JURÍDICO,read
```

| Coluna | Descrição |
|---|---|
| `email` | E-mail do grupo ou usuário M365 |
| `pasta` | Caminho relativo dentro da biblioteca (separador `/`) |
| `permissao` | `write` (Editor) ou `read` (Leitor) |

- Linhas vazias e linhas começando com `#` são ignoradas
- Valores inválidos em `permissao` caem para `write` com aviso

---

## Configuração

Copie `example.env` para `.env` na raiz do repositório e preencha:

```dotenv
AZURE_TENANT_ID=...
AZURE_CLIENT_ID=...
AZURE_CLIENT_SECRET=...

SHAREPOINT_HOSTNAME=neomotelevadores.sharepoint.com
SHAREPOINT_SITE_PATH=/sites/ArquivosNeomot
SHAREPOINT_LIBRARY_NAME=Documentos Compartilhados
```

### Permissões necessárias no App Registration

**Azure Portal → Entra ID → App registrations → seu app → API permissions:**

```
Microsoft Graph → Application → Sites.ReadWrite.All  ✓ Admin consent granted
```

---

## Erros conhecidos

| Erro | Causa | Solução |
|---|---|---|
| `404 Not Found` | Nome da pasta no CSV diferente do SharePoint | Verifique acentos, maiúsculas e espaços |
| `429 Too Many Requests` | Rate limit do Graph API | Re-execute; as linhas com erro são exibidas no resumo final |
| `400 invalidRequest` | Grupo sem e-mail (grupo de segurança puro) | Use grupos M365 (com e-mail) |
