# Arquitetura do erwin Mart Server

> Documento de referência para compreensão da plataforma e integração via API.

---

## O que é o erwin Mart Server

O **erwin Mart Server** é o componente de repositório centralizado do **erwin Data Modeler (erwin DM)**. Ele armazena modelos de dados em um banco relacional (SQL Server, Oracle ou PostgreSQL) e os disponibiliza para múltiplos usuários de forma colaborativa, com controle de versão, permissões e governança.

A solução é composta por três camadas principais:

```
┌──────────────────────────────────────────────────────────────┐
│                       CLIENTES                               │
│                                                              │
│   erwin DM (desktop)    MartPortal (browser)    Scripts      │
│   (modelagem visual)    (consulta / catalogação) (automação) │
└──────────┬──────────────────────┬───────────────────┬────────┘
           │  SCAPI / COM         │  REST / HTTPS      │ REST
           ▼                      ▼                    ▼
┌──────────────────────────────────────────────────────────────┐
│                    erwin Mart Server                         │
│                                                              │
│  ┌─────────────────┐    ┌──────────────────────────────────┐ │
│  │  MartServer     │    │  MartServerCloud                 │ │
│  │  (API REST)     │    │  (Autenticação JWT)              │ │
│  │                 │    │                                  │ │
│  │  /api/report/   │    │  /csrf           (XSRF token)   │ │
│  │  /service/...   │    │  /jwt/authenticate/login         │ │
│  └────────┬────────┘    └──────────────────────────────────┘ │
└───────────┼──────────────────────────────────────────────────┘
            │
            ▼
┌───────────────────────┐
│   Banco de Dados      │
│   (SQL Server /       │
│    Oracle / PostgreSQL)│
└───────────────────────┘
```

---

## Componentes

### MartServer — API REST

Contexto: `/MartServer`

É a interface principal para consulta e operação sobre os modelos e catálogos. Toda comunicação é autenticada via **Bearer JWT** e protegida por **XSRF token**.

| Grupo | Prefixo | Propósito |
|-------|---------|-----------|
| Relatórios | `/api/report/` | Listar e gerar relatórios sobre modelos |
| Catálogo | `/service/catalog/` | Navegação na árvore de catálogos |
| CGM | `/service/cgm/` | Gerenciamento de glossário de dados corporativo |
| Glossário | `/service/glossary/` | Termos de negócio e mapeamentos |
| Portal | `/service/portal/` | Configurações e metadata browser |
| Health | `/health/readiness` | Verificação de disponibilidade do serviço |

### MartServerCloud — Autenticação

Contexto: `/MartServerCloud`

Responsável exclusivamente pelo fluxo de autenticação. Emite tokens JWT utilizados em todas as chamadas subsequentes à API REST.

### MartPortal — Interface Web

Interface visual para usuários não-modeladores. Permite navegar no catálogo de modelos, pesquisar metadados, visualizar diagramas e gerenciar termos de negócio sem instalar o erwin DM.

---

## Autenticação

O Mart Server usa **JWT (JSON Web Token)** combinado com proteção **CSRF**. O fluxo é obrigatório para qualquer chamada autenticada à API.

### Fluxo em 2 etapas

```
Cliente                              Mart Server
  │                                       │
  │── GET /MartServer/csrf ──────────────>│
  │<── 403 + header XSRF-TOKEN ──────────│  ← token anti-CSRF
  │                                       │
  │── POST /MartServerCloud/jwt/authenticate/login
  │   Header: X-XSRF-TOKEN: <xsrf>       │
  │   Body:   { username, password }      │
  │<── 200 + { "id_token": "<jwt>" } ────│  ← Bearer token
  │                                       │
  │── GET /MartServer/api/report/...      │
  │   Header: Authorization: Bearer <jwt>│
  │   Header: X-XSRF-TOKEN: <xsrf>       │
  │<── 200 + dados ─────────────────────│
```

### Detalhes dos endpoints de autenticação

**Etapa 1 — CSRF token**

```
GET https://<host>/MartServer/csrf

Resposta:
  HTTP 403
  Header: XSRF-TOKEN: <uuid>
```

> O status 403 é esperado. O token está no **header** da resposta, mesmo em erro.

**Etapa 2 — Login JWT**

```
POST https://<host>/MartServerCloud/jwt/authenticate/login

Headers:
  Content-Type: application/json
  X-XSRF-TOKEN: <xsrf_obtido_na_etapa_1>

Body:
  {
    "username": "Administrator",
    "password": "senha"
  }

Resposta 200:
  {
    "id_token": "<jwt_bearer_token>"
  }
```

### Uso do Bearer token nas chamadas subsequentes

```
GET https://<host>/MartServer/api/report/generateReport/Models

Headers:
  Authorization: Bearer <id_token>
  X-XSRF-TOKEN: <xsrf>
```

---

## Endpoints Principais

### Listar modelos do repositório

```
GET /MartServer/api/report/generateReport/Models
```

Retorna XML com todos os modelos cadastrados no Mart, incluindo:
- `Catalog_Name` — nome do modelo
- `Catalog_Path` — caminho completo no repositório (ex: `Mart/Modelos`)
- `UpdatedOn` — data/hora da última modificação

### Árvore de catálogos

```
GET /MartServer/service/catalog/tree
```

Retorna a hierarquia de pastas e modelos do repositório.

### Verificação de saúde

```
GET /MartServer/health/readiness
```

Verifica se o serviço está disponível. Não requer autenticação.

---

## Integração via SCAPI e ERXML (Windows COM)

Para exportar o conteúdo de um modelo para XML legível, o processo usa dois objetos COM registrados pela instalação do erwin DM:

```
Script Python (Windows)
        │
        │  win32com.client.Dispatch(...)
        ▼
┌──────────────────────────────────────────────────┐
│  erwin DM (instalado localmente)                 │
│                                                  │
│  ┌─────────────────────────────────┐             │
│  │  erwin9.SCAPI.9.0               │             │
│  │  Abre modelo do Mart via locator│             │
│  │  Salva .erwin temporário local  │             │
│  └────────────────┬────────────────┘             │
│                   │ .erwin temporário            │
│  ┌────────────────▼────────────────┐             │
│  │  ERXML.XMLERwinLink             │             │
│  │  Converte .erwin → XML UTF-8    │             │
│  │  StandAloneExport(...)          │             │
│  └─────────────────────────────────┘             │
└──────────────────────────────────────────────────┘
        │
        ▼
   output/xml/<Catalog_Path>/<Catalog_Name>.xml
```

**Locator SCAPI** — formato usado para referenciar um modelo no Mart:

```
mart://Mart/<Catalog_Path_sem_prefixo>/<Catalog_Name>?TRC=NO;SRV=<host>;PRT=443;ASR=MartServer;SSL=YES;UID=<user>;PSW=<pass>
```

Exemplo:
```
mart://Mart/Modelos/eMovies?TRC=NO;SRV=mart.empresa.com.br;PRT=443;ASR=MartServer;SSL=YES;UID=Administrator;PSW=senha
```

---

## Documentação OpenAPI

A documentação completa da API REST está disponível em:

| Formato | URL |
|---------|-----|
| JSON (OpenAPI 3) | `https://<host>/MartServer/v3/api-docs` |
| Swagger UI | `https://<host>/MartServer/swagger-ui/index.html` |

A documentação lista todos os endpoints disponíveis com schemas de request/response, incluindo os grupos: `report-controller`, `mart-catalog-controller`, `cgm-controller`, `glossary-controller`, `portal`, `home` e `license`.

---

## Estrutura de Pastas do Repositório

Os modelos são organizados em uma hierarquia de pastas dentro do Mart. O campo `Catalog_Path` retornado pela API reflete essa estrutura:

```
Mart/
├── Modelos/
│   ├── eMovies
│   └── Contrato_Pessoa_Produto
├── Ambiente/
│   └── Homologacao/
│       └── exemploMongo
└── API2/
    └── ...
```

Ao exportar, a ferramenta replica essa estrutura localmente em `output/xml/`.

---

## Fluxo Completo de Exportação

```
┌─────────────────────────────────────────────────────────────┐
│  erwin_save_xml.py                                          │
│                                                             │
│  1. Carrega .env                                            │
│  2. Gera tokens (CSRF + JWT) se não informados              │
│  3. GET /MartServer/api/report/generateReport/Models        │
│     → Lista todos os modelos com Catalog_Path e UpdatedOn   │
│  4. Filtra por data (opcional)                              │
│  5. Para cada modelo:                                       │
│     a. Monta locator SCAPI                                  │
│     b. SCAPI.PersistenceUnits.Add(locator) → abre modelo    │
│     c. PersistenceUnit.Save(temp.erwin)                     │
│     d. ERXML.StandAloneExport(temp.erwin, saida.xml)        │
│     e. Remove temp.erwin                                    │
│  6. Salva resumo em output/mart_report/mart_models.xml      │
└─────────────────────────────────────────────────────────────┘
```

---

## Requisitos para Integração

| Requisito | Detalhes |
|-----------|----------|
| Sistema Operacional | Windows (obrigatório para COM/SCAPI/ERXML) |
| erwin DM | Versão 15.2 instalada (registra DLLs COM) |
| Python | >= 3.14 |
| pywin32 | >= 311 |
| Rede | HTTPS na porta 443 para o host do Mart Server |
| Credenciais | Usuário com permissão de leitura no Mart |
