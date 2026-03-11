# erwinDM Save to XML

Utilitario Python para:
- listar modelos do Mart Server em XML (`mart_models.xml`)
- filtrar modelos por `UpdatedOn`
- montar locator SCAPI com `Catalog_Path` + `Catalog_Name`
- exportar modelos do Mart para XML legivel
- exportar modelo local `.erwin` para XML legivel

## Requisitos

- erwin DM 15.2 instalado (registra as DLLs COM)
- Python >= 3.14
- [uv](https://docs.astral.sh/uv/)

## Instalacao

```bash
uv sync
```

## Configuracao Segura (.env)

Os dados sensiveis (URL, token, usuario, senha) devem ficar no arquivo `.env`.

1. Copie o template:

```bash
copy .env.example .env
```

2. Preencha com seus valores reais.

Variaveis usadas pelo script:

- `MART_URL`
- `MART_BEARER_TOKEN`
- `MART_XSRF_TOKEN`
- `MART_CONN_STR`
- `MART_UPDATED_ON_EXACT` (opcional, formato `MM/DD/YYYY`)
- `MART_UPDATED_ON_MIN` (opcional, formato `MM/DD/YYYY`)
- `MART_UPDATED_ON_MAX` (opcional, formato `MM/DD/YYYY`)

Regras de filtro de data:
- Exata: use apenas `MART_UPDATED_ON_EXACT`
- Minima: use `MART_UPDATED_ON_MIN`
- Intervalo: use `MART_UPDATED_ON_MIN` + `MART_UPDATED_ON_MAX`
- `EXACT` nao pode ser combinado com `MIN/MAX`

O arquivo `.env` esta no `.gitignore` e nao deve ser versionado.

## Arquitetura Atual

Fluxo principal para Mart:
1. `listar_modelos_mart(...)` chama `GET /MartServer/api/report/generateReport/Models`
2. `filtrar_modelos_mart_por_data(...)` seleciona modelos por `UpdatedOn`:
  data exata, minima ou intervalo
3. `montar_locator_mart(...)` gera:
   `mart://Mart/<Catalog_Path_sem_prefixo>/<Catalog_Name>?<mart_conn_str>`
4. `mart_exportar_todos_xml(...)` itera nos modelos filtrados, abre via SCAPI, salva `.erwin` temporario e converte para XML via `ERXML.XMLERwinLink`

Fluxo para arquivo local:
1. `erwin_to_xml(...)` chama `ERXML.XMLERwinLink.StandAloneExport`
2. XML e indentado e salvo em UTF-8

## Funcoes Disponiveis

### `erwin_to_xml(caminho_modelo_erwin, caminho_saida_xml) -> bool`
Exporta um arquivo `.erwin` local para XML legivel.

### `listar_modelos_mart(mart_url, bearer_token, xsrf_token, caminho_saida) -> bool`
Baixa a lista de modelos do Mart e salva o XML de relatorio.

### `filtrar_modelos_mart_por_data(mart_url, bearer_token, xsrf_token, data_atualizacao_min=None, data_atualizacao_max=None, data_atualizacao_exata=None) -> list[dict[str, str]]`
Retorna lista com `Catalog_Name`, `Catalog_Path` e `UpdatedOn` com filtro por data exata, minima ou intervalo.

### `montar_locator_mart(catalog_path, catalog_name, mart_conn_str) -> str`
Monta o locator Mart usado na conexao SCAPI.

### `mart_exportar_todos_xml(mart_url, bearer_token, xsrf_token, mart_conn_str, caminho_saida_dir, caminho_temp_dir, data_atualizacao_min=None, data_atualizacao_max=None, data_atualizacao_exata=None) -> dict[str, bool]`
Exporta para XML todos os modelos retornados (ou filtrados por data), usando `Catalog_Path` e `Catalog_Name` para montar o locator.

## Organizacao de Pastas de Saida

Quando executado pelo `__main__`:

- relatorio de modelos:
  `output/mart_report/mart_models.xml`
- XMLs exportados dos modelos:
  `output/xml/*.xml`
- temporarios `.erwin`:
  `temp/*.erwin` (removidos apos exportacao)

## Exemplo de Uso

```python
from erwin_save_xml import (
    listar_modelos_mart,
    filtrar_modelos_mart_por_data,
    montar_locator_mart,
    mart_exportar_todos_xml,
)

MART_URL = "https://mart.exemplo.com.br"
BEARER = "<token>"
XSRF = "<xsrf>"
MART_CONN = "TRC=NO;SRV=mart.exemplo.com.br;PRT=443;ASR=MartServer;SSL=YES;UID=...;PSW=..."

# 1) Salva o relatorio completo
listar_modelos_mart(MART_URL, BEARER, XSRF, "output/mart_report/mart_models.xml")

# 2) Filtra por UpdatedOn
filtrados = filtrar_modelos_mart_por_data(MART_URL, BEARER, XSRF, "03/11/2026")
for m in filtrados:
    locator = montar_locator_mart(m["Catalog_Path"], m["Catalog_Name"], MART_CONN)
    print(locator)

# 3) Exporta modelos filtrados para output/xml
resultado = mart_exportar_todos_xml(
    mart_url=MART_URL,
    bearer_token=BEARER,
    xsrf_token=XSRF,
    mart_conn_str=MART_CONN,
    caminho_saida_dir="output/xml",
    caminho_temp_dir="temp",
    data_atualizacao_min="03/11/2026",
)
print(resultado)
```

Para executar pelo bloco `__main__`, configure o `.env` e rode:

```bash
uv run python erwin_save_xml.py
```

## Arquitetura COM Utilizada

| ProgID | DLL | Uso |
| --- | --- | --- |
| `erwin9.SCAPI.9.0` | `erwin9.dll` | Abrir modelo Mart e salvar `.erwin` temporario |
| `erwin9.SCAPI.PropertyBag.9.0` | `erwin9.dll` | Configuracao de `PersistenceUnits.Create(...)` |
| `ERXML.XMLERwinLink` | `ERXML.dll` | Conversao `.erwin` -> XML (`StandAloneExport`) |

## Estrutura do Projeto

```text
erwinDM_save_to_xml/
├── erwin_save_xml.py
├── output/
│   ├── mart_report/
│   │   └── mart_models.xml
│   └── xml/
│       └── *.xml
├── temp/
├── documentation/
├── PLANO.md
└── README.md
```
