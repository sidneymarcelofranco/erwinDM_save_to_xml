# erwin Data Modeler Read Mart Portal and Save to XML

Utilitario Python para:
- listar modelos do Mart Server em XML (`mart_models.xml`)
- filtrar modelos por `UpdatedOn` (data exata, mínima ou intervalo)
- montar locator SCAPI com `Catalog_Path` + `Catalog_Name`
- exportar modelos do Mart para XML legível
- exportar modelo local `.erwin` para XML legível

## Status do Projeto

✅ **Implementação Completa**
- [x] Integração com API do Mart Server (listar modelos)
- [x] Três modos de filtro por data (exata, mínima, intervalo)
- [x] Abertura de modelos Mart via SCAPI
- [x] Conversão para XML via ERXML
- [x] Credenciais externalizadas em `.env`
- [x] Injeção automática de credenciais na string de conexão
- [x] Sobrescrita automática de XMLs sem prompts
- [x] Estrutura de pastas organizada (mart_report, xml, temp)
- [x] Tratamento de erros COM e HTTP

**Pronto para usar:** Execute `uv run python erwin_save_xml.py` com `.env` preenchido.

## Requisitos

- erwin DM 15.2 instalado (registra as DLLs COM)
- Python >= 3.14

## Instalação

### Opção 1: Com `uv` (recomendado)

Se tiver [uv](https://docs.astral.sh/uv/) instalado:

```bash
uv sync
```

### Opção 2: Com `pip` (padrão Python)

Se preferir usar apenas pip:

```bash
pip install -r requirements.txt
python erwin_save_xml.py
```

**Nota:** Se receber erro de COM no primeiro `python erwin_save_xml.py`, execute (apenas uma vez):

```bash
python -c "import pywin32_postinstall; pywin32_postinstall.main(['-install'])"
```

Depois tente novamente. Na maioria dos casos **não é necessário** esse passo extra.

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
- `MART_CONN_STR` (base sem `UID` e `PSW`)
- `USER_MART` (opcional, mas recomendado)
- `PASS_MART` (opcional, mas recomendado)
- `MART_UPDATED_ON_EXACT` (opcional, formato `MM/DD/YYYY`)
- `MART_UPDATED_ON_MIN` (opcional, formato `MM/DD/YYYY`)
- `MART_UPDATED_ON_MAX` (opcional, formato `MM/DD/YYYY`)

Regras de filtro de data:
- Exata: use apenas `MART_UPDATED_ON_EXACT`
- Minima: use `MART_UPDATED_ON_MIN`
- Intervalo: use `MART_UPDATED_ON_MIN` + `MART_UPDATED_ON_MAX`
- `EXACT` nao pode ser combinado com `MIN/MAX`

Regras da conexao Mart:
- `MART_CONN_STR` pode ser informado sem credenciais
- se `USER_MART` e `PASS_MART` forem informados, o script injeta `UID` e `PSW` automaticamente
- se `MART_CONN_STR` ja vier com `UID`/`PSW`, eles sao removidos e substituidos pelos valores de `USER_MART`/`PASS_MART` quando presentes

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

## Organização de Pastas de Saída

Quando executado pelo `__main__`:

- **Relatório de modelos**:
  `output/mart_report/mart_models.xml`

- **XMLs exportados dos modelos** (com estrutura espelhada do Mart):
  `output/xml/<Catalog_Path>/<Catalog_Name>.xml`
  
  Exemplo: `output/xml/Mart/Modelos/eMovies.xml`

- **Temporários `.erwin`** (removidos após exportação):
  `temp/*.erwin` (sempre vazio após conclusão)

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

## Como Executar

Configure o `.env` com seus dados sensíveis (veja seção **Configuração Segura**) e execute:

### Com `uv`:

```bash
uv run python erwin_save_xml.py
```

### Com `pip`:

```bash
python erwin_save_xml.py
```

### Cenários de Execução

#### 1. Exportar **TODOS** os modelos (sem filtro de data)

Deixe os três campos de data vazios no `.env`:

```env
MART_UPDATED_ON_EXACT=
MART_UPDATED_ON_MIN=
MART_UPDATED_ON_MAX=
```

Resultado: Todos os modelos da lista original serão exportados.

#### 2. Exportar por data **EXATA**

Configure apenas `MART_UPDATED_ON_EXACT`:

```env
MART_UPDATED_ON_EXACT=03/03/2026
MART_UPDATED_ON_MIN=
MART_UPDATED_ON_MAX=
```

Resultado: Apenas modelos com `UpdatedOn` == `03/03/2026` serão exportados.

#### 3. Exportar por data **MÍNIMA** (>=)

Configure apenas `MART_UPDATED_ON_MIN`:

```env
MART_UPDATED_ON_EXACT=
MART_UPDATED_ON_MIN=03/11/2026
MART_UPDATED_ON_MAX=
```

Resultado: Apenas modelos com `UpdatedOn` >= `03/11/2026` serão exportados.

#### 4. Exportar por **INTERVALO** de datas (MIN até MAX)

Configure `MART_UPDATED_ON_MIN` e `MART_UPDATED_ON_MAX`:

```env
MART_UPDATED_ON_EXACT=
MART_UPDATED_ON_MIN=03/06/2026
MART_UPDATED_ON_MAX=03/09/2026
```

Resultado: Apenas modelos com `UpdatedOn` entre `03/06/2026` e `03/09/2026` (inclusive) serão exportados.

### Fluxo de Execução

1. **Carrega configuração**: Lê `.env` com credenciais, URL, protocolos e filtros
2. **Lista modelos**: Chamar API do Mart Server (endpoint `/MartServer/api/report/generateReport/Models`)
3. **Salva relatório**: Guarda lista completa em `output/mart_report/mart_models.xml`
4. **Filtra por data** (opcional): Se algum filtro estiver preenchido, aplica a seleção
5. **Exporta por modelo**: Para cada modelo filtrado:
   - Monta locator SCAPI com credenciais
   - Abre via SCAPI e salva `.erwin` temporário
   - Converte `.erwin` para XML via ERXML
   - Limpa arquivo temporário
6. **Resultado**: XMLs exportados em `output/xml/`

### Arquivos de Saída

Após execução, a estrutura de pastas replica a organização do Mart Server:

| Caminho | Propósito |
| --- | --- |
| `output/mart_report/mart_models.xml` | Lista completa de modelos do Mart (sem filtro) |
| `output/xml/<Catalog_Path>/<Catalog_Name>.xml` | Cada modelo exportado com estrutura espelhada do Mart |
| `temp/` | Vazio (arquivos `.erwin` são limpados após conversão) |

**Exemplo prático:**

Modelo com `Catalog_Path = "Mart/Modelos"` e `Catalog_Name = "eMovies"`:
```
output/xml/Mart/Modelos/eMovies.xml
```

Modelo com `Catalog_Path = "Mart/Ambiente/Homologacao"` e `Catalog_Name = "exemploMongo"`:
```
output/xml/Mart/Ambiente/Homologacao/exemploMongo.xml
```

Os diretórios são criados automaticamente se não existirem.

### Tratamento de Erros

- **Erro de conexão Mart**: Verifique `MART_URL`, `MART_BEARER_TOKEN` e `MART_XSRF_TOKEN`
- **Erro de credenciais**: Verifique `USER_MART` e `PASS_MART` contra o Mart Server
- **Erro de protocolo**: Revise se `MART_PROTOCOL` é `https` ou `http` conforme seu servidor
- **Arquivo XML já existe**: O script sobrescreve automaticamente sem pedir confirmação

### Exemplo de Resultado

Após executar com `MART_UPDATED_ON_EXACT=03/03/2026`:

```
output/
├── mart_report/
│   └── mart_models.xml                          # Lista com ~9 modelos (conforme API)
└── xml/
    └── Mart/
        ├── Modelos/
        │   ├── Contrato_Pessoa_Produto.xml      # 3 modelos com UpdatedOn = 03/03/2026
        │   ├── eMovies.xml
        │   └── ...
        ├── Ambiente/
        │   └── Homologacao/
        │       └── exemploMongo.xml             # Se tivesse UpdatedOn = 03/03/2026
        └── API2/
            └── ...
temp/                                            # Vazio (limpo após export)
```

Cada XML está organizado exatamente como no Mart Server, facilitando a navegação e sincronização.

## Arquitetura COM Utilizada

| ProgID | DLL | Uso |
| --- | --- | --- |
| `erwin9.SCAPI.9.0` | `erwin9.dll` | Abrir modelo Mart e salvar `.erwin` temporario |
| `erwin9.SCAPI.PropertyBag.9.0` | `erwin9.dll` | Configuracao de `PersistenceUnits.Create(...)` |
| `ERXML.XMLERwinLink` | `ERXML.dll` | Conversao `.erwin` -> XML (`StandAloneExport`) |

## Estrutura do Projeto

```text
erwinDM_save_to_xml/
├── erwin_save_xml.py              # Módulo principal
├── test_mart_conn.py              # Teste de conectividade (opcional)
├── .env                           # Credenciais e filtros (NÃO versionar)
├── .env.example                   # Template de .env
├── requirements.txt               # Dependências Python
├── input/
│   └── eMovies.erwin              # Exemplo de modelo local
├── output/
│   ├── mart_report/
│   │   └── mart_models.xml        # Lista de modelos (API response)
│   └── xml/
│       └── Mart/                  # Estrutura espelhada do servidor
│           ├── Modelos/
│           │   ├── eMovies.xml
│           │   └── ...
│           └── Ambiente/
│               └── Homologacao/
│                   └── ...
├── temp/
│   └── *.erwin                    # Temporários (limpados após export)
├── documentation/                 # Docs da API do erwin DM
├── PLANO.md
└── README.md
```

## Diagnóstico Rápido

### Verificar Instalação

Antes de executar, valide se tudo está configurado:

```powershell
# Verificar versão do Python
python --version          # Deve ser >= 3.14

# Verificar se pywin32 está instalado
python -c "import win32com; print('✅ win32com OK')"

# Verificar se erwin DM está registrado
python -c "import win32com.client; app = win32com.client.Dispatch('erwin9.SCAPI.9.0'); print('✅ SCAPI OK')"

# Verificar se ERXML está disponível
python -c "import win32com.client; xml = win32com.client.Dispatch('ERXML.XMLERwinLink'); print('✅ ERXML OK')"
```

Se algum comando falhar, volte à seção **Instalação** e reinstale corretamente.

### Validar conectividade Mart

Execute o script de teste de conexão (se disponível):

```bash
python test_mart_conn.py
```

Verifique:
- ✅ HTTP 200 da API do Mart Server
- ✅ Credenciais corretas no `.env`
- ✅ Tokens (Bearer + XSRF) válidos

### Validar estrutura de pastas

Após primeira execução, confirme que existem:

```
output/mart_report/mart_models.xml       # Deve ter ~2500 bytes
output/xml/                              # Deve conter XMLs dos modelos
temp/                                    # Deve estar vazio
```

### Erros Comuns

| Erro | Solução |
| --- | --- |
| `FileNotFoundError: .env` | Crie `.env` a partir de `.env.example` |
| `HTTP 401 Unauthorized` | Verifique `MART_BEARER_TOKEN` e `MART_XSRF_TOKEN` |
| `COM Exception` | Verifique se erwin DM 15.2 está instalado e DLLs registradas; ou registre novamente com `python -m pip show pywin32` + postinstall |
| `No module named 'win32com'` | Rode `pip install -r requirements.txt` e execute `python -m pip install --upgrade pywin32` |
| `No module named 'pywin32_postinstall'` | Execute `python -m pip install --upgrade pywin32` e depois `python Scripts/pywin32_postinstall.py -install` dentro de site-packages |
| Sem modelos exportados | Confirme se filtros de data correspondem aos modelos do Mart |
| `Permissão negada ao sobrescrever XML` | Feche arquivos XML abertos em outros programas |
