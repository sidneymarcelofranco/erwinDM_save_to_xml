# Quick Start

## Pré-requisitos

- Windows com **erwin DM 15.2** instalado
- **Python >= 3.14**
- **Streamlit** (instalado via `requirements.txt`)

---

## 1. Instalar dependências

```bash
pip install -r requirements.txt
```

> Erro COM após instalar? Rode uma única vez:
>
> ```bash
> python -c "import pywin32_postinstall; pywin32_postinstall.main(['-install'])"
> ```

---

## 2. Configurar `.env`

```bash
copy .env.example .env
```

Edite `.env` com seus dados:

```env
MART_URL=mart.suaempresa.com.br
USER_MART=Administrator
PASS_MART=sua_senha
MART_CONN_STR=TRC=NO;PRT=443;ASR=MartServer;SSL=YES
MART_PROTOCOL=https
```

> `MART_BEARER_TOKEN` e `MART_XSRF_TOKEN` são **opcionais** — se vazios, o token é gerado automaticamente via login.

---

## 3. (Opcional) Definir filtro de data

Para exportar apenas modelos atualizados em uma data específica:

```env
MART_UPDATED_ON_EXACT=03/03/2026
```

Deixe os três campos em branco para exportar **todos** os modelos.

---

## 4. Validar antes de executar

**Sempre execute o teste de conectividade antes de rodar o script principal.**
Ele verifica rede, autenticação e os objetos COM do erwin DM.

```bash
python test_mart_conn.py
```

Saída esperada (console mínimo):

```text
Validando pre-requisitos — log: log/test_mart_conn_2026-03-12.log

  [OK  ] CSRF token
  [OK  ] Login JWT
  [OK  ] API de modelos
  [OK  ] COM: erwin9.SCAPI.9.0
  [OK  ] COM: ERXML.XMLERwinLink

  Tudo pronto. Execute: python erwin_save_xml.py
```

Se algum item aparecer com `[FAIL]`, corrija antes de prosseguir.
O log detalhado fica em `log/test_mart_conn_YYYY-MM-DD.log`.

---

## 5. Executar

```bash
python erwin_save_xml.py
```

Ou execute pela interface web (Streamlit), a partir da raiz do projeto:

```bash
streamlit run app/app.py
```

Com `uv`:

```bash
uv run streamlit run app/app.py
```

Saída no console (mínima):

```text
RESUMO
----------------------------------------
  [OK   ] eMovies
  [OK   ] Contrato_Pessoa_Produto

  2 exportado(s)  |  0 com falha
  Log detalhado: log/erwin_2026-03-12.log
```

Todos os detalhes de execução ficam em `log/erwin_YYYY-MM-DD.log`.

---

## Arquivos gerados

```text
output/
├── mart_report/mart_models.xml          # lista de modelos do Mart
└── xml/Mart/<pasta>/<modelo>.xml        # XMLs exportados

log/
├── test_mart_conn_YYYY-MM-DD.log        # log do teste de conectividade
└── erwin_YYYY-MM-DD.log                 # log da execucao principal
```
