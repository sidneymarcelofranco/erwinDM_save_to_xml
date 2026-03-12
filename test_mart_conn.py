"""
test_mart_conn.py
-----------------
Valida os recursos necessarios para executar erwin_save_xml.py:

  1. Conectividade de rede com o Mart Server
  2. Endpoint CSRF  (GET /MartServer/csrf)
  3. Login JWT      (POST /MartServerCloud/jwt/authenticate/login)
  4. API de modelos (GET /MartServer/api/report/generateReport/Models)
  5. COM: SCAPI     (erwin9.SCAPI.9.0)
  6. COM: ERXML     (ERXML.XMLERwinLink)

Uso:
    python test_mart_conn.py

Console: apenas resultado de cada teste (OK / FAIL).
Log detalhado: log/test_mart_conn_YYYY-MM-DD.log
"""

import os
import sys
import ssl
import json
import logging
import urllib.request
import urllib.error
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR  = os.path.join(BASE_DIR, "log")

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

def _configurar_log() -> logging.Logger:
    os.makedirs(LOG_DIR, exist_ok=True)
    nome_arquivo = datetime.now().strftime("test_mart_conn_%Y-%m-%d.log")
    caminho_log  = os.path.join(LOG_DIR, nome_arquivo)

    fmt_arquivo = logging.Formatter(
        "%(asctime)s [%(levelname)-8s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    fh = logging.FileHandler(caminho_log, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt_arquivo)

    logger = logging.getLogger("test_mart_conn")
    logger.setLevel(logging.DEBUG)
    logger.addHandler(fh)
    logger.propagate = False
    return logger


# ---------------------------------------------------------------------------
# Carrega .env (mesmo loader do script principal)
# ---------------------------------------------------------------------------

def _carregar_dotenv(caminho: str) -> None:
    if not os.path.isfile(caminho):
        return
    with open(caminho, "r", encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or linha.startswith("#") or "=" not in linha:
                continue
            chave, valor = linha.split("=", 1)
            chave = chave.strip()
            valor = valor.strip().strip('"').strip("'")
            if chave and chave not in os.environ:
                os.environ[chave] = valor


# ---------------------------------------------------------------------------
# SSL (mesmo padrao do script principal)
# ---------------------------------------------------------------------------

_SSL_CTX = ssl.create_default_context()
_SSL_CTX.check_hostname = False
_SSL_CTX.verify_mode    = ssl.CERT_NONE

# ---------------------------------------------------------------------------
# Helpers de console (mínimo)
# ---------------------------------------------------------------------------

def _ok(label: str) -> None:
    print(f"  [OK  ] {label}")

def _fail(label: str, detalhe: str = "") -> None:
    sufixo = f" — {detalhe}" if detalhe else ""
    print(f"  [FAIL] {label}{sufixo}")

# ---------------------------------------------------------------------------
# Base URL
# ---------------------------------------------------------------------------

def _base_url() -> str:
    mart_url = os.getenv("MART_URL", "").strip()
    if not mart_url:
        raise RuntimeError("MART_URL nao configurado no .env")
    host = mart_url
    if "://" in host:
        host = host.split("://", 1)[1]
    host = host.strip().strip("/")
    protocolo = os.getenv("MART_PROTOCOL", "https").strip().lower()
    return f"{protocolo}://{host}"

# ---------------------------------------------------------------------------
# Testes
# ---------------------------------------------------------------------------

def test_csrf(log: logging.Logger) -> str | None:
    base = _base_url()
    url  = f"{base}/MartServer/csrf"
    log.info("GET %s", url)
    try:
        req = urllib.request.Request(url, method="GET")
        req.add_header("accept", "*/*")
        xsrf = None
        try:
            with urllib.request.urlopen(req, context=_SSL_CTX, timeout=10) as r:
                xsrf = r.headers.get("XSRF-TOKEN")
                log.debug("HTTP %s", r.status)
        except urllib.error.HTTPError as e:
            xsrf = e.headers.get("XSRF-TOKEN")
            log.debug("HTTP %s (esperado) — XSRF-TOKEN no header", e.code)
        if xsrf:
            log.info("XSRF-TOKEN recebido: %s...", xsrf[:8])
            _ok("CSRF token")
            return xsrf
        log.error("XSRF-TOKEN ausente no header da resposta")
        _fail("CSRF token", "header XSRF-TOKEN nao retornado")
        return None
    except Exception as e:
        log.exception("Erro ao obter CSRF token")
        _fail("CSRF token", str(e))
        return None


def test_login(xsrf_token: str, log: logging.Logger) -> str | None:
    username = os.getenv("USER_MART", "").strip()
    password = os.getenv("PASS_MART", "").strip()
    if not username or not password:
        log.error("USER_MART ou PASS_MART nao configurados no .env")
        _fail("Login JWT", "USER_MART/PASS_MART ausentes no .env")
        return None
    base = _base_url()
    url  = f"{base}/MartServerCloud/jwt/authenticate/login"
    log.info("POST %s (usuario: %s)", url, username)
    try:
        payload = json.dumps({"username": username, "password": password}).encode("utf-8")
        req = urllib.request.Request(url, data=payload, method="POST")
        req.add_header("accept", "*/*")
        req.add_header("Content-Type", "application/json")
        req.add_header("X-XSRF-TOKEN", xsrf_token)
        with urllib.request.urlopen(req, context=_SSL_CTX, timeout=10) as r:
            log.debug("HTTP %s", r.status)
            corpo = r.read().decode("utf-8")
        dados  = json.loads(corpo)
        bearer = dados.get("id_token", "")
        if bearer:
            log.info("Bearer token recebido: %s...", bearer[:12])
            _ok("Login JWT")
            return bearer
        log.error("id_token ausente na resposta: %s", corpo[:200])
        _fail("Login JWT", "id_token ausente na resposta")
        return None
    except urllib.error.HTTPError as e:
        log.error("HTTP %s: %s", e.code, e.reason)
        _fail("Login JWT", f"HTTP {e.code} — verifique credenciais")
        return None
    except Exception as e:
        log.exception("Erro no login JWT")
        _fail("Login JWT", str(e))
        return None


def test_models_api(bearer: str, xsrf_token: str, log: logging.Logger) -> bool:
    base = _base_url()
    url  = f"{base}/MartServer/api/report/generateReport/Models"
    log.info("GET %s", url)
    try:
        req = urllib.request.Request(url)
        req.add_header("accept", "*/*")
        req.add_header("Authorization", f"Bearer {bearer}")
        req.add_header("X-XSRF-TOKEN", xsrf_token)
        with urllib.request.urlopen(req, context=_SSL_CTX, timeout=15) as r:
            conteudo = r.read()
            log.info("HTTP %s — %s bytes", r.status, f"{len(conteudo):,}")
        _ok("API de modelos")
        return True
    except urllib.error.HTTPError as e:
        log.error("HTTP %s: %s", e.code, e.reason)
        _fail("API de modelos", f"HTTP {e.code}")
        return False
    except Exception as e:
        log.exception("Erro ao consultar API de modelos")
        _fail("API de modelos", str(e))
        return False


def test_com_scapi(log: logging.Logger) -> bool:
    try:
        import win32com.client
        win32com.client.Dispatch("erwin9.SCAPI.9.0")
        log.info("erwin9.SCAPI.9.0 instanciado com sucesso")
        _ok("COM: erwin9.SCAPI.9.0")
        return True
    except ImportError:
        log.error("pywin32 nao instalado")
        _fail("COM: erwin9.SCAPI.9.0", "pip install pywin32")
        return False
    except Exception as e:
        log.error("COM SCAPI falhou: %s", e)
        _fail("COM: erwin9.SCAPI.9.0", str(e))
        return False


def test_com_erxml(log: logging.Logger) -> bool:
    try:
        import win32com.client
        win32com.client.Dispatch("ERXML.XMLERwinLink")
        log.info("ERXML.XMLERwinLink instanciado com sucesso")
        _ok("COM: ERXML.XMLERwinLink")
        return True
    except ImportError:
        log.error("pywin32 nao instalado")
        _fail("COM: ERXML.XMLERwinLink", "pip install pywin32")
        return False
    except Exception as e:
        log.error("COM ERXML falhou: %s", e)
        _fail("COM: ERXML.XMLERwinLink", str(e))
        return False

# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

def main() -> None:
    _carregar_dotenv(os.path.join(BASE_DIR, ".env"))
    log = _configurar_log()
    log.info("Iniciando validacao de pre-requisitos")

    try:
        base = _base_url()
        log.info("Mart URL: %s", base)
    except RuntimeError as e:
        print(f"[ERRO] {e}")
        sys.exit(1)

    nome_log = datetime.now().strftime("test_mart_conn_%Y-%m-%d.log")
    print(f"Validando pre-requisitos — log: log/{nome_log}\n")

    resultados: dict[str, bool] = {}

    xsrf = test_csrf(log)
    resultados["CSRF token"] = xsrf is not None

    if xsrf:
        bearer = test_login(xsrf, log)
        resultados["Login JWT"] = bearer is not None
    else:
        resultados["Login JWT"] = False
        log.warning("Login JWT pulado — CSRF token nao disponivel")
        _fail("Login JWT", "pulado")
        bearer = None

    if bearer and xsrf:
        resultados["API de modelos"] = test_models_api(bearer, xsrf, log)
    else:
        resultados["API de modelos"] = False
        log.warning("API de modelos pulada — autenticacao falhou")
        _fail("API de modelos", "pulado")

    resultados["COM: SCAPI"] = test_com_scapi(log)
    resultados["COM: ERXML"] = test_com_erxml(log)

    tudo_ok = all(resultados.values())
    log.info("Resultado final: %s", "OK" if tudo_ok else "FALHA")

    print()
    if tudo_ok:
        print("  Tudo pronto. Execute: python erwin_save_xml.py")
    else:
        falhas = [k for k, v in resultados.items() if not v]
        print(f"  Corrija antes de executar: {', '.join(falhas)}")
    print()

    sys.exit(0 if tudo_ok else 1)


if __name__ == "__main__":
    main()
