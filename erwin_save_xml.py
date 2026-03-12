import win32com.client
import os
import json
import logging
import urllib.request
import ssl
import xml.etree.ElementTree as ET
from datetime import datetime


# =============================================================================
# CONSTANTES - ProgIDs suportados pelo erwin DM
#
# ERXML (exportacao/importacao de XML legivel):
#   'ERXML.XMLERwinLink'       -> interface nativa de export XML (ERXML.dll)
#   'ERXML.XMLERwinLink.9.0'   -> versao versionada
#
# ProgIDs confirmados no registro do Windows para erwin 15.2.
# =============================================================================
ERWIN_XML_PROGID = 'ERXML.XMLERwinLink'

# Contexto SSL com verificacao desabilitada para ambientes com cert. interno.
# Criado uma vez e reutilizado em todas as requisicoes HTTP.
_SSL_CTX = ssl.create_default_context()
_SSL_CTX.check_hostname = False
_SSL_CTX.verify_mode = ssl.CERT_NONE

# Formato de data do campo UpdatedOn retornado pela API do Mart.
# Exemplo: "02/10/2026 12:17:35 PM"
_MART_DATE_FMT = "%m/%d/%Y %I:%M:%S %p"

# Logger do modulo — NullHandler por padrao (padrao de biblioteca Python).
# Os handlers reais (arquivo + console) sao configurados pelo __main__.
_log = logging.getLogger("erwin_save_xml")
_log.addHandler(logging.NullHandler())


# =============================================================================
# LOGGING
# =============================================================================

def _configurar_log(log_dir: str) -> None:
    """
    Configura dois handlers para o logger do modulo:
    - Arquivo : log_dir/erwin_YYYY-MM-DD.log  (nivel DEBUG — tudo)
    - Console : somente WARNING e acima (erros que exigem atencao)
    """
    os.makedirs(log_dir, exist_ok=True)
    nome_arquivo = datetime.now().strftime("erwin_%Y-%m-%d.log")
    caminho_log  = os.path.join(log_dir, nome_arquivo)

    fmt_arquivo  = logging.Formatter(
        "%(asctime)s [%(levelname)-8s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    fmt_console  = logging.Formatter("[%(levelname)s] %(message)s")

    fh = logging.FileHandler(caminho_log, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt_arquivo)

    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)
    ch.setFormatter(fmt_console)

    _log.setLevel(logging.DEBUG)
    _log.addHandler(fh)
    _log.addHandler(ch)
    _log.propagate = False

    _log.info("Log iniciado: %s", caminho_log)


# =============================================================================
# HELPERS
# =============================================================================

def _normalizar_host_mart(mart_url: str) -> str:
    """Normaliza MART_URL para host sem protocolo e sem barra final."""
    host = mart_url.strip()
    if "://" in host:
        host = host.split("://", 1)[1]
    return host.strip().strip("/")


def _montar_base_url_api(mart_url: str, protocolo: str = "https") -> str:
    host = _normalizar_host_mart(mart_url)
    prot = (protocolo or "https").strip().lower()
    return f"{prot}://{host}"


def _parse_data_filtro(data_str: str, nome_campo: str) -> datetime:
    try:
        return datetime.strptime(data_str, "%m/%d/%Y")
    except ValueError as e:
        raise ValueError(f"{nome_campo} invalida: '{data_str}'. Use MM/DD/YYYY") from e


def _carregar_dotenv(caminho_env: str) -> None:
    """Carrega variaveis de um arquivo .env simples (KEY=VALUE)."""
    if not os.path.isfile(caminho_env):
        return

    with open(caminho_env, "r", encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or linha.startswith("#") or "=" not in linha:
                continue
            chave, valor = linha.split("=", 1)
            chave = chave.strip()
            valor = valor.strip().strip('"').strip("'")
            if chave and chave not in os.environ:
                os.environ[chave] = valor


def _env_obrigatorio(nome: str) -> str:
    valor = os.getenv(nome, "").strip()
    if not valor:
        raise RuntimeError(f"Variavel obrigatoria ausente: {nome}")
    return valor


# =============================================================================
# AUTENTICACAO
# =============================================================================

def _obter_xsrf_token(base_url: str) -> str:
    """GET /MartServer/csrf — retorna XSRF-TOKEN do header (resposta pode ser 403/401)."""
    csrf_url = f"{base_url}/MartServer/csrf"
    req = urllib.request.Request(csrf_url, method="GET")
    req.add_header("accept", "*/*")
    xsrf_token = None
    try:
        with urllib.request.urlopen(req, context=_SSL_CTX, timeout=30) as resp:
            xsrf_token = resp.headers.get("XSRF-TOKEN")
    except urllib.error.HTTPError as e:
        # O endpoint retorna 403/401 mas envia XSRF-TOKEN no header mesmo assim
        xsrf_token = e.headers.get("XSRF-TOKEN")
    if not xsrf_token:
        raise RuntimeError("XSRF-TOKEN nao encontrado no header da resposta do CSRF")
    return xsrf_token


def _gerar_token_mart(base_url: str, username: str, password: str) -> tuple[str, str]:
    """
    Autentica no Mart Server e retorna (bearer_token, xsrf_token).

    Etapa 1: GET /MartServer/csrf  -> captura XSRF-TOKEN do header (retorna 403,
             mas o header XSRF-TOKEN e enviado mesmo assim).
    Etapa 2: POST /MartServerCloud/jwt/authenticate/login -> retorna { "id_token": "..." }.
    """
    # Etapa 1: XSRF token
    _log.info("Obtendo XSRF token: %s/MartServer/csrf", base_url)
    xsrf_token = _obter_xsrf_token(base_url)
    _log.info("XSRF-TOKEN obtido")

    # Etapa 2: login JWT
    login_url = f"{base_url}/MartServerCloud/jwt/authenticate/login"
    _log.info("Autenticando: %s", login_url)
    payload = json.dumps({"username": username, "password": password}).encode("utf-8")
    req2 = urllib.request.Request(login_url, data=payload, method="POST")
    req2.add_header("accept", "*/*")
    req2.add_header("Content-Type", "application/json")
    req2.add_header("X-XSRF-TOKEN", xsrf_token)
    with urllib.request.urlopen(req2, context=_SSL_CTX, timeout=30) as resp2:
        corpo = resp2.read().decode("utf-8")
    try:
        dados = json.loads(corpo)
        bearer = dados.get("id_token") or ""
    except Exception:
        bearer = corpo.strip()

    if not bearer:
        raise RuntimeError("Bearer token nao encontrado na resposta do login")
    _log.info("Autenticacao concluida com sucesso")
    return bearer, xsrf_token


# =============================================================================
# CONEXAO / LOCATOR
# =============================================================================

def _montar_conn_str_mart(
    conn_str_base: str,
    user_mart: str | None,
    pass_mart: str | None,
    mart_host: str,
) -> str:
    """
    Monta a connection string final do Mart aplicando UID/PSW separadamente.

    Regras:
    - Remove UID/PSW/SRV existentes em conn_str_base.
    - Usa o host informado em MART_URL para SRV.
    - Se USER_MART e PASS_MART vierem preenchidos, aplica ambos no final.
    - Se nao vierem, mantem apenas o base sem UID/PSW.
    """
    partes_limpa = []
    for parte in conn_str_base.split(";"):
        p = parte.strip()
        if not p:
            continue
        chave = p.split("=", 1)[0].strip().upper()
        if chave in {"UID", "PSW", "SRV"}:
            continue
        partes_limpa.append(p)

    partes_limpa.append(f"SRV={_normalizar_host_mart(mart_host)}")

    if user_mart and pass_mart:
        partes_limpa.append(f"UID={user_mart}")
        partes_limpa.append(f"PSW={pass_mart}")

    return ";".join(partes_limpa)


def montar_locator_mart(catalog_path: str, catalog_name: str, mart_conn_str: str) -> str:
    """
    Monta o locator SCAPI no formato:
    mart://Mart/<Catalog_Path_sem_prefixo>/<Catalog_Name>?<mart_conn_str>
    """
    path_relativo = catalog_path.removeprefix("Mart/")
    return f"mart://Mart/{path_relativo}/{catalog_name}?{mart_conn_str}"


# =============================================================================
# XML
# =============================================================================

def _formatar_xml(caminho: str) -> None:
    """Le um arquivo XML, indenta e reescreve no mesmo caminho (UTF-8)."""
    tree = ET.parse(caminho)
    ET.indent(tree, space="  ")
    tree.write(caminho, encoding="utf-8", xml_declaration=True)


# =============================================================================
# API REST
# =============================================================================

def _fetch_modelos_mart(mart_url: str, bearer_token: str, xsrf_token: str) -> bytes | None:
    """
    Consulta GET /MartServer/api/report/generateReport/Models e retorna os
    bytes da resposta, ou None em caso de falha.
    """
    protocolo = os.getenv("MART_PROTOCOL", "https")
    base_url = _montar_base_url_api(mart_url, protocolo)
    endpoint = f"{base_url}/MartServer/api/report/generateReport/Models"
    _log.info("Consultando Mart: %s", endpoint)
    try:
        req = urllib.request.Request(endpoint)
        req.add_header("accept", "*/*")
        req.add_header("Authorization", f"Bearer {bearer_token}")
        req.add_header("X-XSRF-TOKEN", xsrf_token)
        with urllib.request.urlopen(req, context=_SSL_CTX, timeout=30) as resp:
            status   = resp.status
            conteudo = resp.read()
        _log.info("HTTP %s — %s bytes recebidos", status, f"{len(conteudo):,}")
        return conteudo
    except urllib.error.HTTPError as e:
        _log.error("HTTP %s: %s", e.code, e.reason)
    except urllib.error.URLError as e:
        _log.error("Falha de conexao: %s", e.reason)
    except Exception:
        _log.exception("Erro inesperado em _fetch_modelos_mart")
    return None


# =============================================================================
# COM / EXPORT
# =============================================================================

def _exportar_via_com(locator: str, caminho_saida_xml: str) -> bool:
    """
    Chama ERXML.XMLERwinLink.StandAloneExport e pos-processa o XML gerado.

    Aceita tanto caminhos locais quanto locators Mart (Mart://Mart/...).
    """
    if not caminho_saida_xml.lower().endswith(".xml"):
        _log.warning("O caminho de saida nao possui extensao .xml.")
    try:
        # Evita dialogo de confirmacao de sobrescrita no COM.
        if os.path.exists(caminho_saida_xml):
            _log.info("Removendo arquivo existente: %s", caminho_saida_xml)
            # os.unlink ou os.remove — deleta um único arquivo do disco.
            os.unlink(caminho_saida_xml)

        _log.info("Exportando via %s", ERWIN_XML_PROGID)
        _log.debug("Entrada : %s", locator)
        _log.debug("Saida   : %s", caminho_saida_xml)
        obj_xml = win32com.client.Dispatch(ERWIN_XML_PROGID)
        obj_xml.StandAloneExport(locator, caminho_saida_xml, 0)
        # COM nao retorna status — formata diretamente; erro de I/O indica falha
        _formatar_xml(caminho_saida_xml)
        tamanho = os.path.getsize(caminho_saida_xml)
        _log.info("XML exportado: %s (%s bytes)", caminho_saida_xml, f"{tamanho:,}")
        return True
    except Exception:
        _log.exception("Erro na exportacao XML: %s", caminho_saida_xml)
        return False


def _mart_exportar_modelo(mart_locator: str, caminho_saida_xml: str,
                          caminho_temp: str) -> bool:
    """
    Abre um modelo do Mart via SCAPI, salva como .erwin em caminho_temp,
    exporta para XML via ERXML e remove o arquivo temporario.

    StandAloneExport nao suporta locators Mart — por isso o modelo e baixado
    primeiro como arquivo local antes da conversao XML.
    """
    try:
        _log.info("SCAPI: conectando ao Mart")
        scapi    = win32com.client.Dispatch('erwin9.SCAPI.9.0')
        prop_bag = win32com.client.Dispatch('erwin9.SCAPI.PropertyBag.9.0')
        prop_bag.Add("Model_Type", "Combined")

        scapi.PersistenceUnits.Create(prop_bag)
        pu = scapi.PersistenceUnits.Add(mart_locator, "RDO=No")

        if scapi.Sessions.Count > 0:
            scapi.Sessions.Clear()

        # Abre M0 (logico) e M1 (fisico) — padrao requerido para modelo Combined
        sessao_m0 = scapi.Sessions.Add()
        sessao_m1 = scapi.Sessions.Add()
        ret_m0 = sessao_m0.Open(pu, 0)
        ret_m1 = sessao_m1.Open(pu, 1)
        _log.debug("SCAPI: sessao M0=%s  M1=%s", ret_m0, ret_m1)

        _log.info("SCAPI: salvando local em %s", caminho_temp)
        # Salvando em arquivo local temporario, usar disposicao de arquivo.
        pu.Save(caminho_temp, "OVF=Yes")
        _log.debug("SCAPI: Save() concluido")
        scapi.Sessions.Clear()

        tamanho_temp = os.path.getsize(caminho_temp)
        _log.debug("SCAPI: arquivo temp %s bytes", f"{tamanho_temp:,}")
        if tamanho_temp == 0:
            _log.error("SCAPI: Save() nao gravou dados no arquivo temp")
            os.unlink(caminho_temp)
            return False

    except Exception:
        _log.exception("SCAPI Mart: falha ao abrir/salvar modelo")
        if os.path.exists(caminho_temp):
            os.unlink(caminho_temp)
        return False

    resultado = _exportar_via_com(caminho_temp, caminho_saida_xml)
    if os.path.exists(caminho_temp):
        os.unlink(caminho_temp)
    return resultado


# =============================================================================
# API PUBLICA
# =============================================================================

def listar_modelos_mart(mart_url: str, bearer_token: str, xsrf_token: str,
                        caminho_saida: str) -> bool:
    """
    Consulta a API REST do erwin Mart Server para listar os modelos disponiveis
    e salva a resposta XML indentada em arquivo.

    Endpoint: GET /MartServer/api/report/generateReport/Models

    Parametros:
        mart_url     : str -- URL base do Mart Server.
                             Ex: 'https://mart.empresa.com.br'
        bearer_token : str -- Token Bearer de autenticacao.
        xsrf_token   : str -- Token XSRF (anti-CSRF).
        caminho_saida: str -- Caminho do arquivo de saida (.xml).

    Retorna:
        True  se a resposta foi salva com sucesso.
        False em caso de falha.
    """
    conteudo = _fetch_modelos_mart(mart_url, bearer_token, xsrf_token)
    if conteudo is None:
        return False

    try:
        pasta_saida = os.path.dirname(caminho_saida)
        if pasta_saida:
            os.makedirs(pasta_saida, exist_ok=True)
        tree = ET.ElementTree(ET.fromstring(conteudo))
        ET.indent(tree, space="  ")
        tree.write(caminho_saida, encoding="utf-8", xml_declaration=True)
        _log.info("Resposta salva em: %s", caminho_saida)
        return True
    except Exception:
        _log.exception("Erro ao salvar lista de modelos")
        return False


def filtrar_modelos_mart_por_data(
    mart_url: str,
    bearer_token: str,
    xsrf_token: str,
    data_atualizacao_min: str | None = None,
    data_atualizacao_max: str | None = None,
    data_atualizacao_exata: str | None = None,
) -> list[dict[str, str]]:
    """
    Retorna os modelos do Mart filtrando por UpdatedOn.

    Modos de filtro (inclusivos):
    - Data exata: data_atualizacao_exata
    - Data minima: data_atualizacao_min
    - Intervalo: data_atualizacao_min + data_atualizacao_max

    Retorno:
        lista de dict com chaves:
        - Catalog_Name
        - Catalog_Path
        - UpdatedOn
    """
    conteudo = _fetch_modelos_mart(mart_url, bearer_token, xsrf_token)
    if conteudo is None:
        return []

    if not any([data_atualizacao_min, data_atualizacao_max, data_atualizacao_exata]):
        raise ValueError("Informe ao menos um filtro de data: exata, min ou max")

    if data_atualizacao_exata and (data_atualizacao_min or data_atualizacao_max):
        raise ValueError("Use data exata OU intervalo (min/max), nao ambos")

    dt_exata = _parse_data_filtro(data_atualizacao_exata, "data_atualizacao_exata") if data_atualizacao_exata else None
    dt_min = _parse_data_filtro(data_atualizacao_min, "data_atualizacao_min") if data_atualizacao_min else None
    dt_max = _parse_data_filtro(data_atualizacao_max, "data_atualizacao_max") if data_atualizacao_max else None

    if dt_min and dt_max and dt_min.date() > dt_max.date():
        raise ValueError("data_atualizacao_min nao pode ser maior que data_atualizacao_max")

    root = ET.fromstring(conteudo)
    modelos_filtrados: list[dict[str, str]] = []

    for model in root.findall("Model"):
        catalog_path = model.findtext("Catalog_Path", "")
        catalog_name = model.findtext("Catalog_Name", "")
        updated_on   = model.findtext("UpdatedOn", "")

        if not catalog_path or not catalog_name or not updated_on:
            continue

        try:
            modelo_dt = datetime.strptime(updated_on, _MART_DATE_FMT)
        except ValueError:
            continue

        data_modelo = modelo_dt.date()

        if dt_exata and data_modelo != dt_exata.date():
            continue
        if dt_min and data_modelo < dt_min.date():
            continue
        if dt_max and data_modelo > dt_max.date():
            continue

        modelos_filtrados.append(
            {
                "Catalog_Name": catalog_name,
                "Catalog_Path": catalog_path,
                "UpdatedOn":    updated_on,
            }
        )

    return modelos_filtrados


def erwin_to_xml(caminho_modelo_erwin: str, caminho_saida_xml: str) -> bool:
    """
    Exporta um modelo erwin DM local para XML indentado usando ERXML.XMLERwinLink.

    Usa a interface nativa 'ERXML.XMLERwinLink' (StandAloneExport) que produz
    XML UTF-8 valido, conforme ao schema oficial do erwin DM:
        erwinSchema.xsd / EMX.xsd / EM2.xsd / EMXProps.xsd / UDP.xsd

    Assinatura COM:
        StandAloneExport(ERModelName: str, xmlFileName: str, iExpandMacro: int)

    Parametros:
        caminho_modelo_erwin : str -- Caminho absoluto do .erwin de entrada.
                                      Ex: '.\\input\\eMovies.erwin'
        caminho_saida_xml    : str -- Caminho absoluto do .xml de saida.
                                      Ex: '.\\output\\eMovies.xml'

    Retorna:
        True  se o arquivo XML foi gerado com sucesso.
        False em caso de falha.
    """
    if not os.path.isfile(caminho_modelo_erwin):
        _log.error("Arquivo de entrada nao encontrado: %s", caminho_modelo_erwin)
        return False
    return _exportar_via_com(caminho_modelo_erwin, caminho_saida_xml)


def mart_exportar_todos_xml(
    mart_url: str,
    bearer_token: str,
    xsrf_token: str,
    mart_conn_str: str,
    caminho_saida_dir: str,
    caminho_temp_dir: str,
    data_atualizacao_min: str | None = None,
    data_atualizacao_max: str | None = None,
    data_atualizacao_exata: str | None = None,
) -> dict[str, bool]:
    """
    Lista todos os modelos do Mart Server e exporta cada um para XML indentado.

    Replica a estrutura de pastas do Mart Server na saida:
        Exemplo: Mart/Modelos/eMovies -> output/xml/Mart/Modelos/eMovies.xml

    Utiliza o locator COM no formato:
        mart://Mart/<Catalog_Path_sem_prefixo>/<Catalog_Name>?<mart_conn_str>

    Parametros:
        mart_url           : str  -- URL base do Mart Server.
        bearer_token       : str  -- Token Bearer de autenticacao.
        xsrf_token         : str  -- Token XSRF anti-CSRF.
        mart_conn_str      : str  -- Params de conexao COM.
                                     Ex: 'TRC=NO;SRV=mart.empresa.com;PRT=443;
                                          ASR=MartServer;SSL=YES;UID=...;PSW=...'
        caminho_saida_dir  : str  -- Diretorio de saida dos XMLs gerados.
        caminho_temp_dir   : str  -- Diretorio para arquivos .erwin temporarios.
                                     Os arquivos sao removidos apos exportacao.
        data_atualizacao_min : str | None
                     -- Data minima (UpdatedOn >= min).
        data_atualizacao_max : str | None
                     -- Data maxima (UpdatedOn <= max).
        data_atualizacao_exata: str | None
                     -- Data exata (UpdatedOn == exata).

        Regras:
        - Exata e exclusiva (nao combinar com min/max).
        - Min e max podem ser combinadas para formar intervalo.

    Retorna:
        dict {Catalog_Name: bool} com o resultado de cada exportacao.
    """
    conteudo = _fetch_modelos_mart(mart_url, bearer_token, xsrf_token)
    if conteudo is None:
        return {}

    root = ET.fromstring(conteudo)

    modelos_base: list[dict[str, str]]
    if any([data_atualizacao_min, data_atualizacao_max, data_atualizacao_exata]):
        modelos_base = filtrar_modelos_mart_por_data(
            mart_url=mart_url,
            bearer_token=bearer_token,
            xsrf_token=xsrf_token,
            data_atualizacao_min=data_atualizacao_min,
            data_atualizacao_max=data_atualizacao_max,
            data_atualizacao_exata=data_atualizacao_exata,
        )
    else:
        modelos_base = []
        for model in root.findall("Model"):
            catalog_path = model.findtext("Catalog_Path", "")
            catalog_name = model.findtext("Catalog_Name", "")
            updated_on   = model.findtext("UpdatedOn", "")
            if catalog_path and catalog_name:
                modelos_base.append(
                    {
                        "Catalog_Name": catalog_name,
                        "Catalog_Path": catalog_path,
                        "UpdatedOn":    updated_on,
                    }
                )

    modelos = []
    for model in modelos_base:
        catalog_path = model["Catalog_Path"]
        catalog_name = model["Catalog_Name"]
        updated_on   = model.get("UpdatedOn", "")

        if any([data_atualizacao_min, data_atualizacao_max, data_atualizacao_exata]) and updated_on:
            _log.info("Selecionado por data: %s (%s)", catalog_name, updated_on)

        modelos.append((catalog_path, catalog_name))

    _log.info("%d modelo(s) a exportar", len(modelos))
    os.makedirs(caminho_saida_dir, exist_ok=True)
    os.makedirs(caminho_temp_dir,  exist_ok=True)

    resultados: dict[str, bool] = {}
    for catalog_path, catalog_name in modelos:
        locator = montar_locator_mart(catalog_path, catalog_name, mart_conn_str)
        _log.debug("Locator: %s", locator)

        # Replicar estrutura de pastas do Mart no output
        dir_saida_model = os.path.join(caminho_saida_dir, catalog_path)
        os.makedirs(dir_saida_model, exist_ok=True)

        caminho_saida = os.path.join(dir_saida_model, f"{catalog_name}.xml")
        caminho_temp  = os.path.join(caminho_temp_dir,  f"{catalog_name}.erwin")

        _log.debug("Catalog_Path=%s  Catalog_Name=%s", catalog_path, catalog_name)
        _log.debug("Saida: %s", caminho_saida)
        resultados[catalog_name] = _mart_exportar_modelo(locator, caminho_saida, caminho_temp)

    return resultados


# =============================================================================
# ENTRY POINT
# =============================================================================
if __name__ == "__main__":

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    _carregar_dotenv(os.path.join(BASE_DIR, ".env"))

    LOG_DIR    = os.path.join(BASE_DIR, "log")
    OUTPUT_DIR = os.path.join(BASE_DIR, "output")
    REPORT_DIR = os.path.join(OUTPUT_DIR, "mart_report")
    XML_DIR    = os.path.join(OUTPUT_DIR, "xml")
    TEMP       = os.path.join(BASE_DIR, "temp")

    _configurar_log(LOG_DIR)

    os.makedirs(REPORT_DIR, exist_ok=True)
    os.makedirs(XML_DIR,    exist_ok=True)
    os.makedirs(TEMP,       exist_ok=True)

    MART_URL   = _env_obrigatorio("MART_URL")
    _protocolo = os.getenv("MART_PROTOCOL", "https")
    BEARER_ENV = os.getenv("MART_BEARER_TOKEN", "").strip()
    XSRF_ENV   = os.getenv("MART_XSRF_TOKEN",   "").strip()

    if not BEARER_ENV:
        _log.info("MART_BEARER_TOKEN nao informado — gerando token automaticamente")
        _user  = _env_obrigatorio("USER_MART")
        _pass  = _env_obrigatorio("PASS_MART")
        BEARER, XSRF = _gerar_token_mart(_montar_base_url_api(MART_URL, _protocolo), _user, _pass)
    elif not XSRF_ENV:
        _log.info("MART_XSRF_TOKEN nao informado — obtendo XSRF token")
        XSRF   = _obter_xsrf_token(_montar_base_url_api(MART_URL, _protocolo))
        BEARER = BEARER_ENV
    else:
        BEARER = BEARER_ENV
        XSRF   = XSRF_ENV

    MART_CONN_BASE = _env_obrigatorio("MART_CONN_STR")
    USER_MART      = os.getenv("USER_MART", "").strip() or None
    PASS_MART      = os.getenv("PASS_MART", "").strip() or None
    MART_CONN      = _montar_conn_str_mart(MART_CONN_BASE, USER_MART, PASS_MART, MART_URL)
    DATA_EXATA     = os.getenv("MART_UPDATED_ON_EXACT", "").strip() or None
    DATA_MIN       = os.getenv("MART_UPDATED_ON_MIN",   "").strip() or None
    DATA_MAX       = os.getenv("MART_UPDATED_ON_MAX",   "").strip() or None

    # --- 1. Salvar lista de modelos ---
    listar_modelos_mart(
        mart_url=MART_URL,
        bearer_token=BEARER,
        xsrf_token=XSRF,
        caminho_saida=os.path.join(REPORT_DIR, "mart_models.xml"),
    )

    # --- 2. Exportar todos os modelos do Mart para XML ---
    resultados = mart_exportar_todos_xml(
        mart_url=MART_URL,
        bearer_token=BEARER,
        xsrf_token=XSRF,
        mart_conn_str=MART_CONN,
        caminho_saida_dir=XML_DIR,
        caminho_temp_dir=TEMP,
        data_atualizacao_min=DATA_MIN,
        data_atualizacao_max=DATA_MAX,
        data_atualizacao_exata=DATA_EXATA,
    )

    # Console: apenas resumo final
    print("\nRESUMO")
    print("-" * 40)
    for nome, ok in resultados.items():
        status = "OK   " if ok else "FALHA"
        print(f"  [{status}] {nome}")
    total_ok    = sum(1 for ok in resultados.values() if ok)
    total_falha = len(resultados) - total_ok
    print(f"\n  {total_ok} exportado(s)  |  {total_falha} com falha")
    print(f"  Log detalhado: {os.path.join(LOG_DIR, datetime.now().strftime('erwin_%Y-%m-%d.log'))}")
