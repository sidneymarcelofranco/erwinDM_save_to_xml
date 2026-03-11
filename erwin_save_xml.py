import win32com.client
import os
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


def _formatar_xml(caminho: str) -> None:
    """Le um arquivo XML, indenta e reescreve no mesmo caminho (UTF-8)."""
    tree = ET.parse(caminho)
    ET.indent(tree, space="  ")
    tree.write(caminho, encoding="utf-8", xml_declaration=True)


def _fetch_modelos_mart(mart_url: str, bearer_token: str, xsrf_token: str) -> bytes | None:
    """
    Consulta GET /MartServer/api/report/generateReport/Models e retorna os
    bytes da resposta, ou None em caso de falha.
    """
    protocolo = os.getenv("MART_PROTOCOL", "https")
    base_url = _montar_base_url_api(mart_url, protocolo)
    endpoint = f"{base_url}/MartServer/api/report/generateReport/Models"
    print(f"[INFO] Consultando Mart: {endpoint}")
    try:
        req = urllib.request.Request(endpoint)
        req.add_header("accept", "*/*")
        req.add_header("Authorization", f"Bearer {bearer_token}")
        req.add_header("X-XSRF-TOKEN", xsrf_token)
        with urllib.request.urlopen(req, context=_SSL_CTX, timeout=30) as resp:
            status = resp.status
            conteudo = resp.read()
        print(f"[INFO] HTTP {status} -- {len(conteudo):,} bytes recebidos")
        return conteudo
    except urllib.error.HTTPError as e:
        print(f"[ERRO] HTTP {e.code}: {e.reason}")
    except urllib.error.URLError as e:
        print(f"[ERRO] Falha de conexao: {e.reason}")
    except Exception as e:
        print(f"[EXCECAO] {e}")
    return None


def _exportar_via_com(locator: str, caminho_saida_xml: str) -> bool:
    """
    Chama ERXML.XMLERwinLink.StandAloneExport e pós-processa o XML gerado.

    Aceita tanto caminhos locais quanto locators Mart (Mart://Mart/...).
    """
    if not caminho_saida_xml.lower().endswith(".xml"):
        print("[AVISO] O caminho de saida nao possui extensao .xml.")
    try:
        print(f"[INFO] Exportando via {ERWIN_XML_PROGID}")
        print(f"[INFO] Entrada : {locator}")
        print(f"[INFO] Saida   : {caminho_saida_xml}")
        obj_xml = win32com.client.Dispatch(ERWIN_XML_PROGID)
        obj_xml.StandAloneExport(locator, caminho_saida_xml, 0)
        # COM nao retorna status — formata diretamente; erro de I/O indica falha
        _formatar_xml(caminho_saida_xml)
        tamanho = os.path.getsize(caminho_saida_xml)
        print(f"[OK] XML exportado: {caminho_saida_xml} ({tamanho:,} bytes)")
        return True
    except Exception as e:
        print(f"[EXCECAO] Erro na exportacao XML: {e}")
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
        print(f"[INFO] SCAPI: conectando ao Mart")
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
        print(f"[INFO] SCAPI: sessao M0={ret_m0}  M1={ret_m1}")

        print(f"[INFO] SCAPI: salvando local em {caminho_temp}")
        # Salvando em arquivo local temporario, usar disposicao de arquivo.
        pu.Save(caminho_temp, "OVF=Yes")
        print(f"[INFO] SCAPI: Save() concluido")
        scapi.Sessions.Clear()

        tamanho_temp = os.path.getsize(caminho_temp)
        print(f"[INFO] SCAPI: arquivo temp {tamanho_temp:,} bytes")
        if tamanho_temp == 0:
            print(f"[ERRO] SCAPI: Save() nao gravou dados no arquivo temp")
            os.unlink(caminho_temp)
            return False

    except Exception as e:
        print(f"[EXCECAO] SCAPI Mart: {e}")
        if os.path.exists(caminho_temp):
            os.unlink(caminho_temp)
        return False

    resultado = _exportar_via_com(caminho_temp, caminho_saida_xml)
    if os.path.exists(caminho_temp):
        os.unlink(caminho_temp)
    return resultado


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
        print(f"[OK] Resposta salva em: {caminho_saida}")
        return True
    except Exception as e:
        print(f"[EXCECAO] {e}")
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
        updated_on = model.findtext("UpdatedOn", "")

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

        if True:
            modelos_filtrados.append(
                {
                    "Catalog_Name": catalog_name,
                    "Catalog_Path": catalog_path,
                    "UpdatedOn": updated_on,
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
        print(f"[ERRO] Arquivo de entrada nao encontrado: {caminho_modelo_erwin}")
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
            updated_on = model.findtext("UpdatedOn", "")
            if catalog_path and catalog_name:
                modelos_base.append(
                    {
                        "Catalog_Name": catalog_name,
                        "Catalog_Path": catalog_path,
                        "UpdatedOn": updated_on,
                    }
                )

    modelos = []
    for model in modelos_base:
        catalog_path = model["Catalog_Path"]
        catalog_name = model["Catalog_Name"]
        updated_on = model.get("UpdatedOn", "")

        if any([data_atualizacao_min, data_atualizacao_max, data_atualizacao_exata]) and updated_on:
            print(f"[INFO] Selecionado por data: {catalog_name} ({updated_on})")

        modelos.append((catalog_path, catalog_name))

    print(f"[INFO] {len(modelos)} modelo(s) a exportar")
    os.makedirs(caminho_saida_dir, exist_ok=True)
    os.makedirs(caminho_temp_dir,  exist_ok=True)

    resultados: dict[str, bool] = {}
    for catalog_path, catalog_name in modelos:
        locator = montar_locator_mart(catalog_path, catalog_name, mart_conn_str)
        print(f"[INFO] Locator montado: {locator}")
        caminho_saida = os.path.join(caminho_saida_dir, f"{catalog_name}.xml")
        caminho_temp  = os.path.join(caminho_temp_dir,  f"{catalog_name}.erwin")
        resultados[catalog_name] = _mart_exportar_modelo(locator, caminho_saida, caminho_temp)

    return resultados


# =============================================================================
# EXEMPLO DE USO
# =============================================================================
if __name__ == "__main__":

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    _carregar_dotenv(os.path.join(BASE_DIR, ".env"))

    OUTPUT_DIR = os.path.join(BASE_DIR, "output")
    REPORT_DIR = os.path.join(OUTPUT_DIR, "mart_report")
    XML_DIR    = os.path.join(OUTPUT_DIR, "xml")
    TEMP       = os.path.join(BASE_DIR, "temp")
    os.makedirs(REPORT_DIR, exist_ok=True)
    os.makedirs(XML_DIR,    exist_ok=True)
    os.makedirs(TEMP,       exist_ok=True)

    MART_URL     = _env_obrigatorio("MART_URL")
    BEARER       = _env_obrigatorio("MART_BEARER_TOKEN")
    XSRF         = _env_obrigatorio("MART_XSRF_TOKEN")
    MART_CONN_BASE = _env_obrigatorio("MART_CONN_STR")
    USER_MART      = os.getenv("USER_MART", "").strip() or None
    PASS_MART      = os.getenv("PASS_MART", "").strip() or None
    MART_CONN      = _montar_conn_str_mart(MART_CONN_BASE, USER_MART, PASS_MART, MART_URL)
    DATA_EXATA   = os.getenv("MART_UPDATED_ON_EXACT", "").strip() or None
    DATA_MIN_ENV = os.getenv("MART_UPDATED_ON_MIN", "").strip()
    DATA_MIN     = DATA_MIN_ENV or None
    DATA_MAX_ENV = os.getenv("MART_UPDATED_ON_MAX", "").strip()
    DATA_MAX     = DATA_MAX_ENV or None

    # --- 1. Salvar lista de modelos ---
    listar_modelos_mart(
        mart_url=MART_URL,
        bearer_token=BEARER,
        xsrf_token=XSRF,
        caminho_saida=os.path.join(REPORT_DIR, "mart_models.xml"),
    )

    # --- 2. Exportar todos os modelos do Mart para XML ---
    #        Filtra apenas modelos atualizados a partir de 02/10/2026
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

    print("\n[RESUMO]")
    for nome, ok in resultados.items():
        status = "OK" if ok else "FALHA"
        print(f"  [{status}] {nome}")
