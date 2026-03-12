"""
app/app.py
----------
Interface Streamlit para erwin_save_xml.py.

Execute a partir da raiz do projeto:
    streamlit run app/app.py
"""

import base64
import sys
import subprocess
import threading
import time
from datetime import datetime
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

BASE_DIR   = Path(__file__).resolve().parent.parent
ENV_FILE   = BASE_DIR / ".env"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR    = BASE_DIR / "log"
SCRIPT     = BASE_DIR / "erwin_save_xml.py"

# ---------------------------------------------------------------------------
# Estado global persistente entre reruns via sys.modules
# (globals() não funciona no modelo de execução do Streamlit)
# ---------------------------------------------------------------------------

_EXEC_KEY = "__erwin_app_exec_state__"
if _EXEC_KEY not in sys.modules:
    sys.modules[_EXEC_KEY] = {          # type: ignore[assignment]
        "proc":       None,             # subprocess.Popen em andamento
        "finished":   False,
        "returncode": None,
        "log_file":   None,             # Path do arquivo de log desta execução
        "log_offset": 0,               # bytes já lidos antes de iniciar
    }

_G: dict = sys.modules[_EXEC_KEY]      # type: ignore[assignment]


def _log_file_hoje() -> Path:
    nome = datetime.now().strftime("erwin_%Y-%m-%d.log")
    return LOG_DIR / nome


def _ler_log_novo() -> str:
    """Lê apenas o conteúdo novo do arquivo de log desde o início da execução."""
    log_path: Path | None = _G["log_file"]
    if log_path is None or not log_path.exists():
        return ""
    try:
        with open(log_path, "rb") as f:
            f.seek(_G["log_offset"])
            dados = f.read()
        return dados.decode("utf-8", errors="replace")
    except OSError:
        return ""


def _limpar_logs_em_disco() -> tuple[int, int]:
    """Limpa (trunca) todos os arquivos .log em LOG_DIR.

    Retorna (quantidade_limpos, quantidade_erros).
    """
    if not LOG_DIR.exists():
        return (0, 0)

    limpos = 0
    erros = 0
    for log_file in LOG_DIR.glob("*.log"):
        try:
            log_file.write_text("", encoding="utf-8")
            limpos += 1
        except OSError:
            erros += 1
    return (limpos, erros)


def _run_script_thread() -> None:
    """Executa o script em background thread."""
    proc = subprocess.Popen(
        [sys.executable, str(SCRIPT)],
        cwd=str(BASE_DIR),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    _G["proc"] = proc
    proc.wait()
    _G["returncode"] = proc.returncode
    _G["finished"]   = True
    _G["proc"]       = None


# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="erwin Mart → XML",
    page_icon="🗂️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Session state defaults
# ---------------------------------------------------------------------------

if "logs_exec"       not in st.session_state:
    st.session_state.logs_exec       = ""
if "exec_status"     not in st.session_state:
    st.session_state.exec_status     = None
if "auto_switch_tab" not in st.session_state:
    st.session_state.auto_switch_tab = False
if "running"         not in st.session_state:
    st.session_state.running         = False

# ---------------------------------------------------------------------------
# CSS: overlay de loading e bloqueio da sidebar durante execução
# ---------------------------------------------------------------------------

if st.session_state.running:
    st.markdown("""
    <style>
    /* Bloqueia a sidebar durante execução */
    section[data-testid="stSidebar"] {
        pointer-events: none !important;
        opacity: 0.45 !important;
        transition: opacity 0.3s;
    }
    /* Overlay semi-transparente sobre o conteúdo principal */
    .main .block-container::before {
        content: "";
        position: fixed;
        inset: 0;
        background: rgba(14, 17, 23, 0.55);
        backdrop-filter: blur(2px);
        z-index: 999;
        pointer-events: none;
    }
    </style>
    """, unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# .env helpers
# ---------------------------------------------------------------------------

ENV_FIELDS: list[tuple[str, str, str, bool]] = [
    # (chave, label, placeholder, is_password)
    ("MART_URL",              "Mart Server URL",           "mart.empresa.com.br",                  False),
    ("USER_MART",             "Usuário",                   "Administrator",                         False),
    ("PASS_MART",             "Senha",                     "",                                      True),
    ("MART_CONN_STR",         "Connection String COM",     "TRC=NO;PRT=443;ASR=MartServer;SSL=YES", False),
    ("MART_PROTOCOL",         "Protocolo",                 "https",                                 False),
    ("MART_BEARER_TOKEN",     "Bearer Token (opcional)",   "",                                      True),
    ("MART_XSRF_TOKEN",       "XSRF Token (opcional)",     "",                                      True),
    ("MART_UPDATED_ON_EXACT", "Data Exata  (MM/DD/YYYY)",  "",                                      False),
    ("MART_UPDATED_ON_MIN",   "Data Mínima (MM/DD/YYYY)",  "",                                      False),
    ("MART_UPDATED_ON_MAX",   "Data Máxima (MM/DD/YYYY)",  "",                                      False),
]


def ler_env() -> dict[str, str]:
    if not ENV_FILE.is_file():
        return {}
    resultado: dict[str, str] = {}
    for linha in ENV_FILE.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        k, v = linha.split("=", 1)
        resultado[k.strip()] = v.strip()
    return resultado


def salvar_env(valores: dict[str, str]) -> None:
    linhas_orig = (
        ENV_FILE.read_text(encoding="utf-8").splitlines(keepends=True)
        if ENV_FILE.is_file() else []
    )
    atualizadas: set[str] = set()
    novas: list[str] = []
    for linha in linhas_orig:
        stripped = linha.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            novas.append(linha)
            continue
        k = stripped.split("=", 1)[0].strip()
        if k in valores:
            novas.append(f"{k}={valores[k]}\n")
            atualizadas.add(k)
        else:
            novas.append(linha)
    for k, v in valores.items():
        if k not in atualizadas:
            novas.append(f"{k}={v}\n")
    ENV_FILE.write_text("".join(novas), encoding="utf-8")


# ---------------------------------------------------------------------------
# Sidebar — configuração + botões
# ---------------------------------------------------------------------------

env_atual    = ler_env()
btn_salvar   = False
btn_executar = False
form_valores: dict[str, str] = {}

with st.sidebar:
    st.title("⚙️ Configuração")
    st.caption("Parâmetros gravados em `.env`")

    with st.form("form_env"):
        for chave, label, placeholder, is_pass in ENV_FIELDS:
            valor = env_atual.get(chave, "")
            form_valores[chave] = st.text_input(
                label=label,
                value=valor,
                placeholder=placeholder,
                type="password" if is_pass else "default",
                key=f"field_{chave}",
            )

        st.markdown("---")
        col_s, col_r = st.columns(2)
        btn_salvar = col_s.form_submit_button(
            "💾 Salvar",
            use_container_width=True,
            disabled=st.session_state.running,
        )
        btn_executar = col_r.form_submit_button(
            "⏳ Executando..." if st.session_state.running else "▶ Executar",
            use_container_width=True,
            type="primary",
            disabled=st.session_state.running,
        )

    if btn_salvar or btn_executar:
        salvar_env(form_valores)
        if btn_salvar and not btn_executar:
            st.success("✅ .env salvo")

    st.markdown("---")
    st.caption(f"Script: `{SCRIPT.name}`")
    st.caption(f"Output: `{OUTPUT_DIR.relative_to(BASE_DIR)}`")

# ---------------------------------------------------------------------------
# Lógica de execução em background thread
# ---------------------------------------------------------------------------

if btn_executar and not st.session_state.running:
    log_path = _log_file_hoje()
    _G["finished"]   = False
    _G["returncode"] = None
    _G["proc"]       = None
    _G["log_file"]   = log_path
    # Registra offset atual do arquivo (para exibir só o novo conteúdo)
    _G["log_offset"] = log_path.stat().st_size if log_path.exists() else 0

    st.session_state.logs_exec   = ""
    st.session_state.exec_status = None
    st.session_state.running     = True

    threading.Thread(target=_run_script_thread, daemon=True).start()
    st.rerun()

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------

tab_exec, tab_saida = st.tabs(["📋 Execução", "📁 Estrutura"])

# Após execução bem-sucedida, muda automaticamente para aba Estrutura
if st.session_state.auto_switch_tab:
    st.session_state.auto_switch_tab = False
    components.html("""
    <script>
      setTimeout(function () {
        var tabs = window.parent.document.querySelectorAll('button[data-baseweb="tab"]');
        if (tabs && tabs.length >= 2) { tabs[1].click(); }
      }, 200);
    </script>
    """, height=0)

# ---------------------------------------------------------------------------
# Tab: Execução
# ---------------------------------------------------------------------------

with tab_exec:
    # Cabeçalho com botão Limpar Log
    col_titulo, col_limpar = st.columns([8, 1])
    col_titulo.subheader("Console")
    existe_log_em_disco = LOG_DIR.exists() and any(LOG_DIR.glob("*.log"))
    if not st.session_state.running and (st.session_state.logs_exec or existe_log_em_disco):
        if col_limpar.button("🗑️ Limpar", use_container_width=True, help="Limpa o log exibido"):
            limpos, erros = _limpar_logs_em_disco()
            st.session_state.logs_exec   = ""
            st.session_state.exec_status = None
            if limpos > 0 and erros == 0:
                st.success(f"{limpos} arquivo(s) .log limpo(s) em /log.")
            elif limpos > 0 and erros > 0:
                st.warning(f"{limpos} arquivo(s) limpo(s), {erros} com erro.")
            elif erros > 0:
                st.error(f"Não foi possível limpar {erros} arquivo(s) .log.")

    # ---- Estado: executando ------------------------------------------------
    if st.session_state.running:
        col_spinner, col_stop = st.columns([5, 1])
        with col_spinner:
            st.status("⏳ Executando script… acompanhe o log abaixo.", state="running")
        with col_stop:
            if st.button("⏹ Parar", type="secondary", use_container_width=True,
                         help="Interrompe a execução imediatamente"):
                proc = _G.get("proc")
                if proc is not None:
                    proc.kill()
                log_novo = _ler_log_novo()
                _G["finished"]   = True
                _G["returncode"] = -1
                st.session_state.running     = False
                st.session_state.exec_status = -1
                st.session_state.logs_exec   = log_novo + "\n\n⛔ Execução interrompida pelo usuário."
                st.rerun()

        # Log em tempo real lido direto do arquivo
        log_novo = _ler_log_novo()
        if log_novo.strip():
            st.code(log_novo, language="text")
        else:
            st.info("Aguardando saída do script…")

        # Scroll automático para o final
        components.html("""
        <script>
          setTimeout(function () {
            var codes = window.parent.document.querySelectorAll('pre');
            if (codes.length) { codes[codes.length - 1].scrollTop = 9999999; }
          }, 150);
        </script>
        """, height=0)

        # Verifica conclusão
        if _G["finished"]:
            st.session_state.logs_exec   = _ler_log_novo()
            st.session_state.exec_status = _G["returncode"]
            st.session_state.running     = False
            if _G["returncode"] == 0:
                st.session_state.auto_switch_tab = True
            st.rerun()
        else:
            time.sleep(0.5)
            st.rerun()

    # ---- Estado: ocioso / finalizado ---------------------------------------
    else:
        if st.session_state.exec_status is not None:
            if st.session_state.exec_status == 0:
                st.success("Exportação concluída com sucesso.")
            elif st.session_state.exec_status == -1:
                st.warning("Execução interrompida pelo usuário.")
            else:
                st.error(f"Finalizado com código de saída {st.session_state.exec_status}.")

        if st.session_state.logs_exec:
            st.code(st.session_state.logs_exec, language="text")
        else:
            st.info("Preencha a configuração e clique em **▶ Executar**.")

        # Último arquivo de log detalhado
        if LOG_DIR.exists():
            logs_arquivo = sorted(LOG_DIR.glob("erwin_*.log"), reverse=True)
            if logs_arquivo:
                with st.expander(f"📄 Log detalhado: {logs_arquivo[0].name}"):
                    st.code(
                        logs_arquivo[0].read_text(encoding="utf-8", errors="replace"),
                        language="text",
                    )

# ---------------------------------------------------------------------------
# Tab: Estrutura
# ---------------------------------------------------------------------------

def listar_recursivo(path: Path, nivel: int = 0) -> list[tuple[int, Path]]:
    """Retorna lista plana de (nivel, path) para toda a arvore."""
    items: list[tuple[int, Path]] = []
    try:
        entries = sorted(
            path.iterdir(),
            key=lambda p: (p.is_file(), p.name.lower()),
        )
    except PermissionError:
        return items
    for entry in entries:
        items.append((nivel, entry))
        if entry.is_dir():
            items.extend(listar_recursivo(entry, nivel + 1))
    return items


def abrir_no_navegador(caminho: Path) -> None:
    """Abre o arquivo em nova aba via Blob URL para evitar pagina em branco."""
    conteudo = caminho.read_text(encoding="utf-8", errors="replace")
    b64 = base64.b64encode(conteudo.encode("utf-8")).decode()
    mime = "application/xml" if caminho.suffix == ".xml" else "text/plain"
    components.html(f"""
    <script>
            (function () {{
                const b64 = "{b64}";
                const bin = atob(b64);
                const bytes = new Uint8Array(bin.length);
                for (let i = 0; i < bin.length; i++) {{
                    bytes[i] = bin.charCodeAt(i);
                }}

                const blob = new Blob([bytes], {{ type: "{mime};charset=utf-8" }});
                const url = URL.createObjectURL(blob);
                const aba = window.open(url, "_blank");

                if (!aba) {{
                    const a = document.createElement("a");
                    a.href = url;
                    a.target = "_blank";
                    a.rel = "noopener noreferrer";
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                }}

                setTimeout(function () {{ URL.revokeObjectURL(url); }}, 60000);
            }})();
    </script>
    """, height=0)


with tab_saida:
    st.subheader("📂 Estrutura de saída")

    if not OUTPUT_DIR.exists() or not any(OUTPUT_DIR.iterdir()):
        st.info("Execute o script para gerar arquivos.")
    else:
        for nivel, entry in listar_recursivo(OUTPUT_DIR):
            indent = "\u00a0" * (nivel * 6)   # non-breaking spaces para indentação

            if entry.is_dir():
                st.markdown(
                    f"<span style='color:#888;font-size:0.9em'>"
                    f"{indent}📂 <b>{entry.name}</b></span>",
                    unsafe_allow_html=True,
                )
            else:
                icon  = "📋" if entry.suffix == ".xml" else "📄"
                tamanho_kb = entry.stat().st_size / 1024

                c_label, c_btn = st.columns([8, 1])
                c_label.markdown(
                    f"<span style='font-size:0.9em'>"
                    f"{indent}{icon} {entry.name} "
                    f"<span style='color:#aaa'>({tamanho_kb:.1f} KB)</span>"
                    f"</span>",
                    unsafe_allow_html=True,
                )
                if c_btn.button("🔗", key=f"open_{entry}", help=f"Abrir {entry.name} em nova aba"):
                    abrir_no_navegador(entry)
