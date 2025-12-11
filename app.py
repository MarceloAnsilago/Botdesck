import os
import time
import json
import base64
import random
import threading
import queue
import re
from copy import deepcopy
from io import BytesIO
from typing import Optional, Tuple

import pandas as pd
from flask import Flask, render_template, request, jsonify, Response, send_file

# ===== Selenium (WhatsApp Web) =====
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager


app = Flask(__name__)
SELECTORS_FILE = os.path.join(app.root_path, "selectors.json")
DEFAULT_SELECTORS = {
    # Editor de mensagem (contenteditable no footer)
    "chat_input": {"by": "CSS_SELECTOR", "value": "footer [contenteditable='true'][data-tab]"},
    # Botão de enviar (por data-icon/aria-label)
    "send_button": {"by": "CSS_SELECTOR", "value": "footer button[data-icon='send'], footer button[aria-label*='Enviar'], footer button[aria-label*='Send']"},
    "sidebar": {"by": "ID", "value": "side"},
}
ALLOWED_BY = {"XPATH", "CSS_SELECTOR", "ID", "CLASS_NAME", "NAME", "TAG_NAME"}
FALLBACK_SEND_SELECTORS = [
    # CSS direto pelo ícone
    (By.CSS_SELECTOR, "footer button[data-icon='send']"),
    (By.CSS_SELECTOR, "footer span[data-icon='send']"),
    (By.CSS_SELECTOR, "button[aria-label*='Enviar']"),
    (By.CSS_SELECTOR, "button[aria-label*='Send']"),
    (By.CSS_SELECTOR, "span[data-icon='send']"),
    (By.CSS_SELECTOR, "button[data-icon='send']"),
    # Botao direto que voce enviou
    (By.XPATH, "//*[@id='main']/footer/div[1]/div/span/div/div/div/div[4]/div/span/button"),
    # Retangulo do botao (versoes antigas)
    (By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div/div/div[4]/div/span/button/div/div'),
    # Interior do mesmo botao (caso o wrapper mude)
    (By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div/div/div[4]/div/span/button/div/div/div[1]/span'),
    # Alvos genericos por aria-label/data-icon
    (By.XPATH, "//button[@aria-label='Enviar' or @aria-label='Send']"),
    (By.XPATH, "//footer//button[contains(@aria-label,'Enviar') or contains(@aria-label,'Send')]"),
    (By.XPATH, "//footer//*[contains(@data-icon,'send')]"),
]
# Fallbacks para o input de chat
FALLBACK_CHAT_INPUT_SELECTORS = [
    (By.CSS_SELECTOR, "div[contenteditable='true'][data-tab]"),
    (By.CSS_SELECTOR, "footer [contenteditable='true']"),
    (By.XPATH, "//*[@id='main']//div[@contenteditable='true' and @data-tab]"),
]

def load_selectors(force_default: bool = False):
    base = deepcopy(DEFAULT_SELECTORS)
    if force_default or not os.path.exists(SELECTORS_FILE):
        return base

    try:
        with open(SELECTORS_FILE, "r", encoding="utf-8") as f:
            stored = json.load(f)
    except Exception:
        return base

    merged = {}
    for key, default in base.items():
        raw = stored.get(key, {})
        by = raw.get("by", default["by"]).upper()
        value = raw.get("value", default["value"])
        if by not in ALLOWED_BY:
            by = default["by"]
        merged[key] = {
            "by": by,
            "value": value if value else default["value"],
        }
    return merged

def persist_selectors():
    try:
        with open(SELECTORS_FILE, "w", encoding="utf-8") as f:
            json.dump(SELECTORS, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log(f"[WARN] Não foi possível salvar seletores: {e}")

def resolve_selector(key: str) -> Tuple[str, str]:
    selector = SELECTORS.get(key) or DEFAULT_SELECTORS.get(key)
    if selector is None:
        raise KeyError(f"Seletor desconhecido: {key}")
    by_name = selector.get("by", "XPATH").upper()
    value = selector.get("value", "")
    if by_name not in ALLOWED_BY:
        by_name = "XPATH"
    return getattr(By, by_name), value

SELECTORS = load_selectors()

# ===== Estado global simples =====
STATE = {
    "df": None,                # pandas.DataFrame
    "driver": None,            # selenium webdriver
    "logs": queue.Queue(),     # fila de logs para SSE
    "sending_thread": None,    # thread de envio
    "stop_flag": False,        # para parar envio se necessário
}


# ============ Helpers ============
def log(msg: str):
    """Enfileira mensagens para o front (SSE)."""
    try:
        STATE["logs"].put_nowait(msg)
    except Exception:
        pass


def df_to_visible_html(df: Optional[pd.DataFrame]) -> str:
    if df is not None and not df.empty:
        df_visivel = df.copy()
        if "Dec. Rebanho" in df_visivel.columns:
            df_visivel = df_visivel.drop(columns=["Dec. Rebanho"])
        for col in df_visivel.columns:
            df_visivel[col] = df_visivel[col].astype(str)
        return df_visivel.to_html(index=False, escape=False)
    return "<p>Nenhum dado disponível.</p>"


def number_looks_invalid(driver: webdriver.Chrome) -> bool:
    """Detecta mensagens típicas de número inválido no WhatsApp."""
    text = driver.page_source.lower()
    phrases = [
        "número inválido",
        "invalid phone number",
        "não é possível enviar mensagens para este número",
        "não é possível enviar mensagens",
        "number invalid"
    ]
    return any(phrase in text for phrase in phrases)


def wait_for_message_sent(driver: webdriver.Chrome, previous: int, timeout: int = 5) -> bool:
    """Garante que um novo bubble de saída apareceu após clicar em enviar."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: len(d.find_elements(By.XPATH, "//div[contains(@class,'message-out')]")) > previous
        )
        return True
    except TimeoutException:
        return False


def preprocess_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    colunas = [
        'Nome do Titular da Ficha de bovideos',
        'Nome da Propriedade',
        'Endereço da Prop.',
        'Dec. Rebanho',
        'Telefone 1',
        'Telefone 2',
        'Celular'
    ]
    for col in colunas:
        if col not in df.columns:
            raise ValueError(f"Coluna ausente: {col}")

    df = df[colunas]
    df = pd.melt(
        df,
        id_vars=colunas[:4],
        value_vars=colunas[4:],
        value_name='Telefone'
    ).drop(columns=['variable'])

    df['Nome'] = df.apply(
        lambda row: f"{row[colunas[0]]} - {row[colunas[1]]} - {row[colunas[2]]}",
        axis=1
    )
    df = df.drop(columns=colunas[:3])

    # Normalização de telefone
    df["Telefone"] = df["Telefone"].apply(normalize_phone_number)
    df = df[df["Telefone"].notna() & (df["Telefone"] != "")]
    df["Status"] = "Fila de envio"
    return df[["Status", "Nome", "Telefone", "Dec. Rebanho"]]


def _normalize_col_key(name: str) -> str:
    """Remove caracteres nÇ£o alfanumÇ¸ricos para facilitar matching de cabeÇ§alhos."""
    return re.sub(r"[^a-z0-9]+", "", str(name).lower())


def normalize_phone_number(val: str) -> Optional[str]:
    """Normaliza telefone BR e descarta padrões claramente inválidos."""
    digits = re.sub(r"[^0-9]", "", str(val or ""))
    if not digits:
        return None

    if digits.startswith("55") and len(digits) >= 12:
        digits = digits[2:]
    if len(digits) > 11:
        digits = digits[-11:]
    if len(digits) not in (10, 11):
        return None
    if set(digits) == {"0"}:
        return None
    if digits.startswith("00"):
        return None
    if len(digits) == 10 and re.match(r"^\d{6}0000$", digits):
        return None
    if len(digits) == 11 and re.match(r"^\d{7}0000$", digits):
        return None

    phone = "+55" + digits
    phone = phone[:3] + " " + phone[3:5] + " " + phone[5:]
    return phone
def preprocess_outros_animais_inadimplentes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Trata planilhas HTML de outros animais (ex.: lista suinos.html) filtrando inadimplentes.
    Espera colunas como Declarou, Nome do Titular da Explora‡Æo, Endere‡o, Munic¡pio + Cidade/Distrito e telefones.
    """
    col_candidates = {
        "declarou": ["declarou", "declarao", "declaracao"],
        "nome": [
            "nomedotitulardaexploracao",
            "nomedotitular",
            "nomedotitulardaexplorao",
        ],
        "endereco": ["endereco", "enderecodaexploracao", "endereo"],
        "municipio": [
            "municipiocidadedistrito",
            "municpiocidadedistrito",
            "municipio",
            "municipioecidade",
            "cidade",
        ],
        "telefone1": ["telefone1", "telefoneprimario"],
        "telefone2": ["telefone2", "telefonesecundario"],
        "celular": ["celular", "telefonecelular"],
    }

    def find_col(target_key: str) -> Optional[str]:
        desired = {key for key in col_candidates[target_key]}
        for col in df.columns:
            if _normalize_col_key(col) in desired:
                return col
        return None

    col_declarou = find_col("declarou")
    col_nome = find_col("nome")
    col_endereco = find_col("endereco")
    col_municipio = find_col("municipio")
    col_tel1 = find_col("telefone1")
    col_tel2 = find_col("telefone2")
    col_cel = find_col("celular")

    required = [col_declarou, col_nome, col_endereco, col_municipio]
    if any(col is None for col in required):
        raise ValueError("Colunas obrigatÇ®rias da lista de outros animais nÇ£o foram encontradas.")

    telefone_cols = [c for c in [col_tel1, col_tel2, col_cel] if c is not None]
    if not telefone_cols:
        raise ValueError("NÇ§o hÇ¸ colunas de telefone reconhecidas.")

    df_local = df[[col_declarou, col_nome, col_endereco, col_municipio] + telefone_cols].copy()
    dec_numeric = pd.to_numeric(df_local[col_declarou], errors="coerce")
    df_local = df_local[dec_numeric == -1]
    if df_local.empty:
        raise ValueError("Nenhum registro inadimplente encontrado para tratar.")

    df_local = pd.melt(
        df_local,
        id_vars=[col_declarou, col_nome, col_endereco, col_municipio],
        value_vars=telefone_cols,
        value_name="Telefone",
    ).drop(columns=["variable"])

    df_local["Telefone"] = df_local["Telefone"].apply(normalize_phone_number)
    df_local = df_local[df_local["Telefone"].notna() & (df_local["Telefone"] != "")]

    df_local["Nome"] = df_local.apply(
        lambda row: f"{row[col_nome]} - {row[col_endereco]} - {row[col_municipio]}",
        axis=1,
    )
    df_local["Dec. Rebanho"] = -1
    df_local["Status"] = "Fila de envio"
    return df_local[["Status", "Nome", "Telefone", "Dec. Rebanho"]]


def get_stats(df: Optional[pd.DataFrame]):
    try:
        if df is None or df.empty:
            return {
                "total": 0,
                "fila": {"qtd": 0, "perc": "0%"},
                "enviados": {"qtd": 0, "perc": "0%"},
                "invalidos": {"qtd": 0, "perc": "0%"},
            }

        total = len(df)
        if total == 0:
            return {
                "total": 0,
                "fila": {"qtd": 0, "perc": "0%"},
                "enviados": {"qtd": 0, "perc": "0%"},
                "invalidos": {"qtd": 0, "perc": "0%"},
            }

        if "Status" in df.columns:
            status_series = df["Status"].fillna("").astype(str)
        else:
            status_series = pd.Series([""] * total)

        fila = int((status_series == "Fila de envio").sum())
        enviados = int((status_series == "Enviado").sum())
        invalidos = int((status_series == "Número inválido").sum())

        def perc(qtd: int) -> str:
            return f"{(qtd / total * 100):.1f}%" if total > 0 else "0%"

        return {
            "total": total,
            "fila": {"qtd": fila, "perc": perc(fila)},
            "enviados": {"qtd": enviados, "perc": perc(enviados)},
            "invalidos": {"qtd": invalidos, "perc": perc(invalidos)},
        }
    except Exception as e:
        return {"erro": str(e)}


def start_driver() -> webdriver.Chrome:
    path = ChromeDriverManager().install()
    options = Options()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    # Em desktops/VM, pode ser útil manter visível (sem headless):
    driver = webdriver.Chrome(service=Service(path), options=options)
    driver.get('https://web.whatsapp.com')
    return driver


from selenium.webdriver.common.keys import Keys

def enviar(driver: webdriver.Chrome, numero: str, mensagem: str) -> bool:
    try:
        numero = numero.replace(" ", "")
        from urllib.parse import quote
        mensagem_q = quote(mensagem)

        # Abre a conversa com o texto pre-preenchido
        driver.get(f"https://web.whatsapp.com/send?phone={numero}&text={mensagem_q}")

        # Conta quantas mensagens "message-out" existem antes do envio
        previous = len(driver.find_elements(By.XPATH, "//div[contains(@class,'message-out')]"))

        # Espera o editor aparecer e garante foco (com fallbacks)
        input_by, input_value = resolve_selector("chat_input")
        chat_candidates = [(input_by, input_value)] + FALLBACK_CHAT_INPUT_SELECTORS
        chat_input = None
        for by_opt, value_opt in chat_candidates:
            try:
                chat_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((by_opt, value_opt))
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", chat_input)
                chat_input.click()
                break
            except Exception:
                chat_input = None
        if chat_input is None:
            log("[WARN] NÇåo encontrei o editor de mensagem.")
            return False

        # Checa numero invalido antes de tentar enviar
        if number_looks_invalid(driver):
            log(f"[WARN] Numero invalido detectado para {numero}.")
            return False

        # Aguarda botao enviar ficar clicavel, com fallback de seletores
        send_by, send_value = resolve_selector("send_button")
        send_candidates = [(send_by, send_value)]
        for alt in FALLBACK_SEND_SELECTORS:
            if alt not in send_candidates:
                send_candidates.append(alt)

        last_error = None
        send_btn = None
        for by_opt, value_opt in send_candidates:
            try:
                send_btn = WebDriverWait(driver, 8).until(
                    EC.element_to_be_clickable((by_opt, value_opt))
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", send_btn)
                send_btn.click()
                break
            except Exception as click_err:
                last_error = click_err
                send_btn = None

        if send_btn is None:
            # Tenta clique via JS com seletores mais amplos
            js_selectors = [
                "button[aria-label*='Enviar']",
                "button[aria-label*='Send']",
                "span[data-icon='send']",
                "button[data-icon='send']",
                "footer button[aria-label]"
            ]
            clicked_js = False
            for sel in js_selectors:
                try:
                    clicked_js = driver.execute_script(
                        "const el = document.querySelector(arguments[0]); if(el){el.click(); return true;} return false;",
                        sel,
                    )
                    if clicked_js:
                        break
                except Exception as js_err:
                    last_error = js_err

            if not clicked_js:
                # Plano B: enter no editor
                try:
                    chat_input.send_keys(Keys.ENTER)
                except Exception:
                    log(f"[WARN] Falha ao clicar no botao (ultimo erro: {last_error}) e ao enviar com ENTER ({numero}).")
                    return False

        # Confirma que apareceu uma nova mensagem enviada
        if not wait_for_message_sent(driver, previous, timeout=10):
            log(f"[WARN] Nao consegui confirmar envio para {numero}.")
            return False

        time.sleep(0.8)
        return True

    except Exception as e:
        print(f"[ERRO] Falha ao enviar para {numero}: {e}")
        log(f"[ERRO] Falha ao enviar para {numero}: {e}")
        return False


def run_sending(
    tempo_min: int,
    tempo_max: int,
    mensagem_template: str,
    mensagens_min: int,
    mensagens_max: int,
    minutos_min: int,
    minutos_max: int,
):
    log("[DEBUG] Iniciando envio de mensagens...")
    df = STATE["df"]
    if df is None or df.empty:
        log("Nenhum dado disponível para envio.")
        return

    driver = STATE["driver"]
    if driver is None:
        log("⚠️ Driver/WhatsApp não inicializado. Clique em 'Iniciar WhatsApp Web' antes.")
        return

    try:
        # Espera até o WhatsApp estar pronto (área lateral carregada)
        sidebar_by, sidebar_value = resolve_selector("sidebar")
        while len(driver.find_elements(sidebar_by, sidebar_value)) < 1:
            time.sleep(1)

        log("WhatsApp Web está pronto para envio.")

        batch_min = max(1, mensagens_min)
        batch_max = max(batch_min, mensagens_max)
        pause_min = max(0, minutos_min)
        pause_max = max(pause_min, minutos_max)
        batch_target = random.randint(batch_min, batch_max)
        batch_counter = 0

        # Itera somente quem está na fila
        fila_indices = list(df[df['Status'] == 'Fila de envio'].index)
        for posicao, idx in enumerate(fila_indices):
            row = df.loc[idx]
            if STATE["stop_flag"]:
                log("Envio interrompido pelo usuário.")
                break

            numero = row['Telefone']
            nome = row['Nome']
            mensagem = (
                mensagem_template
                .replace("-&numero", numero)
                .replace("-&nome", nome)
            )

            log(f"Enviando para {nome} ({numero})")
            sucesso = enviar(driver, numero, mensagem)

            if sucesso:
                df.at[idx, 'Status'] = 'Enviado'
                log(f"✅ Mensagem enviada para {nome}")

                delay = random.randint(tempo_min, tempo_max)
                for i in range(delay, 0, -1):
                    log(f"Próximo envio em {i} segundos...")
                    time.sleep(1)
                log("⏭️ Próximo contato...")

            else:
                df.at[idx, 'Status'] = 'Número inválido'
                log(f"⚠️ Falha ao enviar para {nome} ({numero}) — número inválido.")
                log("⏭️ Pulando para o próximo contato...")

            batch_counter += 1
            if (
                batch_counter >= batch_target
                and pause_max > 0
                and posicao != len(fila_indices) - 1
            ):
                pause_minutes = random.randint(pause_min, pause_max)
                if pause_minutes > 0:
                    pause_seconds = pause_minutes * 60
                    log(f"⏸️ Pausando por {pause_minutes} minutos.")
                    for remaining in range(pause_seconds, 0, -1):
                        log(f"Pausa ativa - voltando em {remaining} segundos...")
                        time.sleep(1)
                    log("▶️ Retomando envios...")
                batch_counter = 0
                batch_target = random.randint(batch_min, batch_max)

            # Salva progresso sempre
            try:
                out = BytesIO()
                df.to_excel(out, index=False)
                out.seek(0)
                with open("BancoProd.xlsx", "wb") as f:
                    f.write(out.read())
            except Exception as e:
                log(f"[ERRO] Falha ao salvar BancoProd.xlsx: {e}")

        STATE["df"] = df

    except Exception as e:
        log(f"[ERRO] Envio interrompido: {e}")


# ============ Rotas ============
@app.get("/selectors")
def get_selectors():
    return jsonify(SELECTORS)


@app.post("/selectors")
def update_selectors():
    data = request.get_json()
    if not isinstance(data, dict):
        return jsonify({"ok": False, "msg": "Formato inválido."}), 400

    updated = {}
    for key, default in DEFAULT_SELECTORS.items():
        entry = data.get(key)
        if not isinstance(entry, dict):
            continue
        by = (entry.get("by") or default["by"]).upper()
        value = str(entry.get("value", "")).strip()
        if not value:
            return jsonify({"ok": False, "msg": f"Valor requerido para {key}."}), 400
        if by not in ALLOWED_BY:
            return jsonify({"ok": False, "msg": f"Tipo de seletor inválido para {key}."}), 400
        updated[key] = {"by": by, "value": value}

    if not updated:
        return jsonify({"ok": False, "msg": "Nenhum seletor informado."}), 400

    SELECTORS.update(updated)
    persist_selectors()
    return jsonify({"ok": True, "msg": "Seletores atualizados com sucesso."})


@app.post("/selectors/reset")
def reset_selectors():
    global SELECTORS
    SELECTORS = load_selectors(force_default=True)
    persist_selectors()
    return jsonify({"ok": True, "msg": "Seletores restaurados para os padrões."})


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/upload")
def upload():
    """
    Suporta upload de arquivo direto (form-data: file),
    ou base64 em JSON: { "nome_arquivo": "...", "conteudo_base64": "data:...base64,..." }
    """
    try:
        # 1) Via form-data
        if "file" in request.files:
            f = request.files["file"]
            name = f.filename.lower()
            buf = BytesIO(f.read())
        else:
            # 2) Via JSON base64
            data = request.get_json(silent=True) or {}
            nome_arquivo = data.get("nome_arquivo", "")
            conteudo_base64 = data.get("conteudo_base64", "")
            if not conteudo_base64:
                return jsonify({"ok": False, "msg": "Nada enviado."}), 400
            decoded = base64.b64decode(conteudo_base64.split(",")[1])
            buf = BytesIO(decoded)
            name = nome_arquivo.lower()

        if name.endswith(".csv"):
            df = pd.read_csv(buf)
        elif name.endswith(".xlsx"):
            df = pd.read_excel(buf, engine="openpyxl")
        elif name.endswith(".html"):
            df = pd.read_html(buf)[0]
        else:
            return jsonify({"ok": False, "msg": "Formato de arquivo não suportado."}), 400

        STATE["df"] = df
        return jsonify({"ok": True, "msg": "Arquivo carregado com sucesso.", "rows": len(df)})

    except Exception as e:
        return jsonify({"ok": False, "msg": f"Erro ao carregar: {e}"}), 500


@app.post("/tratar")
def tratar():
    """Mantem a mesma semantica do projeto original: filtro: sim | nao | -1 | continuar; agrupar: bool"""
    data = request.get_json()
    filtro = (data or {}).get("filtro")
    agrupar = bool((data or {}).get("agrupar", True))
    try:
        df = STATE["df"]
        if df is None or df.empty:
            return jsonify({"ok": False, "msg": "Carregue um arquivo antes."}), 400

        if filtro in ["sim", "nao", "-1"]:
            df = preprocess_dataframe(df)
            dec_rebanho = pd.to_numeric(df["Dec. Rebanho"], errors="coerce")
            if filtro == "sim":
                df = df[dec_rebanho == 1]
            elif filtro == "nao":
                df = df[dec_rebanho == 0]
            else:
                df = df[dec_rebanho == -1]
        elif filtro == "outros_inadimplentes":
            df = preprocess_outros_animais_inadimplentes(df)
        elif filtro == "continuar":
            if "Status" not in df.columns:
                return jsonify({"ok": False, "msg": "Arquivo nao contem dados ja tratados para continuar."}), 400
            df = df[df["Status"] == "Fila de envio"]
        else:
            # nenhum filtro definido: apenas normaliza se ainda nao tiver Status
            if "Status" not in df.columns:
                try:
                    df = preprocess_dataframe(df)
                except Exception:
                    df = preprocess_outros_animais_inadimplentes(df)

        if agrupar:
            df = (
                df.groupby(["Status", "Telefone"])["Nome"]
                .apply(lambda x: " || ".join(x))
                .reset_index()
            )
            df["Status"] = "Fila de envio"

        STATE["df"] = df.reset_index(drop=True)
        return jsonify({"ok": True, "msg": f"{len(df)} contatos preparados para envio."})
    except Exception as e:
        return jsonify({"ok": False, "msg": f"Erro ao tratar dados: {e}"}), 500


@app.get("/tabela")
def tabela():
    html = df_to_visible_html(STATE["df"])
    return html


@app.get("/stats")
def stats():
    return jsonify(get_stats(STATE["df"]))


@app.post("/start_whatsapp")
def start_whatsapp():
    try:
        if STATE["driver"] is None:
            STATE["driver"] = start_driver()
        # Espera elemento "side" (lista de conversas) aparecer (limite brando)
        # Deixa o front avisar o usuário para escanear QR, se necessário.
        return jsonify({"ok": True, "msg": "WhatsApp Web aberto. Escaneie o QR (se necessário)."})
    except Exception as e:
        return jsonify({"ok": False, "msg": f"Erro ao iniciar WhatsApp: {e}"}), 500


@app.post("/send")
def send():
    data = request.get_json() or {}
    tempo_min = max(1, int(data.get("tempo_min", 25)))
    tempo_max = max(tempo_min, int(data.get("tempo_max", 37)))
    mensagem_template = str(data.get("mensagem_template", "Olá -&nome, contato: -&numero"))
    mensagens_min = max(1, int(data.get("mensagens_min", 30)))
    mensagens_max = max(mensagens_min, int(data.get("mensagens_max", 40)))
    minutos_min = max(0, int(data.get("minutos_min", 1)))
    minutos_max = max(minutos_min, int(data.get("minutos_max", 3)))

    if STATE["sending_thread"] and STATE["sending_thread"].is_alive():
        return jsonify({"ok": False, "msg": "Envio já está em andamento."}), 400

    STATE["stop_flag"] = False
    t = threading.Thread(
        target=run_sending,
        args=(
            tempo_min,
            tempo_max,
            mensagem_template,
            mensagens_min,
            mensagens_max,
            minutos_min,
            minutos_max,
        ),
        daemon=True,
    )
    STATE["sending_thread"] = t
    t.start()
    return jsonify({"ok": True, "msg": "Envio iniciado."})


@app.post("/stop")
def stop():
    STATE["stop_flag"] = True
    return jsonify({"ok": True, "msg": "Sinal de parada enviado."})


@app.get("/download")
def download():
    df = STATE["df"]
    if df is None or df.empty:
        return jsonify({"ok": False, "msg": "Sem dados."}), 400
    out = BytesIO()
    df.to_excel(out, index=False)
    out.seek(0)
    return send_file(out, as_attachment=True, download_name="BancoProd.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/stream")
def stream():
    """Server-Sent Events: o front consome /stream para receber logs em tempo real."""
    def event_stream():
        while True:
            msg = STATE["logs"].get()
            yield f"data: {msg}\n\n"
    return Response(event_stream(), mimetype="text/event-stream")


if __name__ == "__main__":
    # Em rede local, pode pôr host="0.0.0.0"
    app.run(debug=True, host="127.0.0.1", port=5000)
