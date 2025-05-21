# main.py
import webview
import pandas as pd
import urllib.parse
import time
import random
import os
import base64
from io import BytesIO
import json
from threading import Thread
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

class API:
    def __init__(self):
        self.df = None
        self.driver = None

    def obter_dados_tratados(self):
        if self.df is not None and not self.df.empty:
            df_visivel = self.df.copy()
            if "Dec. Rebanho" in df_visivel.columns:
                df_visivel = df_visivel.drop(columns=["Dec. Rebanho"])
            for col in df_visivel.columns:
                df_visivel[col] = df_visivel[col].astype(str)
            return df_visivel.to_html(index=False, escape=False)
        return "<p>Nenhum dado disponível.</p>"

    def carregar_arquivo(self, nome_arquivo, conteudo_base64):
        try:
            decoded = base64.b64decode(conteudo_base64.split(",")[1])
            buffer = BytesIO(decoded)

            if nome_arquivo.endswith('.csv'):
                self.df = pd.read_csv(buffer)
            elif nome_arquivo.endswith('.xlsx'):
                self.df = pd.read_excel(buffer, engine='openpyxl')
            elif nome_arquivo.endswith('.html'):
                self.df = pd.read_html(buffer)[0]
            else:
                return "Formato de arquivo não suportado."

            return "Arquivo carregado com sucesso."
        except Exception as e:
            return f"Erro ao carregar: {e}"

    def tratar_dados(self, filtro, agrupar=True):
        def preprocess_dataframe(df):
            colunas = ['Nome do Titular da Ficha de bovideos', 'Nome da Propriedade', 'Endereço da Prop.', 'Dec. Rebanho', 'Telefone 1', 'Telefone 2', 'Celular']
            for col in colunas:
                if col not in df.columns:
                    raise ValueError(f"Coluna ausente: {col}")
            df = df[colunas]
            df = pd.melt(df, id_vars=colunas[:4], value_vars=colunas[4:], value_name='Telefone').drop(columns=['variable'])
            df['Nome'] = df.apply(lambda row: f"{row[colunas[0]]} - {row[colunas[1]]} - {row[colunas[2]]}", axis=1)
            df = df.drop(columns=colunas[:3])
            df['Telefone'] = df['Telefone'].astype(str).str.replace(r'[^0-9]', '', regex=True).str.zfill(10)
            df = df[~df['Telefone'].str.match(r'^\d{6}0000$')]
            df['Telefone'] = '+55' + df['Telefone']
            df['Telefone'] = df['Telefone'].apply(lambda telefone: telefone[:5] + telefone[6:] if len(telefone) == 15 else telefone)
            df['Telefone'] = df['Telefone'].str[:3] + ' ' + df['Telefone'].str[3:5] + ' ' + df['Telefone'].str[5:]
            df["Status"] = "Fila de envio"
            return df[["Status", "Nome", "Telefone", "Dec. Rebanho"]]

        try:
            if filtro in ["sim", "nao"]:
                self.df = preprocess_dataframe(self.df)
                if filtro == "sim":
                    self.df = self.df[self.df["Dec. Rebanho"] == 1]
                elif filtro == "nao":
                    self.df = self.df[self.df["Dec. Rebanho"] == 0]
            elif filtro == "continuar":
                if "Status" not in self.df.columns:
                    return "Arquivo não contém dados já tratados para continuar."
                self.df = self.df[self.df["Status"] == "Fila de envio"]

            if agrupar:
                self.df = self.df.groupby(["Status", "Telefone"])["Nome"].apply(lambda x: " || ".join(x)).reset_index()
                self.df["Status"] = "Fila de envio"

            return f"{len(self.df)} contatos preparados para envio."
        except Exception as e:
            return f"Erro ao tratar dados: {str(e)}"
    def obter_estatisticas(self):
        try:
            total = len(self.df)
            fila = len(self.df[self.df["Status"] == "Fila de envio"])
            enviados = len(self.df[self.df["Status"] == "Enviado"])
            invalidos = len(self.df[self.df["Status"] == "Número inválido"])

            def perc(qtd):
                return f"{(qtd / total * 100):.1f}%" if total > 0 else "0%"

            resumo = {
                "total": total,
                "fila": {"qtd": fila, "perc": perc(fila)},
                "enviados": {"qtd": enviados, "perc": perc(enviados)},
                "invalidos": {"qtd": invalidos, "perc": perc(invalidos)}
            }

            return resumo
        except Exception as e:
            return {"erro": str(e)}

    def iniciar_whatsapp(self):
   

        # Corrige o caminho do ChromeDriver
        path = ChromeDriverManager().install()
        folder = os.path.dirname(path)
        chromedriver_path = os.path.join(folder, "chromedriver.exe")  # Força o executável correto

        options = Options()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')

        self.driver = webdriver.Chrome(service=Service(executable_path=chromedriver_path), options=options)
        self.driver.get('https://web.whatsapp.com')

        while len(self.driver.find_elements(By.ID, 'side')) < 1:
            time.sleep(1)

        self._log_frontend("WhatsApp Web está pronto para envio.")
        return "WhatsApp Web pronto para envio."
   
    def _log_frontend(self, mensagem, log_id=None):
            try:
                if log_id:
                    js_code = f"window.adicionarLog({json.dumps(mensagem)}, {json.dumps(log_id)})"
                else:
                    js_code = f"window.adicionarLog({json.dumps(mensagem)})"
                webview.windows[0].evaluate_js(js_code)
            except Exception as e:
                print(f"[LOG ERRO] Falha ao enviar log para o frontend: {e}")

    def enviar_mensagens(self, tempo_min, tempo_max, mensagem_template):
        def send():
            print("[DEBUG] Iniciando envio de mensagens...")
            webview.windows[0].evaluate_js("clearLogs()")

            if self.df is None or self.df.empty:
                self._log_frontend("Nenhum dado disponível para envio.")
                return

            for index, row in self.df[self.df['Status'] == 'Fila de envio'].iterrows():
                webview.windows[0].evaluate_js("clearLogs()")
                numero = row['Telefone']
                nome = row['Nome']
                mensagem = mensagem_template.replace("-&numero", numero).replace("-&nome", nome)

                self._log_frontend(f"Enviando para {nome} ({numero})")
                sucesso = self._enviar(numero, mensagem)

                if sucesso:
                    self.df.at[index, 'Status'] = 'Enviado'
                    self._log_frontend(f"Mensagem enviada para {nome}")

                    delay = random.randint(tempo_min, tempo_max)
                    log_id = f"log_delay_{index}"

                    webview.windows[0].evaluate_js(f"""
                        var el = document.createElement('p');
                        el.id = {json.dumps(log_id)};
                        el.style.fontFamily = 'monospace';
                        el.textContent = 'Próximo envio em {delay} segundos...';
                        document.getElementById('logEnvio').appendChild(el);
                        document.getElementById('logEnvio').scrollTop = document.getElementById('logEnvio').scrollHeight;
                    """)

                    for i in range(delay, 0, -1):
                        webview.windows[0].evaluate_js(
                            f"document.getElementById({json.dumps(log_id)}).textContent = 'Próximo envio em {i} segundos...';"
                        )
                        time.sleep(1)

                    webview.windows[0].evaluate_js("clearDelayLogs()")

                else:
                    self.df.at[index, 'Status'] = 'Número inválido'
                    self._log_frontend(f"⚠️ Falha ao enviar para {nome} ({numero}) — número inválido.")
                    self._log_frontend("⏭️ Pulando para o próximo contato...")

                self.df.to_excel('BancoProd.xlsx', index=False)

        Thread(target=send).start()
        return "Envio iniciado."

   
    def _enviar(self, numero, mensagem):
        try:
            numero = numero.replace(" ", "")
            mensagem = urllib.parse.quote(mensagem)
            self.driver.get(f"https://web.whatsapp.com/send?phone={numero}&text={mensagem}")

            # Aguarda até 8 segundos para encontrar o campo de digitação
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"]'))
            )

            time.sleep(1)
            botao_enviar = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
            )
            botao_enviar.click()
            time.sleep(1)
            return True

        except Exception as e:
            print(f"[ERRO] Número inválido ou falha ao enviar para {numero}: {e}")
            return False



    def baixar_planilha(self):
        try:
            import tkinter as tk
            from tkinter import filedialog

            # Esconde a janela principal do Tkinter
            root = tk.Tk()
            root.withdraw()

            # Abre diálogo de "Salvar Como"
            caminho_salvar = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Planilha Excel", "*.xlsx")],
                title="Salvar planilha como",
                initialfile="BancoProd.xlsx"
            )

            if not caminho_salvar:
                return "Operação cancelada pelo usuário."

            # Salva a planilha no local escolhido
            self.df.to_excel(caminho_salvar, index=False)
            os.startfile(os.path.dirname(caminho_salvar))

            return f"Planilha salva em: {caminho_salvar}"

        except Exception as e:
            return f"Erro ao salvar planilha: {e}"


if __name__ == '__main__':
    api = API()
    webview.create_window("Envio de Mensagens", "interface.html", js_api=api, width=900, height=700)
    webview.start(debug=False)
