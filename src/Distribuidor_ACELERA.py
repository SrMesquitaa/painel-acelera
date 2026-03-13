import time, gspread, os, threading, sys, ctypes, logging, requests
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials


# --- BARRA DE TAREFAS ---
try:
    myappid = "minhaempresa.extrator.planium.1.4"
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except:
    pass


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# --- LOGS ---
def configurar_logs():
    base = os.path.dirname(os.path.abspath(sys.argv[0]))
    caminho_logs = os.path.join(base, "logs")
    os.makedirs(caminho_logs, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_log = os.path.join(caminho_logs, f"execucao_{timestamp}.txt")
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    file_handler = logging.FileHandler(arquivo_log)
    stream_handler = logging.StreamHandler()
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    file_handler.setFormatter(formatter)
    stream_handler.setFormatter(formatter)
    logging.basicConfig(level=logging.INFO, handlers=[file_handler, stream_handler])


# --- CONFIGURAÇÃO ---
URL_PLANILHA = "https://docs.google.com/spreadsheets/d/16xD6EAIkfL7bZYfLSKnilBq4hceJbotSkfDE8MVNXg0/edit"
NOME_ABA = "leads"
URL_WEBHOOK = (
    "https://n8n.jetsalesbrasil.com/webhook/22f5b5ed-7e1f-4dc3-8a8d-e479629b7504"
)
MODO_TESTE = False  # True = envia só 1 lead | False = envia todos
INTERVALO_SEGUNDOS = 20  # ⏱️ Intervalo entre envios (segundos)

# Índices das colunas (A=0, B=1 ...)
COL_DATA = 0
COL_CPF = 1
COL_NOME = 2
COL_EMAIL = 3
COL_TELEFONE = 4
COL_PRACA = 5
COL_VENDEDOR = 6
COL_OBS = 7

robo_rodando = False


def conectar_planilha():
    caminho_json = resource_path("credentials.json")
    if not os.path.exists(caminho_json):
        raise FileNotFoundError(f"credentials.json não encontrado em: {caminho_json}")
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(caminho_json, scope)
    client = gspread.authorize(creds)
    return client.open_by_url(URL_PLANILHA).worksheet(NOME_ABA)


def normalizar_data(valor):
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%y"]:
        try:
            return datetime.strptime(str(valor).strip(), fmt).strftime("%Y-%m-%d")
        except:
            continue
    return str(valor).strip()


def enviar_webhook(linha):
    try:
        data_fmt = datetime.strptime(linha[COL_DATA], "%d/%m/%Y").strftime("%Y-%m-%d")
    except:
        try:
            data_fmt = datetime.strptime(linha[COL_DATA], "%Y-%m-%d").strftime(
                "%Y-%m-%d"
            )
        except:
            data_fmt = linha[COL_DATA]
    payload = {
        "dataCriacao": data_fmt,
        "cpf": linha[COL_CPF],
        "nome": linha[COL_NOME],
        "email": linha[COL_EMAIL],
        "telefone": linha[COL_TELEFONE],
        "praca": linha[COL_PRACA],
        "vendedor": linha[COL_VENDEDOR] if len(linha) > COL_VENDEDOR else "",
        "observacao1": linha[COL_OBS] if len(linha) > COL_OBS else "",
        "conversao": "",
        "filial": "",
        "observacao2": "",
        "remarketing1": "",
        "remarketing2": "",
        "remarketing3": "",
        "supervisor": "",
    }
    response = requests.post(
        URL_WEBHOOK,
        json=payload,
        headers={"Content-Type": "application/json"},
        timeout=15,
    )
    return response.status_code


# --- ROBÔ ---
def run_robo():
    global robo_rodando
    try:
        configurar_logs()
        data_alvo = (datetime.now() - timedelta(days=2)).strftime("%Y-%m-%d")
        data_exibir = (datetime.now() - timedelta(days=2)).strftime("%d/%m/%Y")
        logging.info(f">>> Iniciando - leads de {data_exibir} <<<")
        logging.info(f">>> Intervalo entre envios: {INTERVALO_SEGUNDOS}s <<<")
        try:
            aba = conectar_planilha()
            todas_linhas = aba.get_all_values()
            leads = todas_linhas[1:]
            logging.info(f"Total de linhas na planilha: {len(leads)}")
        except Exception as e:
            finalizar(f"Erro ao acessar a planilha:\n{e}", erro=True)
            return
        leads_filtrados = []
        for linha in leads:
            while len(linha) < 8:
                linha.append("")
            data_lead = normalizar_data(linha[COL_DATA])
            nome = linha[COL_NOME].strip()
            logging.info(
                f"Verificando: {nome} | data_lead={data_lead} | data_alvo={data_alvo}"
            )
            if nome and data_lead == data_alvo:
                leads_filtrados.append(linha)
        total_na_data = len(leads_filtrados)
        logging.info(f"Leads encontrados para {data_exibir}: {total_na_data}")
        if total_na_data == 0:
            finalizar(f"Nenhum lead encontrado para {data_exibir}.")
            return
        janela.after(
            0,
            lambda: lbl_status.config(
                text=f"Status: 0/{total_na_data} enviados", foreground="green"
            ),
        )
        total_enviados = 0
        for i, linha in enumerate(leads_filtrados):
            if not robo_rodando:
                logging.info("Parado pelo usuário.")
                break
            nome = linha[COL_NOME].strip()
            cpf = linha[COL_CPF].strip()
            try:
                status = enviar_webhook(linha)
                if status == 200:
                    total_enviados += 1
                    logging.info(
                        f"[{total_enviados}/{total_na_data}] ENVIADO: {nome} | CPF: {cpf}"
                    )
                    janela.after(
                        0,
                        lambda e=total_enviados, t=total_na_data: lbl_status.config(
                            text=f"Status: {e}/{t} enviados", foreground="green"
                        ),
                    )
                else:
                    logging.error(f"FALHA [{status}]: {nome} | CPF: {cpf}")
            except Exception as e:
                logging.error(f"ERRO ao enviar {nome}: {e}")
            if MODO_TESTE:
                logging.info(">>> MODO TESTE: apenas 1 lead enviado. <<<")
                break
            if i < len(leads_filtrados) - 1:
                for _ in range(INTERVALO_SEGUNDOS):
                    if not robo_rodando:
                        break
                    time.sleep(1)
        logging.info(
            f"=== Resumo: {total_enviados} enviados de {total_na_data} leads da data ==="
        )
        finalizar(
            f"Concluído!\n\n📅 Data processada: {data_exibir}\n✅ Leads enviados: {total_enviados}\n📋 Total encontrado: {total_na_data}"
        )

    except Exception as e:
        # garante que QUALQUER erro apareça
        janela.after(0, lambda: messagebox.showerror("Erro crítico", str(e)))
        robo_rodando = False


# --- INTERFACE ---
def finalizar(msg, erro=False):
    global robo_rodando
    robo_rodando = False

    def _atualizar_ui():
        loading.stop()
        btn_executar.config(state="normal")
        lbl_status.config(text="Status: Parado", foreground="black")
        if erro:
            messagebox.showerror("Erro", msg)
        else:
            messagebox.showinfo("Fim", msg)

    janela.after(0, _atualizar_ui)


def acao_executar():
    global robo_rodando
    if robo_rodando:
        return
    configurar_logs()
    robo_rodando = True
    lbl_status.config(text="Status: Iniciando...", foreground="green")
    btn_executar.config(state="disabled")
    loading.start(10)
    threading.Thread(target=run_robo, daemon=True).start()


def acao_parar():
    global robo_rodando
    robo_rodando = False
    lbl_status.config(text="Status: Parando...", foreground="red")


# --- JANELA ---
data_prev = (datetime.now() - timedelta(days=2)).strftime("%d/%m/%Y")
janela = tk.Tk()
janela.title(
    "Extrator Lidera SP v1.5 - MODO TESTE" if MODO_TESTE else "Extrator Lidera SP v1.5"
)
janela.geometry("360x290")
janela.resizable(False, False)

tk.Label(janela, text="Extrator Lidera SP", font=("Arial", 13, "bold")).pack(
    pady=(15, 4)
)
tk.Label(
    janela,
    text=f"Distribuindo leads de:  {data_prev}",
    font=("Arial", 10),
    fg="#2c3e50",
).pack(pady=(0, 4))

if MODO_TESTE:
    tk.Label(
        janela,
        text="⚠️ MODO TESTE — apenas 1 lead será enviado",
        font=("Arial", 9, "bold"),
        fg="#e67e22",
    ).pack(pady=(0, 4))

tk.Label(
    janela,
    text=f"⏱ Intervalo: {INTERVALO_SEGUNDOS}s por lead",
    font=("Arial", 9),
    fg="#7f8c8d",
).pack(pady=(0, 4))

lbl_status = tk.Label(janela, text="Status: Parado", font=("Arial", 9))
lbl_status.pack(pady=4)

loading = ttk.Progressbar(janela, mode="indeterminate", length=220)
loading.pack(pady=4)

btn_executar = tk.Button(
    janela, text="EXECUTAR", command=acao_executar, bg="#27ae60", fg="white", width=15
)
btn_executar.pack(pady=5)

btn_parar = tk.Button(
    janela, text="PARAR", command=acao_parar, bg="#c0392b", fg="white"
)
btn_parar.pack(pady=5)
janela.mainloop()
