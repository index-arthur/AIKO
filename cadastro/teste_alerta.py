import os
import sys
import json
import time
import random
import tempfile
import threading
import subprocess
import webbrowser
import urllib.request
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException

# ==================== CONFIG ====================
VERSION = "4.6"
REPO_OWNER = "index-arthur"
REPO_NAME = "AIKO"
GITHUB_API_LATEST = (
    f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest"
)
RELEASES_URL = f"https://github.com/{REPO_OWNER}/{REPO_NAME}/releases/latest"

# ⚠️ MODO DE TESTE DO UPDATER (deixar False em produção!)
#   True  → força o banner aparecer; o "Atualizar" simula o download
#            (não substitui o .exe de verdade)
#   False → comportamento real (consulta o GitHub, baixa e substitui)
MODO_TESTE_UPDATE = False

GRUPOS_DEFAULT = ["BAR-VN-000", "reserva", "teste", "inativos"]
MODEL_DEFAULT  = ["AXOR 3344", "feller", "basculante", "forwarder"]
PERFIL_DEFAULT = ["equipamentos"]

TUTORIAL_TXT = (
    "INFORMAÇÕES NECESSÁRIAS:\n\n"
    "• Usuário        → Seu usuário sem o @aiko.digital (ex: acosta)\n"
    "• Senha          → Sua senha de acesso\n"
    "• Empresa        → Sigla da empresa (ex: BRC, RAI, QA.BRC)\n"
    "• Equipamento    → Tipo (ex: COMODATO, SERVICO DE CAMPO)\n"
    "• Ticket         → Número do ticket (só os números)\n"
    "• Zendesk        → Número do ticket Zendesk (deixe vazio se não tiver)\n"
    "• Parou no bordo → Se parou em algum bordo, o número dele (0 se não parou)\n"
    "• Qtd. Bordos    → Total de bordos a cadastrar\n"
    "• Grupos         → Termos para achar o grupo no dropdown\n"
    "• Modelos        → Termos para achar o modelo no dropdown\n"
    "• Perfil         → Perfil de rede (padrão: equipamentos)\n"
)

# ==================== UPDATE CHECK ====================
def checar_atualizacao(timeout=5):
    """Consulta a Releases API do GitHub.

    Retorna um dict com:
      tem_update     → bool
      versao_remota  → str  (tag sem 'v')
      exe_url        → str  (URL de download do .exe, ou None)
      release_url    → str  (URL html da release)
    """
    vazio = {"tem_update": False, "versao_remota": None,
             "exe_url": None, "release_url": RELEASES_URL}

    if MODO_TESTE_UPDATE:
        return {"tem_update": True, "versao_remota": "9.9",
                "exe_url": None, "release_url": RELEASES_URL}

    try:
        req = urllib.request.Request(
            GITHUB_API_LATEST,
            headers={
                "User-Agent": "AikoCadastroHUD",
                "Accept": "application/vnd.github+json",
            },
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        tag = (data.get("tag_name") or "").lstrip("v").strip()
        exe_asset = next(
            (a for a in data.get("assets", [])
             if a.get("name", "").lower().endswith(".exe")),
            None,
        )
        return {
            "tem_update": bool(tag) and tag != VERSION,
            "versao_remota": tag or None,
            "exe_url": exe_asset["browser_download_url"] if exe_asset else None,
            "release_url": data.get("html_url") or RELEASES_URL,
        }
    except Exception:
        return vazio

def baixar_nova_versao(exe_url, progresso_cb, timeout=60):
    """Baixa o .exe para um arquivo temporário. Retorna o caminho local."""
    if MODO_TESTE_UPDATE:
        # Simula download com progresso falso
        for i in range(0, 101, 4):
            progresso_cb(i)
            time.sleep(0.08)
        return None

    temp_dir = tempfile.gettempdir()
    destino = os.path.join(temp_dir, f"Cadastro_update_{os.getpid()}.exe")

    req = urllib.request.Request(
        exe_url, headers={"User-Agent": "AikoCadastroHUD"}
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        total = int(resp.headers.get("Content-Length") or 0)
        baixado = 0
        with open(destino, "wb") as f:
            while True:
                chunk = resp.read(64 * 1024)
                if not chunk:
                    break
                f.write(chunk)
                baixado += len(chunk)
                if total:
                    progresso_cb(min(99, baixado / total * 100))
    progresso_cb(100)
    return destino

def aplicar_atualizacao(exe_novo_path):
    """Agenda a substituição do .exe atual via script .bat.

    IMPORTANTE: esta função NÃO encerra o programa. Quem chama é
    responsável por chamar os._exit(0) no thread principal para que
    o processo realmente feche e libere o .exe para substituição.
    """
    if not getattr(sys, "frozen", False):
        raise RuntimeError(
            "Auto-update só funciona no .exe compilado.\n"
            "Em modo dev (python teste_alerta.py), baixe manualmente."
        )

    exe_atual = sys.executable
    temp_dir = tempfile.gettempdir()
    bat_path = os.path.join(temp_dir, f"_aiko_update_{os.getpid()}.bat")

    bat = (
        f'@echo off\r\n'
        f'chcp 65001 > nul\r\n'
        f'timeout /t 3 /nobreak > nul\r\n'
        f':retry\r\n'
        f'move /y "{exe_novo_path}" "{exe_atual}" > nul 2>&1\r\n'
        f'if errorlevel 1 (\r\n'
        f'  timeout /t 1 /nobreak > nul\r\n'
        f'  goto retry\r\n'
        f')\r\n'
        # Espera 4s depois do move para o Windows Defender terminar de
        # escanear o novo arquivo. Sem essa pausa, o 'start' logo abaixo
        # dispara um erro "Failed to load Python DLL" porque o arquivo
        # ainda está lockado pelo AV enquanto o PyInstaller tenta extrair.
        f'timeout /t 4 /nobreak > nul\r\n'
        f'start "" "{exe_atual}"\r\n'
        f'del "%~f0"\r\n'
    )
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(bat)

    subprocess.Popen(
        ["cmd", "/c", bat_path],
        creationflags=0x00000008,  # DETACHED_PROCESS
    )

# ==================== SELENIUM HELPERS ====================
def pausa(min_s=0.5, max_s=3.5):
    time.sleep(random.uniform(min_s, max_s))

def clicar_dropdown(driver, xpaths_botao, xpaths_opcoes, condicao):
    wait_curto = WebDriverWait(driver, 3)
    wait_normal = WebDriverWait(driver, 25)

    for xpath in xpaths_botao:
        try:
            wait_curto.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
            break
        except TimeoutException:
            continue

    opcoes = []
    for xpath in xpaths_opcoes:
        try:
            wait_normal.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            opcoes = driver.find_elements(By.XPATH, xpath + "//li")
            if opcoes:
                break
        except TimeoutException:
            continue

    for opcao in opcoes:
        texto = opcao.text.strip().lower()
        if condicao(texto):
            opcao.click()
            return True
    return False

def login(driver, usuario, senha):
    wait_curto = WebDriverWait(driver, 3)
    wait_normal = WebDriverWait(driver, 25)
    try:
        botao = wait_curto.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='social-azuread-aiko']/span")))
        driver.execute_script("arguments[0].click();", botao)
    except TimeoutException:
        wait_normal.until(EC.visibility_of_element_located((By.ID, "UserName")))\
            .send_keys(usuario)
        driver.find_element(By.ID, "Password").send_keys(senha, Keys.ENTER)

def fechar_emergencias(driver):
    while True:
        botoes = driver.find_elements(
            By.XPATH, '//*[@id="legacy-global-modals"]//button')
        if not botoes:
            break
        for b in botoes:
            try:
                b.click()
                time.sleep(0.5)
            except Exception:
                pass

def click_forcado(driver, el):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        ActionChains(driver).move_to_element(el).pause(0.1).click(el).perform()
    except Exception:
        pass

def clicar_salvar_ultimo_visivel(driver, timeout=8):
    wait_local = WebDriverWait(driver, timeout)
    wait_local.until(EC.visibility_of_element_located((By.ID, "name")))
    spans = driver.find_elements(
        By.XPATH, "//*[@id='app']//span[normalize-space()='Salvar']")
    spans = [s for s in spans if s.is_displayed()]
    span_salvar = spans[-1]
    clicavel = span_salvar.find_element(
        By.XPATH, "./ancestor::button[1] | ./ancestor::a[1]")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", clicavel)
    driver.execute_script("arguments[0].click();", clicavel)

def confirmar_sim_se_existir(driver, timeout=3):
    wait_local = WebDriverWait(driver, timeout)
    seletores = [
        (By.CSS_SELECTOR, "button[data-test-ak-confirm-dialog-btn-confirm]"),
        (By.XPATH, "//button[normalize-space()='Sim']"),
        (By.XPATH, '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[18]/div[2]/div/div[2]/button[2]'),
        (By.XPATH, '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[19]/div[2]/div/div[2]/button[2]'),
    ]
    for by, sel in seletores:
        try:
            sim = wait_local.until(EC.element_to_be_clickable((by, sel)))
            click_forcado(driver, sim)
            time.sleep(0.3)
            return True
        except TimeoutException:
            continue
    return False

def esperar_loading(driver, timeout=25):
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.ID, "waiting-update")))
    except TimeoutException:
        pass

# ==================== AUTOMAÇÃO ====================
def executar_automacao(dados, log):
    empresa = dados["empresa"]
    driver = webdriver.Edge()
    driver.maximize_window()
    driver.get(f"https://{empresa}.br.trackit.host/")

    log("Fazendo login...")
    login(driver, dados["usuario"], dados["senha"])

    wait = WebDriverWait(driver, 20)
    fechar_emergencias(driver)

    log("Abrindo cadastro de bordo...")
    wait.until(EC.presence_of_element_located((By.ID, "nav-menu")))
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="nav-menu"]/div[1]/div[2]/a'))).click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="nav-menu"]/div[1]/div[2]/div/div/section[1]/article/ul/li[4]/a'))).click()

    usa_zendesk = dados["zendesk"].lower() != "n"
    zendesk_txt = f" | #{dados['zendesk']}" if usa_zendesk else ""

    inicio = dados["parou"] + 1
    fim = dados["limite"] + 1

    for k in range(inicio, fim):
        numero = f"{k:02d}"
        bordo = (f"{empresa} | {dados['equipamento']} | "
                 f"HWS-{dados['ticket']}{zendesk_txt} | {numero}")
        log(f"[{k}/{dados['limite']}] {bordo}")

        fechar_emergencias(driver)
        wait.until(EC.presence_of_element_located((By.ID, "app")))
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="app"]/div[1]/div[5]/div/section/div/div[2]/a/span'))).click()
        pausa(1, 1.3)

        fechar_emergencias(driver)
        wait.until(EC.element_to_be_clickable((By.ID, "name"))).send_keys(bordo)
        pausa(1, 1.3)

        fechar_emergencias(driver); esperar_loading(driver)
        clicar_dropdown(driver,
            xpaths_botao=[
                '//*[@id="equipmentModel"]/div[1]',
                '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[4]/div[1]/div/div/div[1]',
            ],
            xpaths_opcoes=[
                '//*[@id="equipmentModel"]/div[3]',
                '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[4]/div[1]/div/div/div[3]',
            ],
            condicao=lambda t: any(p.lower() in t for p in dados["modelos"]))
        pausa(0.5, 0.8)

        fechar_emergencias(driver); esperar_loading(driver)
        clicar_dropdown(driver,
            xpaths_botao=[
                '//*[@id="networkProfiles"]/div[1]',
                '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[4]/div[3]/div/div/div[1]',
            ],
            xpaths_opcoes=[
                '//*[@id="networkProfiles"]/div[3]',
                '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[4]/div[3]/div/div/div[3]',
            ],
            condicao=lambda t: any(p.lower() in t for p in dados["perfil"]))
        pausa(0.5, 0.8)

        fechar_emergencias(driver); esperar_loading(driver)
        clicar_dropdown(driver,
            xpaths_botao=[
                '//*[@id="equipmentGroup"]/div[1]',
                '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[7]/div[1]/div/div/div/div[1]',
            ],
            xpaths_opcoes=[
                '//*[@id="equipmentGroup"]/div[3]',
                '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[7]/div[1]/div/div/div/div[3]',
            ],
            condicao=lambda t: any(p.lower() in t for p in dados["grupos"]))
        pausa(0.5, 0.8)

        fechar_emergencias(driver)
        clicar_salvar_ultimo_visivel(driver, timeout=8)
        confirmar_sim_se_existir(driver, timeout=3)
        wait.until(EC.invisibility_of_element_located((By.ID, "name")))
        pausa(0.5, 1)

    log("Finalizado.")

# ==================== HUD ====================
class CadastroHUD(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Cadastro de Bordo - Aiko  •  v{VERSION}")
        self.geometry("640x740")
        self.minsize(580, 680)
        try:
            self.iconbitmap("negao.ico")
        except Exception:
            pass

        self._estilo()
        self._montar_layout()
        self.after(200, self._verificar_update_async)

    # Paleta do tema escuro
    BG       = "#1e1e1e"   # fundo principal da janela
    SURFACE  = "#2d2d30"   # caixas de texto, log, dropdowns
    BORDER   = "#3e3e42"   # bordas sutis
    TEXT     = "#e0e0e0"   # texto principal
    SUBTLE   = "#9a9a9a"   # texto secundário
    ACCENT   = "#0e639c"   # botão primário
    ACCENT_2 = "#1177bb"   # hover do primário
    UPDATE_BG   = "#3a3500"  # banner amarelo discreto
    UPDATE_FG   = "#ffd966"
    UPDATE_LINK = "#4ea1ff"

    def _estilo(self):
        # Fundo da janela principal
        self.configure(bg=self.BG)

        style = ttk.Style(self)
        # 'clam' aceita customização de cor melhor que 'vista' no Windows
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # ---------- Containers ----------
        style.configure("TFrame", background=self.BG)
        style.configure("TLabelframe",
                        background=self.BG, foreground=self.TEXT,
                        bordercolor=self.BORDER, lightcolor=self.BORDER,
                        darkcolor=self.BORDER)
        style.configure("TLabelframe.Label",
                        background=self.BG, foreground=self.TEXT)

        # ---------- Labels ----------
        style.configure("TLabel",
                        background=self.BG, foreground=self.TEXT,
                        font=("Segoe UI", 9))
        style.configure("Header.TLabel",
                        background=self.BG, foreground=self.TEXT,
                        font=("Segoe UI", 14, "bold"))
        style.configure("Sub.TLabel",
                        background=self.BG, foreground=self.SUBTLE,
                        font=("Segoe UI", 9))

        # ---------- Banner de update ----------
        style.configure("Update.TFrame", background=self.UPDATE_BG)
        style.configure("Update.TLabel",
                        background=self.UPDATE_BG,
                        foreground=self.UPDATE_FG,
                        font=("Segoe UI", 9, "bold"))
        style.configure("UpdateLink.TLabel",
                        background=self.UPDATE_BG,
                        foreground=self.UPDATE_LINK,
                        font=("Segoe UI", 9, "underline"))

        # ---------- Entry ----------
        style.configure("TEntry",
                        fieldbackground=self.SURFACE,
                        background=self.SURFACE,
                        foreground=self.TEXT,
                        insertcolor=self.TEXT,
                        bordercolor=self.BORDER,
                        lightcolor=self.BORDER,
                        darkcolor=self.BORDER)
        style.map("TEntry",
                  fieldbackground=[("readonly", self.SURFACE),
                                   ("disabled", self.BG)],
                  foreground=[("disabled", self.SUBTLE)])

        # ---------- Buttons ----------
        style.configure("TButton",
                        background=self.SURFACE, foreground=self.TEXT,
                        bordercolor=self.BORDER,
                        lightcolor=self.SURFACE, darkcolor=self.SURFACE,
                        focuscolor=self.BORDER,
                        font=("Segoe UI", 9), padding=(10, 4))
        style.map("TButton",
                  background=[("active", self.BORDER),
                              ("pressed", self.BORDER)],
                  foreground=[("disabled", self.SUBTLE)])

        style.configure("Primary.TButton",
                        background=self.ACCENT, foreground="white",
                        bordercolor=self.ACCENT,
                        lightcolor=self.ACCENT, darkcolor=self.ACCENT,
                        font=("Segoe UI", 10, "bold"), padding=(12, 5))
        style.map("Primary.TButton",
                  background=[("active", self.ACCENT_2),
                              ("pressed", self.ACCENT_2)])

        # ---------- Checkbutton ----------
        style.configure("TCheckbutton",
                        background=self.BG, foreground=self.TEXT,
                        indicatorcolor=self.SURFACE,
                        focuscolor=self.BG,
                        font=("Segoe UI", 9))
        style.map("TCheckbutton",
                  background=[("active", self.BG)],
                  indicatorcolor=[("selected", self.ACCENT),
                                  ("!selected", self.SURFACE)])

        # ---------- Progressbar ----------
        style.configure("TProgressbar",
                        background=self.ACCENT,
                        troughcolor=self.SURFACE,
                        bordercolor=self.BORDER,
                        lightcolor=self.ACCENT, darkcolor=self.ACCENT)

        # ---------- Scrollbar (do ScrolledText) ----------
        style.configure("Vertical.TScrollbar",
                        background=self.SURFACE,
                        troughcolor=self.BG,
                        bordercolor=self.BORDER,
                        arrowcolor=self.TEXT,
                        lightcolor=self.SURFACE, darkcolor=self.SURFACE)

    def _montar_layout(self):
        # Banner de update (inicia vazio, empacotado só se houver update)
        self.update_bar = ttk.Frame(self, style="Update.TFrame", padding=(10, 6))
        self.update_lbl = ttk.Label(self.update_bar, style="Update.TLabel", text="")
        self.update_link = ttk.Label(self.update_bar, style="UpdateLink.TLabel",
                                     text="Abrir release →", cursor="hand2")
        self.update_link.bind("<Button-1>", lambda e: webbrowser.open(RELEASES_URL))
        self.update_lbl.pack(side="left")
        self.update_link.pack(side="right")

        # Cabeçalho
        header = ttk.Frame(self, padding=(16, 14, 16, 4))
        header.pack(fill="x")
        ttk.Label(header, text="Automação de Cadastro de Bordo",
                  style="Header.TLabel").pack(anchor="w")
        ttk.Label(header, text=f"Versão {VERSION} — Trackit / Aiko",
                  style="Sub.TLabel").pack(anchor="w")

        # Campos
        body = ttk.Frame(self, padding=(16, 8, 16, 8))
        body.pack(fill="x")
        body.columnconfigure(1, weight=1)

        self.vars = {}

        def row(r, label, key, show=None, default=""):
            ttk.Label(body, text=label).grid(row=r, column=0, sticky="w", pady=4)
            v = tk.StringVar(value=default)
            e = ttk.Entry(body, textvariable=v, show=show)
            e.grid(row=r, column=1, sticky="ew", pady=4, padx=(8, 0))
            self.vars[key] = v

        row(0, "Usuário (sem @aiko.digital):", "usuario")
        row(1, "Senha:", "senha", show="•")
        row(2, "Empresa (sigla):", "empresa")
        row(3, "Equipamento:", "equipamento", default="COMODATO")
        row(4, "Ticket (número):", "ticket")
        row(5, "Ticket Zendesk (vazio = nenhum):", "zendesk")
        row(6, "Parou no bordo (0 = não parou):", "parou", default="0")
        row(7, "Qtd. total de bordos:", "limite")

        def lista(r, label, key, default_list):
            ttk.Label(body, text=label).grid(row=r, column=0, sticky="w", pady=4)
            frame = ttk.Frame(body)
            frame.grid(row=r, column=1, sticky="ew", pady=4, padx=(8, 0))
            frame.columnconfigure(0, weight=1)
            v = tk.StringVar(value=", ".join(default_list))
            e = ttk.Entry(frame, textvariable=v)
            e.grid(row=0, column=0, sticky="ew")
            usar_padrao = tk.BooleanVar(value=True)

            def toggle():
                if usar_padrao.get():
                    v.set(", ".join(default_list))
                    e.configure(state="disabled")
                else:
                    e.configure(state="normal")
            ttk.Checkbutton(frame, text="padrão", variable=usar_padrao,
                            command=toggle).grid(row=0, column=1, padx=(6, 0))
            toggle()
            self.vars[key] = v

        lista(8,  "Grupos:",  "grupos", GRUPOS_DEFAULT)
        lista(9,  "Modelos:", "modelos", MODEL_DEFAULT)
        lista(10, "Perfil:",  "perfil",  PERFIL_DEFAULT)

        # Log
        log_frame = ttk.LabelFrame(self, text="Log", padding=6)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(8, 4))
        self.log = scrolledtext.ScrolledText(log_frame, height=8,
                                             font=("Consolas", 9),
                                             state="disabled", wrap="word",
                                             bg=self.SURFACE, fg=self.TEXT,
                                             insertbackground=self.TEXT,
                                             selectbackground=self.ACCENT,
                                             selectforeground="white",
                                             borderwidth=0,
                                             highlightthickness=0)
        self.log.pack(fill="both", expand=True)

        # Botões
        botoes = ttk.Frame(self, padding=(16, 4, 16, 14))
        botoes.pack(fill="x")
        ttk.Button(botoes, text="Tutorial", command=self._abrir_tutorial)\
            .pack(side="left")
        self.btn_iniciar = ttk.Button(botoes, text="▶  Iniciar automação",
                                      style="Primary.TButton",
                                      command=self._on_iniciar)
        self.btn_iniciar.pack(side="right")

    # ----- Update -----
    def _verificar_update_async(self):
        def worker():
            info = checar_atualizacao()
            if info.get("tem_update"):
                self.after(0, lambda: self._mostrar_update(info))
        threading.Thread(target=worker, daemon=True).start()

    def _mostrar_update(self, info):
        remota = info["versao_remota"]
        self.update_lbl.configure(
            text=f"  Nova versão disponível: {remota}  (você tem {VERSION})"
        )
        # Troca o link padrão por um que abre o diálogo de atualização
        self.update_link.configure(text="Atualizar agora →")
        self.update_link.unbind("<Button-1>")
        self.update_link.bind(
            "<Button-1>", lambda e: self._dialogo_confirmar_update(info)
        )
        # Coloca o banner no topo da janela
        self.update_bar.pack(fill="x", before=self.winfo_children()[1])
        # Abre direto o diálogo de confirmação
        self._dialogo_confirmar_update(info)

    def _dialogo_confirmar_update(self, info):
        """Janela perguntando se quer atualizar agora ou depois."""
        remota = info["versao_remota"]
        top = tk.Toplevel(self)
        top.title("Atualização disponível")
        top.geometry("420x210")
        top.transient(self)
        top.grab_set()
        top.resizable(False, False)
        top.configure(bg=self.BG)

        ttk.Label(top, text=f"Nova versão: {remota}",
                  font=("Segoe UI", 13, "bold")).pack(pady=(18, 2))
        ttk.Label(top, text=f"Você está usando a {VERSION}.").pack()
        ttk.Label(top, text="Atualizar agora vai baixar a nova versão\n"
                            "e reiniciar o programa automaticamente.",
                  justify="center").pack(pady=(12, 0))

        btns = ttk.Frame(top)
        btns.pack(pady=14)

        def on_atualizar():
            top.destroy()
            self._dialogo_baixando_update(info)

        def on_depois():
            top.destroy()

        ttk.Button(btns, text="Atualizar agora", style="Primary.TButton",
                   command=on_atualizar).pack(side="left", padx=6)
        ttk.Button(btns, text="Depois", command=on_depois)\
            .pack(side="left", padx=6)

    def _dialogo_baixando_update(self, info):
        """Janela com barra de progresso durante download/aplicação."""
        exe_url = info.get("exe_url")
        if not exe_url and not MODO_TESTE_UPDATE:
            messagebox.showerror(
                "Sem arquivo disponível",
                "A release mais recente não tem um .exe anexado.\n"
                "Abra a página de releases e baixe manualmente."
            )
            webbrowser.open(info.get("release_url") or RELEASES_URL)
            return

        top = tk.Toplevel(self)
        top.title("Atualizando...")
        top.geometry("420x170")
        top.transient(self)
        top.grab_set()
        top.resizable(False, False)
        top.protocol("WM_DELETE_WINDOW", lambda: None)  # trava X durante download
        top.configure(bg=self.BG)

        ttk.Label(top, text="Baixando a nova versão...",
                  font=("Segoe UI", 10, "bold")).pack(pady=(18, 6))

        pbar = ttk.Progressbar(top, length=360, maximum=100, mode="determinate")
        pbar.pack(pady=6, padx=20, fill="x")

        status = tk.StringVar(value="Conectando...")
        ttk.Label(top, textvariable=status,
                  style="Sub.TLabel").pack(pady=(2, 0))

        def progresso(p):
            self.after(0, lambda p=p: (pbar.configure(value=p),
                                       status.set(f"{p:.0f}%")))

        def worker():
            try:
                caminho = baixar_nova_versao(exe_url, progresso)
                self.after(0, lambda: status.set(
                    "Download concluído. Reiniciando..."))
                time.sleep(0.6)

                if MODO_TESTE_UPDATE:
                    # Em modo teste, não mexe em nada: só avisa.
                    self.after(0, lambda: (
                        top.destroy(),
                        messagebox.showinfo(
                            "Simulação concluída",
                            "MODO_TESTE_UPDATE = True\n\n"
                            "Em produção, o .exe seria substituído e o "
                            "programa reiniciaria automaticamente."
                        )
                    ))
                    return

                # Produção: aplica e encerra (o .bat reinicia)
                aplicar_atualizacao(caminho)
                self.after(0, lambda: os._exit(0))
            except Exception as e:
                erro = str(e)
                self.after(0, lambda: (
                    top.destroy(),
                    messagebox.showerror("Erro na atualização", erro)
                ))

        threading.Thread(target=worker, daemon=True).start()

    # ----- Tutorial -----
    def _abrir_tutorial(self):
        top = tk.Toplevel(self)
        top.title("Tutorial")
        top.geometry("520x420")
        top.transient(self)
        top.configure(bg=self.BG)
        txt = scrolledtext.ScrolledText(top, wrap="word",
                                        font=("Segoe UI", 10),
                                        bg=self.SURFACE, fg=self.TEXT,
                                        insertbackground=self.TEXT,
                                        selectbackground=self.ACCENT,
                                        selectforeground="white",
                                        borderwidth=0,
                                        highlightthickness=0)
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", TUTORIAL_TXT)
        txt.configure(state="disabled")
        ttk.Button(top, text="Fechar", command=top.destroy)\
            .pack(pady=(0, 10))

    # ----- Log -----
    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # ----- Iniciar -----
    def _on_iniciar(self):
        try:
            dados = self._coletar_dados()
        except ValueError as e:
            messagebox.showerror("Dados inválidos", str(e))
            return

        self.btn_iniciar.configure(state="disabled", text="Executando...")
        self._log(f"Iniciando {dados['limite'] - dados['parou']} bordos em "
                  f"{dados['empresa']}...")

        def worker():
            try:
                executar_automacao(
                    dados,
                    log=lambda m: self.after(0, self._log, m),
                )
                self.after(0, lambda: messagebox.showinfo(
                    "Finalizado",
                    f"Foram criados {dados['limite'] - dados['parou']} bordos.\n"
                    f"Empresa: {dados['empresa']}  |  Ticket: {dados['ticket']}"
                ))
            except Exception as e:
                erro = str(e)
                self.after(0, lambda: messagebox.showerror("Erro", erro))
            finally:
                self.after(0, lambda: self.btn_iniciar.configure(
                    state="normal", text="▶  Iniciar automação"))

        threading.Thread(target=worker, daemon=True).start()

    def _coletar_dados(self):
        def req(k, label):
            val = self.vars[k].get().strip()
            if not val:
                raise ValueError(f"Preencha o campo: {label}")
            return val

        usuario = req("usuario", "Usuário") + "@aiko.digital"
        senha = req("senha", "Senha")
        empresa = req("empresa", "Empresa").upper()
        equipamento = req("equipamento", "Equipamento").upper()

        try:
            ticket = int(req("ticket", "Ticket"))
        except ValueError:
            raise ValueError("Ticket deve ser um número")

        zendesk = self.vars["zendesk"].get().strip() or "N"

        try:
            parou = int(self.vars["parou"].get().strip() or "0")
        except ValueError:
            raise ValueError("'Parou no bordo' deve ser um número")

        try:
            limite = int(req("limite", "Qtd. total de bordos"))
        except ValueError:
            raise ValueError("Qtd. total de bordos deve ser um número")

        def lista(k):
            return [x.strip() for x in self.vars[k].get().split(",") if x.strip()]

        return dict(
            usuario=usuario, senha=senha, empresa=empresa,
            equipamento=equipamento, ticket=ticket, zendesk=zendesk,
            parou=parou, limite=limite,
            grupos=lista("grupos"), modelos=lista("modelos"),
            perfil=lista("perfil"),
        )

# ==================== MAIN ====================
if __name__ == "__main__":
    CadastroHUD().mainloop()