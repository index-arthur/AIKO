import time
import random
import threading
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
VERSION = "3.5"
REPO_OWNER = "index-arthur"
REPO_NAME = "AIKO"
VERSION_URL = (
    f"https://raw.githubusercontent.com/{REPO_OWNER}/{REPO_NAME}/main/cadastro/VERSION"
)
RELEASES_URL = f"https://github.com/{REPO_OWNER}/{REPO_NAME}/releases/latest"

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
def checar_atualizacao(timeout=3):
    """Retorna (tem_update: bool, versao_remota: str|None)."""
    try:
        req = urllib.request.Request(
            VERSION_URL, headers={"User-Agent": "AikoCadastroHUD"}
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            remota = resp.read().decode("utf-8").strip()
        return (remota != VERSION, remota)
    except Exception:
        return (False, None)

# ==================== SELENIUM HELPERS ====================
def pausa(min_s=0.5, max_s=3.5):
    time.sleep(random.uniform(min_s, max_s))

def clicar_dropdown(driver, xpaths_botao, xpaths_opcoes, condicao):
    wait_curto = WebDriverWait(driver, 3)
    wait_normal = WebDriverWait(driver, 15)

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
    wait_normal = WebDriverWait(driver, 15)
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

def esperar_loading(driver, timeout=15):
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

    def _estilo(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Sub.TLabel", font=("Segoe UI", 9), foreground="#555")
        style.configure("Update.TFrame", background="#FFF3CD")
        style.configure("Update.TLabel", background="#FFF3CD",
                        foreground="#664D03", font=("Segoe UI", 9, "bold"))
        style.configure("UpdateLink.TLabel", background="#FFF3CD",
                        foreground="#0D6EFD",
                        font=("Segoe UI", 9, "underline"))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))

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
                                             state="disabled", wrap="word")
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
            tem, remota = checar_atualizacao()
            if tem and remota:
                self.after(0, lambda: self._mostrar_update(remota))
        threading.Thread(target=worker, daemon=True).start()

    def _mostrar_update(self, remota):
        self.update_lbl.configure(
            text=f"  Nova versão disponível: {remota}  (você tem {VERSION})"
        )
        # Coloca o banner no topo da janela
        self.update_bar.pack(fill="x", before=self.winfo_children()[1])
        if messagebox.askyesno(
            "Atualização disponível",
            f"Há uma nova versão ({remota}).\n"
            f"Você está usando {VERSION}.\n\n"
            "Abrir a página de releases no navegador?",
        ):
            webbrowser.open(RELEASES_URL)

    # ----- Tutorial -----
    def _abrir_tutorial(self):
        top = tk.Toplevel(self)
        top.title("Tutorial")
        top.geometry("520x420")
        top.transient(self)
        txt = scrolledtext.ScrolledText(top, wrap="word",
                                        font=("Segoe UI", 10))
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
