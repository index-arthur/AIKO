import os
import sys
import json
import time
import tempfile
import threading
import subprocess
import webbrowser
import urllib.request
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

import sv_ttk

from motor_api import executar_automacao_api, montar_nome
from motor_vinculo import executar_vinculacao
from trackit_api_client import TrackitClient, obter_sessao

# ==================== CONFIG ====================
VERSION = "6.0"
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


TUTORIAL_TXT = (
    "INFORMAÇÕES NECESSÁRIAS:\n\n"
    "• Usuário/Senha  → Seu login do TracKit (sem o @aiko.digital).\n"
    "                   A senha só é usada para autenticar: não vai\n"
    "                   para disco nem para o log. Depois do 1º login\n"
    "                   a sessão fica em cache e os campos podem ficar\n"
    "                   vazios. Conta com MFA/SSO cai no navegador.\n"
    "• Empresa        → Sigla da empresa (ex: BRC, RAI, QA.BRC)\n"
    "• Equipamento    → Tipo (ex: COMODATO, SERVICO DE CAMPO)\n"
    "• Ticket         → Número do ticket (só os números)\n"
    "• Zendesk        → Número do ticket Zendesk (deixe vazio se não tiver)\n"
    "• Parou no bordo → Se parou em algum bordo, o número dele (0 se não parou)\n"
    "• Qtd. Bordos    → Total de bordos a cadastrar\n"
    "\n"
    "COMO USAR:\n\n"
    "1. Preencha Empresa (e usuário/senha, se a sessão expirou)\n"
    "   e clique em Conectar.\n"
    "2. Escolha Modelo, Grupo e Perfil nos seletores — eles listam\n"
    "   os cadastros reais do cliente, com busca por nome ou ID.\n"
    "3. Preencha os dados do lote. A prévia mostra como vai ficar\n"
    "   o nome do primeiro equipamento.\n"
    "4. Clique em Simular para conferir sem gravar nada.\n"
    "5. Clique em Cadastrar para valer.\n"
    "\n"
    "Cada equipamento criado é lido de volta e conferido (nome,\n"
    "modelo, perfil e grupo). Divergência aparece no log em vermelho.\n"
    "\n"
    "ABA VINCULAÇÃO:\n\n"
    "Liga os computadores de bordo aos equipamentos. Informe o filtro\n"
    "(ex: HWS-6848) e cole os seriais, um por linha.\n"
    "\n"
    "O pareamento é POR POSIÇÃO: o 1º serial vai para o equipamento 01,\n"
    "o 2º para o 02, e assim por diante. Por isso as contagens precisam\n"
    "bater — se não baterem, ele para e não grava nada.\n"
    "\n"
    "Device que ainda não existe no TracKit é CADASTRADO e vinculado no\n"
    "mesmo passo — a simulação marca esses com [NOVO].\n"
    "\n"
    "Barra antes de gravar quando: o serial já está vinculado a outro\n"
    "equipamento, está repetido na sua lista, está cadastrado em\n"
    "duplicidade no TracKit, o equipamento já tem device, ou as\n"
    "contagens não batem.\n"
    "\n"
    "IMPORTANTE: Simular nunca grava. Confira o de-para no log e só\n"
    "então clique em Vincular. Como device inexistente vira cadastro\n"
    "novo, um IMEI digitado errado não dá erro — vira um device com o\n"
    "número torto. A conferência na simulação é o que evita isso.\n"
)

# ==================== UPDATE CHECK ====================
def _comparar_versoes(remota, local):
    """Retorna True se remota > local (assumindo formato 'x.y' ou 'x.y.z')."""
    try:
        r = tuple(int(p) for p in remota.split("."))
        l = tuple(int(p) for p in local.split("."))
        return r > l
    except (ValueError, AttributeError):
        return False


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
            "tem_update": _comparar_versoes(tag, VERSION),
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

    # Notas sobre o .bat:
    # - 'ping -n N 127.0.0.1 > nul' espera N-1 segundos SEM criar janela
    #   (diferente do 'timeout' que pisca uma janela cmd.exe).
    # - 'setlocal enabledelayedexpansion' permite usar !var! pra ler valor
    #   atualizado dentro de blocos (necessário pro contador de retries).
    # - Limite de 30 tentativas (~30s) impede loop infinito caso o arquivo
    #   esteja permanentemente lockado.
    bat = (
        f'@echo off\r\n'
        f'setlocal enabledelayedexpansion\r\n'
        f'chcp 65001 > nul\r\n'
        f'ping -n 4 127.0.0.1 > nul\r\n'
        f'set tries=0\r\n'
        f':retry\r\n'
        f'move /y "{exe_novo_path}" "{exe_atual}" > nul 2>&1\r\n'
        f'if not errorlevel 1 goto success\r\n'
        f'set /a tries+=1\r\n'
        f'if !tries! geq 30 goto fail\r\n'
        f'ping -n 2 127.0.0.1 > nul\r\n'
        f'goto retry\r\n'
        f':success\r\n'
        # Espera ~4s depois do move pra dar tempo do Defender terminar
        # o scan do novo arquivo antes de iniciar (evita erro Python DLL).
        f'ping -n 5 127.0.0.1 > nul\r\n'
        f'start "" "{exe_atual}"\r\n'
        f'goto end\r\n'
        f':fail\r\n'
        # Desistiu de substituir — abre o exe original mesmo (versão antiga)
        # pra não deixar o usuário sem nada.
        f'start "" "{exe_atual}"\r\n'
        f':end\r\n'
        f'del "%~f0"\r\n'
    )
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(bat)

    CREATE_NO_WINDOW = 0x08000000
    subprocess.Popen(
        ["cmd", "/c", bat_path],
        creationflags=CREATE_NO_WINDOW,
    )

# ==================== HUD ====================
class CadastroHUD(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Cadastro de Bordo - Aiko  •  v{VERSION}")
        self.geometry("700x860")
        self.minsize(620, 560)
        try:
            self.iconbitmap("negao.ico")
        except Exception:
            pass

        self._aplicar_tema(self._tema_atual)
        self._montar_layout()

        # O minimo vem do que o conteudo realmente precisa. Chutar um valor
        # deixava o formulario colapsar (o corpo virava 1px e so sobravam os
        # botoes). O log e o unico que encolhe: por isso entra com pouco e
        # cresce junto com a janela.
        self.update_idletasks()
        self.minsize(640, min(self.winfo_reqheight(), 920))

        self.after(200, self._verificar_update_async)

    def _aplicar_tema(self, nome):
        """
        O sv_ttk cuida de todos os widgets ttk. Aqui so guardamos a paleta
        para os widgets que NAO sao ttk (o log e o Listbox do seletor), que
        continuam precisando de cor na mao.
        """
        self._tema_atual = nome
        for k, v in self.PALETA[nome].items():
            setattr(self, k, v)

        sv_ttk.set_theme("dark" if nome == "escuro" else "light")
        self._estilo()

        # O set_theme do sv_ttk mexe nos widgets nao-ttk depois desta chamada,
        # sobrescrevendo o que configurarmos agora. Por isso as cores do log
        # sao aplicadas tambem no proximo ciclo ocioso - a ultima palavra
        # precisa ser nossa.
        self._cores_nao_ttk()
        self.after_idle(self._cores_nao_ttk)

        if hasattr(self, "btn_tema"):
            self.btn_tema.configure(
                text="☀ Claro" if nome == "escuro" else "🌙 Escuro"
            )

    def _cores_nao_ttk(self):
        """Log e sufixo do usuario: nao sao ttk, entao levam cor na mao."""
        if hasattr(self, "log"):
            self.log.configure(
                bg=self.SURFACE, fg=self.TEXT,
                insertbackground=self.TEXT,
                selectbackground=self.ACCENT,
            )
            for tag, cor in (("ok", self.OK), ("erro", self.ERRO),
                             ("aviso", self.AVISO), ("info", self.SUBTLE)):
                self.log.tag_configure(tag, foreground=cor)
        if hasattr(self, "txt_seriais"):
            self.txt_seriais.configure(
                bg=self.SURFACE, fg=self.TEXT,
                insertbackground=self.TEXT,
                selectbackground=self.ACCENT,
                highlightbackground=self.BORDER, highlightcolor=self.ACCENT,
            )
        if hasattr(self, "_user_suffix"):
            self._user_suffix.configure(bg=self.SURFACE, fg=self.SUBTLE)

    def _toggle_tema(self):
        self._aplicar_tema("claro" if self._tema_atual == "escuro" else "escuro")

    # Paleta apenas para widgets nao-ttk. Segue as cores do Sun Valley
    # para o log e a lista nao destoarem do resto da janela.
    PALETA = {
        "escuro": {
            "BG": "#1c1c1c", "SURFACE": "#2b2b2b", "BORDER": "#3a3a3a",
            "TEXT": "#ffffff", "SUBTLE": "#9a9a9a", "ACCENT": "#0078d4",
            "OK": "#6ccb5f", "ERRO": "#ff6b6b", "AVISO": "#ffd966",
            "UPDATE_BG": "#3a3500", "UPDATE_FG": "#ffd966",
            "UPDATE_LINK": "#4ea1ff",
        },
        "claro": {
            "BG": "#fafafa", "SURFACE": "#ffffff", "BORDER": "#d6d6d6",
            "TEXT": "#1a1a1a", "SUBTLE": "#5d5d5d", "ACCENT": "#0078d4",
            "OK": "#107c10", "ERRO": "#c42b1c", "AVISO": "#9d5d00",
            "UPDATE_BG": "#fff4ce", "UPDATE_FG": "#6b5300",
            "UPDATE_LINK": "#0067c0",
        },
    }

    _tema_atual = "escuro"

    def _estilo(self):
        """
        Só o que o sv_ttk não cobre: o banner de update e um rótulo
        secundário. Antes eram ~90 linhas configurando cor de cada widget.
        """
        style = ttk.Style()
        style.configure("Sub.TLabel", foreground=self.SUBTLE)
        style.configure("Update.TFrame", background=self.UPDATE_BG)
        style.configure(
            "Update.TLabel", background=self.UPDATE_BG,
            foreground=self.UPDATE_FG,
        )
        style.configure(
            "UpdateLink.TLabel", background=self.UPDATE_BG,
            foreground=self.UPDATE_LINK, font=("Segoe UI", 9, "underline"),
        )
        style.configure("Ok.TLabel", foreground=self.OK)
        style.configure("Erro.TLabel", foreground=self.ERRO)

    def _montar_layout(self):
        # Banner de update (fica escondido ate haver update)
        self.update_bar = ttk.Frame(self, style="Update.TFrame", padding=(12, 7))
        self.update_lbl = ttk.Label(self.update_bar, style="Update.TLabel", text="")
        self.update_link = ttk.Label(self.update_bar, style="UpdateLink.TLabel",
                                     text="Abrir release", cursor="hand2")
        self.update_link.bind("<Button-1>", lambda e: webbrowser.open(RELEASES_URL))
        self.update_lbl.pack(side="left")
        self.update_link.pack(side="right")

        # ---------- Cabecalho ----------
        header = ttk.Frame(self, padding=(20, 18, 20, 6))
        header.pack(fill="x")
        header.columnconfigure(0, weight=1)

        titulo = ttk.Frame(header)
        titulo.grid(row=0, column=0, sticky="w")
        ttk.Label(titulo, text="Cadastro de Bordo",
                  style="Title.TLabel").pack(anchor="w")
        ttk.Label(titulo, text="TracKit / Aiko / v" + VERSION,
                  style="Caption.TLabel").pack(anchor="w")

        self.btn_tema = ttk.Button(header, text="Claro",
                                   command=self._toggle_tema, width=10)
        self.btn_tema.grid(row=0, column=1, sticky="ne")

        self.vars = {}

        def campo(pai, r, rotulo, chave, default=""):
            ttk.Label(pai, text=rotulo).grid(
                row=r, column=0, sticky="w", pady=(0, 2))
            v = tk.StringVar(value=default)
            e = ttk.Entry(pai, textvariable=v)
            e.grid(row=r, column=1, sticky="ew", pady=(0, 8), padx=(12, 0))
            self.vars[chave] = v
            return e

        # O rodape e empacotado ANTES do corpo, com side="bottom": assim ele
        # reserva o espaco dele primeiro e nunca some quando a janela encolhe.
        # (Antes o corpo expandia e empurrava o botao Cadastrar para fora.)
        botoes = ttk.Frame(self, padding=(20, 12, 20, 18))
        botoes.pack(side="bottom", fill="x")
        ttk.Button(botoes, text="Tutorial",
                   command=self._abrir_tutorial).pack(side="left")
        self.btn_iniciar = ttk.Button(botoes, text="Cadastrar",
                                      style="Accent.TButton",
                                      command=self._on_acao)
        self.btn_iniciar.pack(side="right")
        self.btn_simular = ttk.Button(
            botoes, text="Simular",
            command=lambda: self._on_acao(dry_run=True))
        self.btn_simular.pack(side="right", padx=(0, 8))

        # Tambem antes do corpo: quem expande tem de ser empacotado por
        # ultimo, senao consome o espaco de quem vier depois.
        self.progresso = ttk.Progressbar(self, mode="determinate")
        self.progresso.pack(side="bottom", fill="x", padx=20, pady=(0, 6))

        corpo = ttk.Frame(self, padding=(20, 0))
        corpo.pack(fill="both", expand=True)
        corpo.columnconfigure(0, weight=1)

        # ---------- 1. Acesso ----------
        acesso = ttk.Labelframe(corpo, text=" Acesso ", padding=(14, 10, 14, 12))
        acesso.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        acesso.columnconfigure(1, weight=1)

        campo(acesso, 0, "Empresa (sigla)", "empresa")

        ttk.Label(acesso, text="Usuario").grid(row=1, column=0, sticky="w",
                                               pady=(0, 2))
        self.vars["usuario"] = tk.StringVar()
        usr = ttk.Entry(acesso, textvariable=self.vars["usuario"])
        usr.grid(row=1, column=1, sticky="ew", pady=(0, 8), padx=(12, 0))
        self._user_suffix = tk.Label(usr, text="@aiko.digital",
                                     bg=self.SURFACE, fg=self.SUBTLE,
                                     font=("Segoe UI", 9), bd=0)
        self._user_suffix.place(relx=1.0, rely=0.5, anchor="e", x=-8)

        ttk.Label(acesso, text="Senha").grid(row=2, column=0, sticky="w",
                                             pady=(0, 2))
        sen_frame = ttk.Frame(acesso)
        sen_frame.grid(row=2, column=1, sticky="ew", pady=(0, 8), padx=(12, 0))
        sen_frame.columnconfigure(0, weight=1)
        self.vars["senha"] = tk.StringVar()
        sen = ttk.Entry(sen_frame, textvariable=self.vars["senha"], show="*")
        sen.grid(row=0, column=0, sticky="ew")

        def olho():
            visivel = sen.cget("show") == ""
            sen.configure(show="*" if visivel else "")
            self.btn_olho.configure(text="ver" if visivel else "ocultar")

        self.btn_olho = ttk.Button(sen_frame, text="ver", command=olho, width=8)
        self.btn_olho.grid(row=0, column=1, padx=(6, 0))

        conexao = ttk.Frame(acesso)
        conexao.grid(row=3, column=1, sticky="ew", padx=(12, 0), pady=(4, 0))
        self.btn_conectar = ttk.Button(conexao, text="Conectar",
                                       style="Accent.TButton",
                                       command=self._on_conectar)
        self.btn_conectar.pack(side="left")
        self.lbl_conexao = ttk.Label(conexao, text="nao conectado",
                                     style="Sub.TLabel")
        self.lbl_conexao.pack(side="left", padx=(12, 0))

        # ---------- Abas ----------
        self.abas = ttk.Notebook(corpo)
        self.abas.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        self.abas.bind("<<NotebookTabChanged>>", self._ao_trocar_aba)

        aba_cad = ttk.Frame(self.abas, padding=(12, 12))
        aba_vin = ttk.Frame(self.abas, padding=(12, 12))
        self.abas.add(aba_cad, text="  Cadastro de Bordo  ")
        self.abas.add(aba_vin, text="  Vinculacao  ")
        aba_cad.columnconfigure(0, weight=1)
        aba_vin.columnconfigure(0, weight=1)

        # ---------- Aba 1: cadastro ----------
        self.bloco_cad = ttk.Labelframe(
            aba_cad, text=" Cadastros do cliente ", padding=(14, 10, 14, 12))
        self.bloco_cad.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        self.bloco_cad.columnconfigure(1, weight=1)

        self.selecao = {"modelo": None, "grupo": None, "perfil": None}
        self.listas = {"modelo": [], "grupo": [], "perfil": []}
        self.btn_sel = {}

        for r, (chave, rotulo) in enumerate(
            (("modelo", "Modelo"), ("grupo", "Grupo"),
             ("perfil", "Perfil de rede"))
        ):
            ttk.Label(self.bloco_cad, text=rotulo).grid(
                row=r, column=0, sticky="w", pady=(0, 6))
            b = ttk.Button(
                self.bloco_cad, text="Conecte para escolher",
                state="disabled",
                command=lambda k=chave, t=rotulo: self._abrir_seletor(k, t),
            )
            b.grid(row=r, column=1, sticky="ew", pady=(0, 6), padx=(12, 0))
            self.btn_sel[chave] = b

        lote = ttk.Labelframe(aba_cad, text=" Lote ", padding=(14, 10, 14, 12))
        lote.grid(row=1, column=0, sticky="ew")
        lote.columnconfigure(1, weight=1)

        campo(lote, 0, "Equipamento", "equipamento", default="COMODATO")
        campo(lote, 1, "Ticket (numero)", "ticket")
        campo(lote, 2, "Zendesk (opcional)", "zendesk")
        campo(lote, 3, "Parou no bordo", "parou", default="0")
        campo(lote, 4, "Qtd. total de bordos", "limite")

        self.lbl_previa = ttk.Label(lote, text="", style="Sub.TLabel")
        self.lbl_previa.grid(row=5, column=1, sticky="w", padx=(12, 0))
        for k in ("empresa", "equipamento", "ticket", "zendesk"):
            self.vars[k].trace_add("write", lambda *_: self._atualizar_previa())

        # ---------- Aba 2: vinculacao ----------
        bloco_vin = ttk.Labelframe(
            aba_vin, text=" Vincular bordos aos equipamentos ",
            padding=(14, 10, 14, 12))
        bloco_vin.grid(row=0, column=0, sticky="ew")
        bloco_vin.columnconfigure(1, weight=1)

        ttk.Label(bloco_vin, text="Filtro do equipamento").grid(
            row=0, column=0, sticky="w", pady=(0, 2))
        self.vars["filtro"] = tk.StringVar()
        ttk.Entry(bloco_vin, textvariable=self.vars["filtro"]).grid(
            row=0, column=1, sticky="ew", pady=(0, 2), padx=(12, 0))
        ttk.Label(
            bloco_vin,
            text="ex: HWS-6848 - pega os equipamentos do lote, em ordem",
            style="Sub.TLabel",
        ).grid(row=1, column=1, sticky="w", padx=(12, 0), pady=(0, 10))

        ttk.Label(bloco_vin, text="Seriais").grid(
            row=2, column=0, sticky="nw", pady=(0, 2))
        self.txt_seriais = tk.Text(
            bloco_vin, height=8, wrap="none", font=("Consolas", 9),
            bg=self.SURFACE, fg=self.TEXT, insertbackground=self.TEXT,
            selectbackground=self.ACCENT, borderwidth=0,
            highlightthickness=1, relief="flat",
        )
        self.txt_seriais.grid(row=2, column=1, sticky="ew", padx=(12, 0))
        self.txt_seriais.bind("<KeyRelease>", lambda e: self._contar_seriais())

        self.lbl_seriais = ttk.Label(bloco_vin, text="", style="Sub.TLabel")
        self.lbl_seriais.grid(row=3, column=1, sticky="w", padx=(12, 0),
                              pady=(6, 0))
        ttk.Label(
            bloco_vin,
            text="Um serial por linha, na ordem: o 1o vai para o equipamento\n"
                 "01, o 2o para o 02, e assim por diante. As contagens\n"
                 "precisam bater - a confirmacao aparece antes de gravar.",
            style="Sub.TLabel", justify="left",
        ).grid(row=4, column=1, sticky="w", padx=(12, 0), pady=(8, 0))

        # ---------- Log ----------
        log_frame = ttk.Labelframe(corpo, text=" Log ", padding=(10, 8))
        log_frame.grid(row=2, column=0, sticky="nsew")
        corpo.rowconfigure(2, weight=1)
        self.log = scrolledtext.ScrolledText(
            log_frame, height=6, font=("Consolas", 9),
            state="disabled", wrap="word", bg=self.SURFACE, fg=self.TEXT,
            insertbackground=self.TEXT, selectbackground=self.ACCENT,
            borderwidth=0, highlightthickness=0, relief="flat",
        )
        self.log.pack(fill="both", expand=True)
        for tag, cor in (("ok", self.OK), ("erro", self.ERRO),
                         ("aviso", self.AVISO), ("info", self.SUBTLE)):
            self.log.tag_configure(tag, foreground=cor)


        self._atualizar_previa()
        # Reaplica o tema agora que os widgets nao-ttk existem: na 1a chamada
        # (dentro do __init__) o log e a lista ainda nao tinham sido criados.
        self._aplicar_tema(self._tema_atual)

    # ----- Abas -----
    def _aba_atual(self):
        try:
            return self.abas.index(self.abas.select())
        except Exception:
            return 0

    def _ao_trocar_aba(self, _evt=None):
        vinculo = self._aba_atual() == 1
        self.btn_iniciar.configure(text="Vincular" if vinculo else "Cadastrar")

    def _on_acao(self, dry_run=False):
        """O botao principal serve as duas abas."""
        if self._aba_atual() == 1:
            self._on_vincular(dry_run=dry_run)
        else:
            self._on_iniciar(dry_run=dry_run)

    def _seriais_digitados(self):
        bruto = self.txt_seriais.get("1.0", "end")
        return [ln.strip() for ln in bruto.splitlines() if ln.strip()]

    def _contar_seriais(self):
        n = len(self._seriais_digitados())
        self.lbl_seriais.configure(
            text="{} serial(is) na lista".format(n) if n else "")

    def _on_vincular(self, dry_run=False):
        filtro = self.vars["filtro"].get().strip()
        seriais = self._seriais_digitados()
        if not filtro:
            messagebox.showerror(
                "Falta o filtro",
                "Informe o filtro do equipamento (ex: HWS-6848).")
            return
        if not seriais:
            messagebox.showerror("Falta a lista", "Cole os seriais, um por linha.")
            return

        empresa = self.vars["empresa"].get().strip().upper()
        if not empresa:
            messagebox.showerror("Falta a empresa", "Informe a sigla da empresa.")
            return

        if not dry_run and not messagebox.askyesno(
            "Confirmar vinculação",
            f"Isto vai VINCULAR {len(seriais)} bordo(s) em {empresa} — "
            f"de verdade, na produção.\n\n"
            f"Filtro: {filtro}\n\nContinuar?",
        ):
            return

        prefixo = self.vars["usuario"].get().strip()
        if prefixo.lower().endswith("@aiko.digital"):
            prefixo = prefixo[: -len("@aiko.digital")]

        dados = dict(
            empresa=empresa,
            usuario=(prefixo + "@aiko.digital") if prefixo else None,
            senha=self.vars["senha"].get() or None,
            filtro=filtro, seriais=seriais,
        )

        for b in (self.btn_iniciar, self.btn_simular):
            b.configure(state="disabled")
        self.progresso["value"] = 0
        self.progresso["maximum"] = max(len(seriais), 1)
        self._log(("[SIMULAÇÃO] " if dry_run else "")
                  + f"Vinculação: {len(seriais)} bordo(s) em {empresa}...")

        def worker():
            try:
                resumo = executar_vinculacao(
                    dados,
                    log=lambda m, t=None: self.after(0, self._log, m, t),
                    dry_run=dry_run,
                    progresso=lambda f, t: self.after(
                        0, lambda: self.progresso.configure(value=f, maximum=t)
                    ),
                )

                def fim():
                    if resumo["problemas"]:
                        messagebox.showwarning(
                            "Nada foi gravado",
                            "A vinculação foi barrada:\n\n- "
                            + "\n- ".join(resumo["problemas"][:6]),
                        )
                    elif resumo["dry_run"]:
                        messagebox.showinfo(
                            "Simulação",
                            f"{resumo['total']} vínculo(s) prontos.\n"
                            "Confira o de-para no log.",
                        )
                    elif resumo["falhas"]:
                        messagebox.showwarning(
                            "Concluído com problemas",
                            f"Vinculados: {resumo['vinculados']}\n"
                            f"COM PROBLEMA: {resumo['falhas']}\n\n"
                            "Veja o log.",
                        )
                    else:
                        messagebox.showinfo(
                            "Finalizado",
                            f"{resumo['vinculados']} bordo(s) vinculados "
                            f"e conferidos.\n"
                            f"{resumo.get('criados', 0)} foram criados agora.",
                        )
                self.after(0, fim)
            except Exception as e:
                erro = str(e)
                self.after(0, lambda: messagebox.showerror("Erro", erro))
            finally:
                def restaurar():
                    self.btn_iniciar.configure(state="normal")
                    self.btn_simular.configure(state="normal")
                self.after(0, restaurar)

        threading.Thread(target=worker, daemon=True).start()

    # ----- Previa do nome, ao vivo -----
    def _atualizar_previa(self):
        try:
            dados = {
                "empresa": self.vars["empresa"].get().strip().upper() or "?",
                "equipamento":
                    self.vars["equipamento"].get().strip().upper() or "?",
                "ticket": self.vars["ticket"].get().strip() or "?",
                "zendesk": self.vars["zendesk"].get().strip(),
            }
            self.lbl_previa.configure(text="Ficara: " + montar_nome(dados, 1))
        except Exception:
            self.lbl_previa.configure(text="")

    # ----- Conexao -----
    def _on_conectar(self):
        empresa = self.vars["empresa"].get().strip()
        if not empresa:
            messagebox.showerror("Falta a empresa",
                                 "Informe a sigla da empresa primeiro.")
            return

        self.btn_conectar.configure(state="disabled", text="Conectando...")
        self.lbl_conexao.configure(text="autenticando...", style="Sub.TLabel")

        prefixo = self.vars["usuario"].get().strip()
        if prefixo.lower().endswith("@aiko.digital"):
            prefixo = prefixo[: -len("@aiko.digital")]
        usuario = (prefixo + "@aiko.digital") if prefixo else None
        senha = self.vars["senha"].get() or None

        def worker():
            try:
                sessao = obter_sessao(
                    empresa.lower(), usuario=usuario, senha=senha,
                    log=lambda m: self.after(0, self._log, m),
                )
                cli = TrackitClient(empresa.lower(), sessao)
                listas = {
                    "modelo": cli.modelos(),
                    "grupo": cli.grupos(),
                    "perfil": cli.perfis_rede(),
                }
                self.after(0, lambda: self._conectado(empresa, listas))
            except Exception as e:
                erro = str(e)

                def falhou():
                    self.lbl_conexao.configure(text="falhou",
                                               style="Erro.TLabel")
                    self._log("Conexao falhou: " + erro, "erro")
                    messagebox.showerror("Nao conectou", erro)
                self.after(0, falhou)
            finally:
                self.after(0, lambda: self.btn_conectar.configure(
                    state="normal", text="Conectar"))

        threading.Thread(target=worker, daemon=True).start()

    def _conectado(self, empresa, listas):
        self.listas = listas
        self.lbl_conexao.configure(
            text="conectado a " + empresa.upper(), style="Ok.TLabel")
        self._log(
            "Conectado: {} modelos, {} grupos, {} perfis.".format(
                len(listas["modelo"]), len(listas["grupo"]),
                len(listas["perfil"])), "ok")
        for chave, botao in self.btn_sel.items():
            botao.configure(state="normal")
            if self.selecao[chave] is None:
                botao.configure(text="Selecionar...")
        # A senha continua no campo (igual ao usuario) para nao ter de
        # digitar de novo ao reconectar. Fica so em memoria enquanto o app
        # esta aberto: nao vai para disco nem para o arquivo de sessao.

    def _abrir_seletor(self, chave, rotulo):
        itens = self.listas.get(chave) or []
        if not itens:
            messagebox.showinfo("Sem dados", "Conecte primeiro.")
            return
        escolhido = self._dialogo_escolha(rotulo, [], itens)
        if escolhido:
            self.selecao[chave] = escolhido
            self.btn_sel[chave].configure(
                text="{}  (id {})".format(
                    (escolhido.get("name") or "").strip(), escolhido["id"]))

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

        ttk.Button(btns, text="Atualizar agora", style="Accent.TButton",
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
    def _log(self, msg, tag=None):
        if tag is None:
            baixo = msg.lower()
            if "falhou" in baixo or "erro" in baixo or "diferente" in baixo:
                tag = "erro"
            elif msg.strip().startswith("[") and " OK " in msg:
                tag = "ok"
            elif "aviso" in baixo or "ja existe" in baixo:
                tag = "aviso"
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag or ())
        self.log.see("end")
        self.log.configure(state="disabled")

    # ----- Escolha quando os termos casam com mais de um cadastro -----
    def _montar_dialogo_escolha(self, rotulo, candidatos, todos=None):
        """
        Constroi (sem esperar) a janela de escolha. Separado do wait_window
        para dar para testar as cores sem travar num loop modal.

        candidatos: o que casou com os termos digitados (pre-selecionado).
        todos:      a lista COMPLETA do tenant, alcancavel pela busca. Sem
                    isso, quando o termo casa com o cadastro errado nao ha
                    como chegar no certo - foi o caso do grupo "reserva".

        Retorna (janela, dict_resultado).
        """
        todos = todos or candidatos
        win = tk.Toplevel(self)
        win.title("Qual {}?".format(rotulo))
        # O Toplevel NAO herda o tema: sem isto o fundo fica no cinza padrao
        # do Windows e o TLabel (que tem background=BG) vira um bloco preto.
        win.configure(bg=self.BG)
        win.transient(self)
        win.resizable(True, True)
        win.minsize(460, 420)

        corpo = ttk.Frame(win, padding=(14, 12, 14, 12))
        corpo.pack(fill="both", expand=True)

        if len(candidatos) == len(todos):
            cabecalho = "Escolha o {}:".format(rotulo.lower())
        elif candidatos:
            cabecalho = "{} cadastro(s) casaram com os termos.".format(
                len(candidatos)
            )
        else:
            cabecalho = "Nenhum cadastro casou com os termos."
        ttk.Label(corpo, text=cabecalho).pack(anchor="w")
        ttk.Label(
            corpo,
            text="Digite para buscar entre os {} do cliente.".format(len(todos)),
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(0, 8))

        busca_var = tk.StringVar()
        busca = ttk.Entry(corpo, textvariable=busca_var)
        busca.pack(fill="x", pady=(0, 8))

        mostrar_todos = tk.BooleanVar(value=not candidatos)
        ttk.Checkbutton(
            corpo,
            text="Mostrar todos (ignorar os termos digitados)",
            variable=mostrar_todos,
            command=lambda: atualizar(),
        ).pack(anchor="w", pady=(0, 6))

        caixa = tk.Frame(
            corpo, bg=self.BORDER, bd=0, highlightthickness=0
        )
        caixa.pack(fill="both", expand=True)

        lst = tk.Listbox(
            caixa,
            height=16,
            width=64,
            bg=self.SURFACE,
            fg=self.TEXT,
            selectbackground=self.ACCENT,
            selectforeground="#ffffff",
            activestyle="none",
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            exportselection=False,
            font=("Consolas", 9),
        )
        lst.pack(fill="both", expand=True, padx=1, pady=1)

        visiveis = []  # o que esta na lista agora, na mesma ordem

        def atualizar(*_):
            termo = busca_var.get().strip().lower()
            # Digitar significa "quero outro": a busca varre a lista INTEIRA.
            # Se ela ficasse restrita aos candidatos, o caso que motivou isto
            # (termo 'reserva' casando com o grupo errado) continuaria sem
            # saida, porque o grupo certo nao esta entre os candidatos.
            if termo or mostrar_todos.get() or not candidatos:
                base = todos
            else:
                base = candidatos
            if termo:
                base = [
                    x for x in base
                    if termo in (x.get("name") or "").lower()
                    or termo == str(x.get("id"))
                ]
            visiveis[:] = base
            lst.delete(0, "end")
            for c in base:
                lst.insert(
                    "end", " {:>12}  {}".format(c.get("id"), c.get("name"))
                )
            if base:
                lst.selection_set(0)
                lst.activate(0)
            contador.configure(text="{} item(ns)".format(len(base)))

        resultado = {"item": None}

        def confirmar(_evt=None):
            sel = lst.curselection()
            if sel and visiveis:
                resultado["item"] = visiveis[sel[0]]
                win.destroy()

        def cancelar(_evt=None):
            win.destroy()

        lst.bind("<Double-Button-1>", confirmar)
        win.bind("<Return>", confirmar)
        win.bind("<Escape>", cancelar)
        win.protocol("WM_DELETE_WINDOW", cancelar)

        bts = ttk.Frame(corpo)
        bts.pack(fill="x", pady=(12, 0))
        ttk.Button(bts, text="Cancelar", command=cancelar).pack(side="left")
        contador = ttk.Label(bts, text="", style="Sub.TLabel")
        contador.pack(side="left", padx=(10, 0))
        ttk.Button(
            bts, text="Usar este", style="Accent.TButton", command=confirmar
        ).pack(side="right")

        busca_var.trace_add("write", atualizar)
        atualizar()

        # centraliza sobre a janela principal
        win.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - win.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - win.winfo_height()) // 3
        win.geometry("+{}+{}".format(max(x, 0), max(y, 0)))

        lst.focus_set()
        return win, resultado

    def _dialogo_escolha(self, rotulo, candidatos, todos=None):
        """Roda no thread da UI. Devolve o item escolhido ou None."""
        win, resultado = self._montar_dialogo_escolha(rotulo, candidatos, todos)
        win.grab_set()
        win.wait_window()
        return resultado["item"]

    def _escolher_do_worker(self, rotulo, candidatos, todos=None):
        """Ponte thread do motor -> thread da UI (Tk nao e thread-safe)."""
        resposta = {}
        pronto = threading.Event()

        def perguntar():
            try:
                resposta["item"] = self._dialogo_escolha(
                    rotulo, candidatos, todos)
            finally:
                pronto.set()

        self.after(0, perguntar)
        pronto.wait()
        return resposta.get("item")

    # ----- Iniciar -----
    def _on_iniciar(self, dry_run=False):
        try:
            dados = self._coletar_dados()
        except ValueError as e:
            messagebox.showerror("Dados inválidos", str(e))
            return

        total = dados["limite"] - dados["parou"]
        if not dry_run:
            if not messagebox.askyesno(
                "Confirmar cadastro",
                f"Isto vai CRIAR {total} equipamento(s) em "
                f"{dados['empresa']} — de verdade, na produção.\n\n"
                f"Ticket: {dados['ticket']}\n\nContinuar?",
            ):
                return

        for b in (self.btn_iniciar, self.btn_simular):
            b.configure(state="disabled")
        self.btn_iniciar.configure(text="Executando...")
        self.progresso["value"] = 0
        self.progresso["maximum"] = max(total, 1)
        self._log(("[SIMULAÇÃO] " if dry_run else "")
                  + f"{total} equipamento(s) em {dados['empresa']}...")

        def worker():
            try:
                resumo = executar_automacao_api(
                    dados,
                    log=lambda m: self.after(0, self._log, m),
                    dry_run=dry_run,
                    progresso=lambda f, t: self.after(
                        0, lambda: self.progresso.configure(value=f, maximum=t)
                    ),
                    escolher=self._escolher_do_worker,
                )
                def fim():
                    if resumo["dry_run"]:
                        messagebox.showinfo(
                            "Simulação",
                            f"{resumo['total'] - resumo['pulados']} seriam "
                            f"criados.\n{resumo['pulados']} já existem.",
                        )
                    elif resumo["falhas"]:
                        messagebox.showwarning(
                            "Concluído com problemas",
                            f"Criados: {resumo['criados']}\n"
                            f"Já existiam: {resumo['pulados']}\n"
                            f"COM PROBLEMA: {resumo['falhas']}\n\n"
                            "Veja o log — cada falha está detalhada.",
                        )
                    else:
                        messagebox.showinfo(
                            "Finalizado",
                            f"Criados e verificados: {resumo['criados']}\n"
                            f"Já existiam: {resumo['pulados']}",
                        )
                self.after(0, fim)
            except Exception as e:
                erro = str(e)
                self.after(0, lambda: messagebox.showerror("Erro", erro))
            finally:
                def restaurar():
                    self.btn_iniciar.configure(
                        state="normal", text="▶  Iniciar automação")
                    self.btn_simular.configure(state="normal")
                self.after(0, restaurar)

        threading.Thread(target=worker, daemon=True).start()

    def _coletar_dados(self):
        def req(k, label):
            val = self.vars[k].get().strip()
            if not val:
                raise ValueError(f"Preencha o campo: {label}")
            return val

        prefixo = self.vars["usuario"].get().strip()
        if prefixo.lower().endswith("@aiko.digital"):
            prefixo = prefixo[: -len("@aiko.digital")]
        usuario = (prefixo + "@aiko.digital") if prefixo else None
        senha = self.vars["senha"].get() or None

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

        # Modelo/grupo/perfil vem dos seletores, ja resolvidos: mandamos o ID.
        # Assim o motor nunca precisa casar termo com nome - some a classe de
        # erro em que "reserva" pegava o grupo parecido.
        faltando = [
            rot
            for chave, rot in (("modelo", "Modelo"), ("grupo", "Grupo"),
                               ("perfil", "Perfil de rede"))
            if not self.selecao.get(chave)
        ]
        if faltando:
            raise ValueError(
                "Escolha antes: {}.\n(Conecte e clique nos seletores.)".format(
                    ", ".join(faltando)
                )
            )

        return dict(
            usuario=usuario, senha=senha,
            empresa=empresa, equipamento=equipamento, ticket=ticket,
            zendesk=zendesk, parou=parou, limite=limite,
            modelos=[str(self.selecao["modelo"]["id"])],
            grupos=[str(self.selecao["grupo"]["id"])],
            perfil=[str(self.selecao["perfil"]["id"])],
        )

# ==================== MAIN ====================
if __name__ == "__main__":
    CadastroHUD().mainloop()