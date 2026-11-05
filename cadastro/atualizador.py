"""
Atualizacao automatica pelo GitHub Releases.

Mora fora da interface porque agora ha duas: a HUD em Tkinter (main.py) e
a casca em WebView2 (app_web.py). Duas copias do updater significariam
duas versoes de VERSION, e a constante tem de bater com a tag da release -
divergiu, ninguem recebe atualizacao e so se descobre semanas depois.

VERSION e a fonte unica. O build.ps1 le daqui e recusa gerar o instalador
se nao casar com o arquivo VERSION.
"""
import json
import os
import subprocess
import sys
import tempfile
import urllib.request

VERSION = "8.0"
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
#
# NUNCA distribuir com True: o app passa a anunciar uma versao 9.9 que nao
# existe e fica pedindo atualizacao para sempre. O build.ps1 recusa buildar
# se isto estiver ligado.
MODO_TESTE_UPDATE = False

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
        # Prefere o instalador. A release pode ter mais de um .exe (o
        # instalador e, por um tempo, o executavel avulso da migracao);
        # pegar "o primeiro .exe" traria o errado dependendo da ordem.
        assets = [a for a in data.get("assets", [])
                  if a.get("name", "").lower().endswith(".exe")]
        exe_asset = next(
            (a for a in assets if "setup" in a.get("name", "").lower()),
            assets[0] if assets else None,
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

def _caminho_curto(caminho):
    """
    Devolve o caminho no formato 8.3 (C:\\Users\\ARTHUR~1\\...).

    Serve para escrever caminhos em .bat sem depender de codepage: o nome
    curto é sempre ASCII. Se o Windows não souber encurtar (volume sem
    suporte a nome curto), devolve o original — melhor tentar do que falhar.
    """
    import ctypes

    buf = ctypes.create_unicode_buffer(600)
    n = ctypes.windll.kernel32.GetShortPathNameW(
        str(caminho), buf, len(buf))
    curto = buf.value if n else ""
    if not curto:
        return caminho
    try:
        curto.encode("ascii")
    except UnicodeEncodeError:
        return caminho
    return curto


def aplicar_atualizacao(instalador_path):
    """Dispara o instalador baixado, em modo silencioso.

    A v6.2 passou a ser distribuída como app instalado, e a atualização
    virou "rodar o instalador novo por cima". Antes o app trocava o próprio
    .exe com um .bat que tentava o `move` 30 vezes até o arquivo destravar —
    frágil e sem jeito de voltar atrás se falhasse no meio.

    O que os parâmetros fazem:
      /SILENT             mostra só a barra de progresso, sem perguntar nada
      /CLOSEAPPLICATIONS  fecha este app para liberar os arquivos
      /NORESTART          nunca reinicia a máquina
      /LOG                deixa rastro em %TEMP% para depurar

    Quem reabre o app é este código, não o instalador. A entrada [Run] com
    `postinstall` do Inno não dispara de forma confiável em modo silencioso
    (testado: o app fechava e não voltava), e "atualizou e a janela sumiu"
    parece defeito para quem está usando. Por isso um `cmd` solto espera o
    instalador terminar e sobe o app de novo — ele sobrevive ao
    /CLOSEAPPLICATIONS porque é outro processo.
    """
    if not getattr(sys, "frozen", False):
        raise RuntimeError(
            "Auto-update só funciona no app instalado.\n"
            "Em modo dev (python main.py), atualize pelo git."
        )

    temp_dir = tempfile.gettempdir()
    log_path = os.path.join(temp_dir, f"aiko_update_{os.getpid()}.log")
    bat_path = os.path.join(temp_dir, f"_aiko_update_{os.getpid()}.bat")

    # Caminho curto (8.3) para tudo que entra no .bat. O perfil do usuario
    # pode ter acento — "C:\Users\ArthurSalomãoSantosC\..." — e o cmd lê
    # .bat na codepage OEM, não em UTF-8: o caminho chega corrompido e o
    # `start` falha em silêncio, sem errorlevel. Testado: o instalador
    # rodava, o app não voltava, e nada indicava o porquê.
    setup_c = _caminho_curto(instalador_path)
    exe_c = _caminho_curto(sys.executable)
    log_c = _caminho_curto(temp_dir) + f"\\aiko_update_{os.getpid()}.log"

    # Um .bat, e não `cmd /c "instalador & start"`: encadear com & exige
    # aspas dentro de aspas, e as regras do cmd comem as externas — nesse
    # caminho o instalador simplesmente não executava.
    bat = (
        "@echo off\r\n"
        f'"{setup_c}" /SILENT /CLOSEAPPLICATIONS /NORESTART /LOG="{log_c}"\r\n'
        # Só reabre se o instalador terminou bem. Se falhou, o app antigo
        # continua instalado e a pessoa reabre pelo atalho de sempre.
        "if errorlevel 1 goto fim\r\n"
        f'start "" "{exe_c}"\r\n'
        ":fim\r\n"
        'del "%~f0"\r\n'
    )
    with open(bat_path, "w", encoding="ascii", errors="strict") as f:
        f.write(bat)

    CREATE_NO_WINDOW = 0x08000000
    subprocess.Popen(["cmd", "/c", bat_path], creationflags=CREATE_NO_WINDOW)
