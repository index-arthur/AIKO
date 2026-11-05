"""
Guarda segredo no disco cifrado com o DPAPI do Windows.

O DPAPI cifra com a conta do Windows de quem salvou: o arquivo so volta a
ser legivel pelo MESMO usuario, na MESMA maquina. Copiar o .dat para
outro PC nao serve de nada. E o mesmo mecanismo que o gerenciador de
senhas do navegador usa.

O que isso NAO protege: codigo rodando como a propria pessoa. Quem ja esta
dentro do perfil dela consegue pedir ao DPAPI para abrir - igual a
qualquer senha salva em navegador. Por isso a tela deixa a escolha
explicita ("lembrar neste PC") e permite esquecer depois.

Centralizado aqui porque ja eram dois segredos: o token do ClickUp e, a
partir da v8, a senha do TracKit. Duas copias do mesmo ctypes seria uma a
mais do que o necessario.
"""
import ctypes
import ctypes.wintypes
import json
import os

PASTA = os.path.join(os.environ.get("APPDATA", "."), "CadastroBordo")


def _arquivo(chave):
    return os.path.join(PASTA, "{}.dat".format(chave))


class _Blob(ctypes.Structure):
    _fields_ = [("cbData", ctypes.wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_char))]


def _dpapi(dados, proteger):
    buf = ctypes.create_string_buffer(dados, len(dados))
    entrada = _Blob(len(dados), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    saida = _Blob()
    fn = (ctypes.windll.crypt32.CryptProtectData if proteger
          else ctypes.windll.crypt32.CryptUnprotectData)
    if not fn(ctypes.byref(entrada), None, None, None, None, 0,
              ctypes.byref(saida)):
        raise ctypes.WinError()
    try:
        return ctypes.string_at(saida.pbData, saida.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(
            ctypes.cast(saida.pbData, ctypes.c_void_p))


def guardar(chave, texto):
    os.makedirs(PASTA, exist_ok=True)
    with open(_arquivo(chave), "wb") as f:
        f.write(_dpapi(str(texto).encode("utf-8"), proteger=True))


def ler(chave):
    """Devolve o texto, ou None. Arquivo de outro usuario/maquina da None."""
    # Nao ha arquivo, ou o DPAPI recusa o blob (perfil ou maquina
    # diferente). Nenhum dos dois e erro a mostrar: e so "nao tenho senha
    # guardada aqui".
    try:
        with open(_arquivo(chave), "rb") as f:
            return _dpapi(f.read(), proteger=False).decode("utf-8")
    except Exception:
        return None


def apagar(chave):
    try:
        os.remove(_arquivo(chave))
    except OSError:
        pass


def tem(chave):
    return os.path.exists(_arquivo(chave))


# ---------- login do TracKit ----------
# Guardado como um par: trocar de usuario invalida a senha guardada, senao
# a pessoa entraria com o usuario novo e a senha do antigo e levaria um
# "credencial invalida" sem entender por que.
CHAVE_LOGIN = "trackit_login"


def salvar_login(usuario, senha):
    guardar(CHAVE_LOGIN, json.dumps({"usuario": usuario, "senha": senha}))


def ler_login():
    bruto = ler(CHAVE_LOGIN)
    if not bruto:
        return None
    try:
        d = json.loads(bruto)
        return (d.get("usuario") or "", d.get("senha") or "")
    except ValueError:
        return None


def esquecer_login():
    apagar(CHAVE_LOGIN)
