"""
Registro dos kits Starlink no board Ativos TI do ClickUp.

Cada kit e um card chamado "<PREFIXO>-<NNN>" (SM-123, SV-016) com o S/N no
campo Nº SÉRIE. O casamento entre a lista colada e o board e PELO S/N:

    S/N ja esta em um card   -> atualiza esse card
    S/N nao esta em nenhum   -> cria o proximo numero da sequencia
    S/N em dois ou mais      -> barra o lote (cadastro duplicado no board)

Assim a mesma aba serve para o kit novo (cria) e para o kit que ja existe e
so mudou de plano, cliente ou ativacao (atualiza), sem ninguem ter de saber
qual e o SM de cada serial.

Formato do que o ClickUp devolve, conferido nos cards reais do board:

    drop_down  -> value = orderindex da opcao (int); ao GRAVAR vai o id
    labels     -> value = lista de ids
    short_text -> value = texto
    currency   -> value = texto com o numero ("863")
    date       -> value = texto com milissegundos
"""
import ctypes
import ctypes.wintypes
import json
import os
import re
import time
from datetime import datetime

import requests

API = "https://api.clickup.com/api/v2"
LISTA_ATIVOS = "901327205770"

CAMPO_TIPO = "5cf19fbc-614b-4692-a287-f21a9ad4dfc3"
CAMPO_SERIE = "9a75dc1c-0c0c-447a-8d00-4d2bce08ec29"
CAMPO_PATRIMONIO = "fea60991-a140-476e-abf8-06b4ac7785c3"
CAMPO_TRACKIT = "5e8e3485-608d-4c91-b5a3-87094a02e710"
TRACKIT_SIM = "5228e41e-17e2-4b8f-97bb-0ddb712de444"

TIPO_OPCAO = {
    "SM": "28df6a06-cb7e-4da6-b1b5-a5418b5a4539",   # Starlink Mini - SM
    "SV": "84f6a189-ac89-443d-8e76-458ef34e9fb1",   # Starlink V4 - SV
}

# Campos do lote, na ordem da tela. As opcoes das listas nao ficam aqui:
# vem do proprio ClickUp ao conectar, porque o board ganha cliente novo
# toda semana.
CAMPOS = [
    {"id": "110d2d1d-9766-4a82-bef7-e428f7f39b71", "rotulo": "NF",
     "tipo": "texto"},
    {"id": "c690f873-f3fe-4713-a01b-d2ba29d2791c",
     "rotulo": "Valor de compra (R$)", "tipo": "moeda"},
    {"id": "6077e645-8232-4f7d-b1dc-4b9b4e51f994",
     "rotulo": "Data de aquisição", "tipo": "data"},
    {"id": "9ea6a330-7794-46f9-9fd8-e41e5092b72e", "rotulo": "Condição",
     "tipo": "lista"},
    {"id": CAMPO_TRACKIT, "rotulo": "Cadastro no Trackit", "tipo": "lista"},
    {"id": "d128d6ab-f84e-4833-9c88-7b4ffa6d9faa", "rotulo": "Cliente",
     "tipo": "etiqueta"},
    {"id": "04e1159d-5813-4dab-af12-b06f3f0d1352", "rotulo": "STATUS (campo)",
     "tipo": "lista"},
    {"id": "1bb58a6c-90a3-49e4-867c-36aaca736438", "rotulo": "Plano de dados",
     "tipo": "lista"},
    {"id": "e0bedbd7-0045-418a-8cd9-38f7f7b5f243", "rotulo": "Pagamento",
     "tipo": "lista"},
    {"id": "f338576d-3eaa-4f41-83e7-00a5386842f6",
     "rotulo": "Status ativação plano", "tipo": "lista"},
    {"id": "32884f11-bd64-4602-bae0-ecaee285ae6e",
     "rotulo": "Data de vinculação", "tipo": "data", "hoje": True},
    {"id": "1bac8b5a-28a6-40f2-8ff9-8137b9557d1c", "rotulo": "Nome da rede",
     "tipo": "texto", "padrao": "Coletor Aiko"},
    {"id": "26efe2d6-7ef8-40f7-b623-c8c18701c48f", "rotulo": "Senha da rede",
     "tipo": "texto", "padrao": "Aiko@2024"},
]

PASTA_CONFIG = os.path.join(os.environ.get("APPDATA", "."), "CadastroBordo")
ARQ_TOKEN = os.path.join(PASTA_CONFIG, "clickup_token.dat")
ARQ_ULTIMOS = os.path.join(PASTA_CONFIG, "clickup_ultimos.json")


# ==================== token (DPAPI) ====================
# O token do ClickUp da acesso a tudo que a pessoa ve no workspace. Fica
# criptografado com o DPAPI do Windows: so o mesmo usuario, na mesma
# maquina, consegue ler de volta. Nunca vai para o log.

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
        ctypes.windll.kernel32.LocalFree(ctypes.cast(saida.pbData,
                                                     ctypes.c_void_p))


def ler_token():
    try:
        with open(ARQ_TOKEN, "rb") as f:
            return _dpapi(f.read(), proteger=False).decode("utf-8")
    except (OSError, ValueError):
        return None


def salvar_token(token):
    os.makedirs(PASTA_CONFIG, exist_ok=True)
    with open(ARQ_TOKEN, "wb") as f:
        f.write(_dpapi(token.strip().encode("utf-8"), proteger=True))


def apagar_token():
    try:
        os.remove(ARQ_TOKEN)
    except OSError:
        pass


def ler_ultimos():
    try:
        with open(ARQ_ULTIMOS, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def salvar_ultimos(dados):
    os.makedirs(PASTA_CONFIG, exist_ok=True)
    with open(ARQ_ULTIMOS, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=1)


# ==================== cliente HTTP ====================

class ClickUp:
    def __init__(self, token):
        self.s = requests.Session()
        self.s.headers["Authorization"] = token.strip()

    def _req(self, metodo, caminho, corpo=None):
        for _ in range(6):
            r = self.s.request(metodo, API + caminho, json=corpo, timeout=30)
            if r.status_code == 429:
                # 100 requisicoes/minuto por token. O header diz quando a
                # janela reabre; sem ele, espera o suficiente para reabrir.
                reabre = r.headers.get("X-RateLimit-Reset")
                espera = 20.0
                if reabre and reabre.isdigit():
                    espera = min(max(int(reabre) - time.time(), 1.0), 61.0)
                time.sleep(espera)
                continue
            if r.status_code >= 400:
                raise RuntimeError("ClickUp {} {} -> {} {}".format(
                    metodo, caminho, r.status_code, r.text[:300]))
            return r.json() if r.content else None
        raise RuntimeError("ClickUp segue recusando por excesso de "
                           "requisicoes - espere 1 minuto e tente de novo.")

    def usuario(self):
        return self._req("GET", "/user")["user"]

    def campos(self):
        return self._req("GET", "/list/{}/field".format(LISTA_ATIVOS))["fields"]

    def statuses(self):
        lista = self._req("GET", "/list/{}".format(LISTA_ATIVOS))
        return [s["status"] for s in
                sorted(lista["statuses"], key=lambda s: s.get("orderindex", 0))]

    def cards(self):
        """Todos os cards do board, com o S/N de cada um."""
        cards, pagina = [], 0
        while True:
            r = self._req("GET", "/list/{}/task?page={}&include_closed=true"
                          "&subtasks=true".format(LISTA_ATIVOS, pagina))
            for t in r.get("tasks", []):
                serie = next((c.get("value") for c in t.get("custom_fields", [])
                              if c.get("id") == CAMPO_SERIE), None)
                cards.append({"id": t["id"], "nome": (t.get("name") or "").strip(),
                              "serie": (serie or "").strip(),
                              "status": (t.get("status") or {}).get("status")})
            if r.get("last_page") or not r.get("tasks"):
                return cards
            pagina += 1

    def criar(self, corpo):
        return self._req("POST", "/list/{}/task".format(LISTA_ATIVOS), corpo)

    def mudar_status(self, card_id, status):
        return self._req("PUT", "/task/{}".format(card_id), {"status": status})

    def gravar_campo(self, card_id, campo_id, valor, e_data=False):
        corpo = {"value": valor}
        if e_data:
            corpo["value_options"] = {"time": False}
        return self._req("POST", "/task/{}/field/{}".format(card_id, campo_id),
                         corpo)

    def card(self, card_id):
        return self._req("GET", "/task/{}".format(card_id))


# ==================== valores ====================

def data_em_ms(dia):
    """Meio-dia local: o ClickUp mostra so a data e o fuso nao vira o dia."""
    return int(datetime(dia.year, dia.month, dia.day, 12).timestamp() * 1000)


def ler_data(texto):
    """'25/08/2026' (ou 25/08/26, 25-08-2026) -> datetime. Erro se invalida."""
    t = texto.strip().replace("-", "/").replace(".", "/")
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(t, fmt)
        except ValueError:
            pass
    raise ValueError("data invalida: {!r} (use dd/mm/aaaa)".format(texto))


def ler_moeda(texto):
    """'863', 'R$ 863,00', '1.234,56', '863.50' -> float."""
    t = texto.replace("R$", "").strip()
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        raise ValueError("valor invalido: {!r}".format(texto))


def formatar_nome(prefixo, numero):
    return "{}-{:03d}".format(prefixo, numero)


def ler_linhas(bruto):
    """
    Cada linha: o S/N e, opcionalmente, o patrimonio ao lado (colado do
    Excel vem separado por tab). Devolve [{serial, patrimonio}].
    """
    itens = []
    for ln in bruto.splitlines():
        partes = [p for p in re.split(r"[\s;,]+", ln.strip()) if p]
        if not partes:
            continue
        itens.append({"serial": partes[0],
                      "patrimonio": partes[1] if len(partes) > 1 else ""})
    return itens


# ==================== plano (sem rede) ====================

def _chave(texto):
    return (texto or "").strip().lower()


def ultimo_numero(cards, prefixo):
    padrao = re.compile(r"^{}-(\d+)$".format(re.escape(prefixo)))
    nums = [int(m.group(1)) for c in cards
            for m in [padrao.match(c["nome"])] if m]
    return max(nums, default=0)


def planejar(cards, itens, prefixo):
    """
    Funcao pura. Devolve (acoes, problemas).

    acoes - [{acao: 'criar'|'atualizar', nome, card_id, serial, patrimonio}]
    """
    problemas = []
    if prefixo not in TIPO_OPCAO:
        problemas.append("Prefixo '{}' invalido - use SM ou SV.".format(prefixo))
    if not itens:
        problemas.append("Nenhum S/N informado.")
    if problemas:
        return [], problemas

    vistos, repetidos = set(), []
    for it in itens:
        if _chave(it["serial"]) in vistos:
            repetidos.append(it["serial"])
        vistos.add(_chave(it["serial"]))
    if repetidos:
        problemas.append("S/N repetido na lista: {}".format(
            ", ".join(sorted(set(repetidos)))))

    por_serie = {}
    for c in cards:
        if c["serie"]:
            por_serie.setdefault(_chave(c["serie"]), []).append(c)

    proximo = ultimo_numero(cards, prefixo)
    acoes = []
    for it in itens:
        achados = por_serie.get(_chave(it["serial"]), [])
        if len(achados) > 1:
            problemas.append("S/N {} esta em {} cards: {}".format(
                it["serial"], len(achados),
                ", ".join(c["nome"] for c in achados)))
            continue
        if achados:
            acoes.append(dict(acao="atualizar", nome=achados[0]["nome"],
                              card_id=achados[0]["id"], **it))
        else:
            proximo += 1
            acoes.append(dict(acao="criar", nome=formatar_nome(prefixo, proximo),
                              card_id=None, **it))
    return acoes, problemas


# ==================== execucao ====================

def executar_clickup(dados, log, dry_run=False, progresso=None):
    """
    dados:
      token, prefixo (SM/SV), itens [{serial, patrimonio}]
      campos      [{campo, valor, texto}] - so os marcados na tela
      status_card nome do status do card, ou None para nao mexer
      trackit_sim True quando o TracKit acabou de cadastrar os kits
    """
    cu = ClickUp(dados["token"])
    prefixo = (dados.get("prefixo") or "SM").strip().upper()
    campos = list(dados.get("campos") or [])
    if dados.get("trackit_sim"):
        # o cadastro no TracKit acabou de dar certo: a tela nao manda aqui
        campos = [c for c in campos if c["campo"]["id"] != CAMPO_TRACKIT]
        campos.append({"campo": {"id": CAMPO_TRACKIT, "tipo": "lista",
                                 "rotulo": "Cadastro no Trackit"},
                       "valor": TRACKIT_SIM, "texto": "SIM"})

    log("ClickUp: lendo o board Ativos TI...")
    cards = cu.cards()
    defs = {f["id"]: f for f in cu.campos()}
    acoes, problemas = planejar(cards, dados.get("itens") or [], prefixo)

    if problemas:
        log("ClickUp: nao vou gravar - resolva antes:", "erro")
        for p in problemas:
            log("  - " + p, "erro")
        return dict(criados=0, atualizados=0, falhas=0, total=len(acoes),
                    dry_run=dry_run, problemas=problemas)

    novos = [a for a in acoes if a["acao"] == "criar"]
    log("ClickUp: {} card(s) novo(s), {} a atualizar.".format(
        len(novos), len(acoes) - len(novos)))
    for a in acoes:
        extra = "  patrimonio {}".format(a["patrimonio"]) if a["patrimonio"] else ""
        log("  {:<9} {:<8} S/N {}{}".format(
            "CRIAR" if a["acao"] == "criar" else "ATUALIZAR", a["nome"],
            a["serial"], extra))
    if dados.get("status_card"):
        log("  status do card: {}".format(dados["status_card"]))
    for c in campos:
        log("  {}: {}".format(c["campo"]["rotulo"], c["texto"]))

    if dry_run:
        log("Simulacao ClickUp: nada foi gravado.")
        return dict(criados=0, atualizados=0, falhas=0, total=len(acoes),
                    dry_run=True, problemas=[])

    criados = atualizados = falhas = 0
    for pos, a in enumerate(acoes, start=1):
        etiqueta = "{}/{}".format(pos, len(acoes))
        gravar = list(campos)
        gravar.append({"campo": {"id": CAMPO_SERIE, "tipo": "texto",
                                 "rotulo": "Nº SÉRIE"},
                       "valor": a["serial"], "texto": a["serial"]})
        if a["patrimonio"]:
            gravar.append({"campo": {"id": CAMPO_PATRIMONIO, "tipo": "texto",
                                     "rotulo": "Nº PATRIMÔNIO"},
                           "valor": a["patrimonio"], "texto": a["patrimonio"]})
        try:
            if a["acao"] == "criar":
                corpo = {
                    "name": a["nome"],
                    "custom_fields":
                        [{"id": CAMPO_TIPO, "value": TIPO_OPCAO[prefixo]}]
                        + [{"id": c["campo"]["id"], "value": c["valor"]}
                           for c in gravar],
                }
                if dados.get("status_card"):
                    corpo["status"] = dados["status_card"]
                card_id = cu.criar(corpo)["id"]
            else:
                card_id = a["card_id"]
                if dados.get("status_card"):
                    cu.mudar_status(card_id, dados["status_card"])
                for c in gravar:
                    cu.gravar_campo(card_id, c["campo"]["id"], c["valor"],
                                    e_data=c["campo"]["tipo"] == "data")

            divergencias = _conferir(cu.card(card_id), a["nome"], gravar,
                                     dados.get("status_card"), defs)
            if divergencias:
                falhas += 1
                log("[{}] {} -> GRAVOU DIFERENTE".format(etiqueta, a["nome"]),
                    "erro")
                for d in divergencias:
                    log("        - " + d, "erro")
            else:
                if a["acao"] == "criar":
                    criados += 1
                else:
                    atualizados += 1
                log("[{}] OK {} :: {} (S/N {})".format(
                    etiqueta, "criado" if a["acao"] == "criar" else "atualizado",
                    a["nome"], a["serial"]))
        except Exception as e:
            falhas += 1
            log("[{}] FALHOU {} :: {}".format(etiqueta, a["nome"], e), "erro")
        if progresso:
            progresso(pos, len(acoes))

    log("ClickUp: fim. {} criado(s), {} atualizado(s), {} com problema.".format(
        criados, atualizados, falhas))
    return dict(criados=criados, atualizados=atualizados, falhas=falhas,
                total=len(acoes), dry_run=False, problemas=[])


def _normalizar(valor_lido, campo, defs):
    """Leva o que o ClickUp devolve para o mesmo formato que foi gravado."""
    tipo = campo["tipo"]
    if valor_lido is None:
        return None
    if tipo == "lista":
        # le orderindex, grava id
        opcoes = ((defs.get(campo["id"]) or {}).get("type_config") or {}) \
            .get("options") or []
        return next((o["id"] for o in opcoes
                     if str(o.get("orderindex")) == str(valor_lido)), valor_lido)
    if tipo == "etiqueta":
        return sorted(valor_lido)
    if tipo == "moeda":
        return round(float(valor_lido), 2)
    if tipo == "data":
        return datetime.fromtimestamp(int(valor_lido) / 1000).date()
    return str(valor_lido).strip()


def _conferir(card, nome, gravar, status, defs):
    """Le o card de volta - escrever sem conferir apodrece."""
    problemas = []
    if (card.get("name") or "").strip() != nome:
        problemas.append("nome: pedi {!r}, ficou {!r}".format(
            nome, card.get("name")))
    if status and ((card.get("status") or {}).get("status") or "") != status:
        problemas.append("status do card: pedi {!r}, ficou {!r}".format(
            status, (card.get("status") or {}).get("status")))
    lidos = {c.get("id"): c.get("value") for c in card.get("custom_fields", [])}
    for c in gravar:
        campo = c["campo"]
        esperado = c["valor"]
        if campo["tipo"] == "etiqueta":
            esperado = sorted(esperado)
        elif campo["tipo"] == "moeda":
            esperado = round(float(esperado), 2)
        elif campo["tipo"] == "data":
            esperado = datetime.fromtimestamp(esperado / 1000).date()
        elif campo["tipo"] == "texto":
            esperado = str(esperado).strip()
        try:
            lido = _normalizar(lidos.get(campo["id"]), campo, defs)
        except (TypeError, ValueError):
            lido = lidos.get(campo["id"])
        if lido != esperado:
            problemas.append("{}: pedi {}, ficou {}".format(
                campo["rotulo"], c["texto"], lidos.get(campo["id"])))
    return problemas
