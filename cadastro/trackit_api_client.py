"""
Cliente HTTP para a API interna do TracKit (endpoints /legacy/Forms/*).
Substitui o preenchimento por Selenium do Cadastro de Bordo.

Estrategia: abre o Edge e VOCE loga na mao (SSO/MFA/usuario+senha, tanto faz).
O script nao manipula credencial - so espera a sessao responder JSON, copia os
cookies para uma requests.Session e faz todo o resto por HTTP.

Uso:
    python trackit_api_client.py --empresa vap --inspect
    python trackit_api_client.py --empresa vap --equipamento "ESCAVADEIRA" \
        --ticket 0307 --modelo "Caminhao 3/4" --grupo "VAP FM2C" \
        --perfil "Padrao" --de 1 --ate 10            # dry-run por padrao
    ... --executar                                    # grava de verdade
"""
import argparse
import io
import json
import os
import re
import sys
import time
from datetime import datetime, timezone

from urllib.parse import urlparse

import requests


class TrackitClient:
    def __init__(self, empresa, session):
        self.base = "https://{}.br.trackit.host/legacy/".format(empresa)
        self.s = session

    # ---------- transporte ----------
    def _get(self, path):
        r = self.s.get(self.base + path, timeout=30)
        self._guard_login(r)
        r.raise_for_status()
        return r.json()

    def _post(self, path, payload):
        r = self.s.post(
            self.base + path,
            data=json.dumps(payload),
            headers={
                "Content-Type": "application/json charset=utf-8",
                "X-Requested-With": "XMLHttpRequest",
            },
            timeout=30,
        )
        self._guard_login(r)
        r.raise_for_status()
        return r.json() if r.content else None

    @staticmethod
    def _guard_login(r):
        """HTML no lugar de JSON pode ser 2 coisas bem diferentes."""
        if "text/html" not in r.headers.get("content-type", ""):
            return
        corpo = r.text[:4000]
        # Sessao expirada aparece de duas formas: o formulario ASP.NET nos
        # tenants antigos, ou a tela do Keycloak nos que usam SSO. Sem o
        # segundo caso o erro sairia como "assinatura errada", mandando
        # procurar bug onde so falta relogar.
        host = urlparse(r.url).netloc.lower()
        if (
            'id="formLogin"' in corpo
            or 'name="UserName"' in corpo
            or not host.endswith("trackit.host")
            or "openid-connect" in r.url
            or "kc-form" in corpo
        ):
            raise RuntimeError("Sessao expirada/invalida - refaca o login.")
        raise RuntimeError(
            "Endpoint respondeu HTML ({} {}). Provavel assinatura errada "
            "(nome/tipo de parametro). Trecho: {}".format(
                r.status_code, r.url, " ".join(corpo.split())[:300]
            )
        )

    # ---------- leitura ----------
    def equipamentos(self):
        return self._get("Forms/Equipment/GetAll")

    def modelos(self):
        return self._get("Operation/EquipmentModel/GetAll")

    def perfis_rede(self):
        return self._get("Forms/Equipment/GetAllNetworkProfiles")

    def grupos(self):
        return self._get("Forms/EquipmentGroup/GetAllEquipmentGroups")

    def equipamentos_com_grupo(self):
        return self._get(
            "Forms/Equipment/GetAllEquipmentWithEquipmentGroupAssociation"
        )

    def equipamento(self, eid):
        return self._get(
            "Forms/Equipment/GetEquipmentById?equipmentId={}".format(eid)
        )

    def associacoes(self, equipment_id):
        return self._get(
            "Forms/EquipmentGroupAssociation/GetAllByEquipmentId"
            "?equipmentId={}".format(equipment_id)
        )

    # ---------- escrita ----------
    def salvar_equipamento(self, equipment):
        return self._post("Forms/Equipment/SaveEquipment", equipment)

    def salvar_associacao(self, associacao):
        """SaveEquipment IGNORA equipmentGroupId - o grupo vem por aqui."""
        return self._post(
            "Forms/EquipmentGroupAssociation/Save", associacao
        )

    def salvar_mdt(self, mdt):
        """mdt = {'deviceID': str, 'type': int, 'equipmentID': str|None}"""
        return self._post(
            "Forms/MobileDataTerminal/SaveMobileDataTerminal", mdt
        )


def _arquivo_sessao(empresa):
    return os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        ".sessao_{}.json".format(empresa.lower()),
    )


def _sessao_de_cookies(cookies, user_agent=None):
    s = requests.Session()
    for c in cookies:
        s.cookies.set(c["name"], c["value"], domain=c.get("domain"))
    if user_agent:
        s.headers["User-Agent"] = user_agent
    return s


def _sessao_viva(s, empresa):
    """Unico teste que vale: um endpoint autenticado devolve JSON?"""
    base = "https://{}.br.trackit.host/legacy/".format(empresa)
    try:
        r = s.get(base + "Forms/Equipment/GetAll", timeout=15)
        return r.ok and "application/json" in r.headers.get("content-type", "")
    except requests.RequestException:
        return False


def carregar_sessao(empresa):
    """Reaproveita a sessao do ultimo login, se ainda estiver valida."""
    caminho = _arquivo_sessao(empresa)
    if not os.path.exists(caminho):
        return None
    try:
        with io.open(caminho, encoding="utf-8") as f:
            dados = json.load(f)
    except Exception:
        return None
    s = _sessao_de_cookies(dados.get("cookies", []), dados.get("userAgent"))
    return s if _sessao_viva(s, empresa) else None


def salvar_sessao(empresa, cookies, user_agent, log=print):
    """
    Grava os cookies em disco. ATENCAO: e um token de sessao - quem tiver o
    arquivo entra como voce. Fica local e some quando a sessao expira.
    """
    caminho = _arquivo_sessao(empresa)
    with io.open(caminho, "w", encoding="utf-8") as f:
        f.write(
            json.dumps(
                {"cookies": cookies, "userAgent": user_agent},
                ensure_ascii=False,
            )
        )
    log("Sessao salva (reutilizada ate expirar).")


def tipo_de_login(empresa, timeout=25):
    """
    Descobre como o tenant autentica. Devolve "formulario" ou "sso".

    Nem todo TracKit tem o formulario ASP.NET do /legacy/Login. Alguns
    clientes (CBM, por exemplo) redirecionam para o Keycloak em
    sso.aiko.digital, onde se entra pelos botoes "Entrar com email da
    AIKO/CBM". Nesses nao ha usuario e senha para postar - so o navegador
    completa o fluxo OIDC.
    """
    base = "https://{}.br.trackit.host/legacy/".format(empresa.lower())
    try:
        r = requests.get(base + "Login", timeout=timeout, allow_redirects=True)
    except requests.RequestException:
        return "formulario"  # na duvida, tenta o caminho barato primeiro

    host_final = urlparse(r.url).netloc.lower()
    if not host_final.endswith("trackit.host"):
        return "sso"
    if "__RequestVerificationToken" not in r.text:
        return "sso"
    return "formulario"


def login_por_formulario(empresa, usuario, senha):
    """
    Login direto por HTTP, sem abrir navegador.

    A senha e usada apenas nesta requisicao: nao vai para disco, nao entra no
    arquivo de sessao (que guarda so cookies) e nao aparece em log.

    Levanta RuntimeError se as credenciais forem recusadas OU se a conta
    exigir MFA/SSO - nesses casos o caminho e o login pelo navegador.
    """
    base = "https://{}.br.trackit.host/legacy/".format(empresa.lower())
    s = requests.Session()

    r = s.get(base + "Login", timeout=30)
    r.raise_for_status()
    achado = re.search(
        r'name="__RequestVerificationToken"[^>]*value="([^"]+)"', r.text
    )
    if not achado:
        raise RuntimeError(
            "Nao encontrei o token do formulario de login - a tela de login "
            "pode ter mudado. Use o login pelo navegador."
        )

    resp = s.post(
        base + "Login",
        data={
            "__RequestVerificationToken": achado.group(1),
            "UserName": usuario,
            "Password": senha,
        },
        headers={"Referer": base + "Login"},
        timeout=30,
        allow_redirects=True,
    )

    if _sessao_viva(s, empresa):
        return s

    corpo = resp.text if resp.content else ""
    if "UserName" in corpo or 'id="formLogin"' in corpo:
        raise RuntimeError(
            "Usuario ou senha recusados pelo TracKit.\n"
            "Se a sua conta usa SSO/MFA, use o login pelo navegador."
        )
    raise RuntimeError(
        "Login nao completou (HTTP {}). Se a conta exige MFA, use o login "
        "pelo navegador.".format(resp.status_code)
    )


def obter_sessao(
    empresa, espera_max=300, forcar_login=False, usuario=None, senha=None,
    log=print,
):
    """
    Ordem de tentativa:
      1. sessao salva do login anterior (se ainda viva)
      2. login por formulario, se vierem usuario/senha
      3. login pelo navegador (SSO/MFA)
    """
    if not forcar_login:
        s = carregar_sessao(empresa)
        if s is not None:
            log("Sessao reaproveitada do login anterior.")
            return s

    if usuario and senha:
        if tipo_de_login(empresa) == "sso":
            log("Este cliente entra por SSO (Keycloak), nao por usuario e "
                "senha. Vou abrir o navegador - use o botao 'Entrar com "
                "email da AIKO'.")
            return login_e_pegar_cookies(empresa, espera_max)

        log("Autenticando como {}...".format(usuario))
        s = login_por_formulario(empresa, usuario, senha)
        try:
            salvar_sessao(
                empresa,
                [
                    {"name": c.name, "value": c.value, "domain": c.domain}
                    for c in s.cookies
                ],
                s.headers.get("User-Agent"),
                log=log,
            )
        except Exception as e:
            log("[aviso] nao salvei a sessao: {}".format(e))
        return s

    return login_e_pegar_cookies(empresa, espera_max)


def login_e_pegar_cookies(empresa, espera_max=300):
    """
    Abre o Edge no TracKit e espera VOCE logar na janela (SSO, MFA, usuario/senha
    - tanto faz). O script nao toca em credencial: so fica testando se a sessao
    ja responde JSON e, quando responder, transplanta os cookies.
    """
    from selenium import webdriver

    base = "https://{}.br.trackit.host/legacy/".format(empresa)
    driver = webdriver.Edge()
    try:
        driver.maximize_window()
        driver.get("https://{}.br.trackit.host/".format(empresa))
        print(
            "\n>>> Faca o login na janela do Edge que abriu.\n"
            "    Aguardando a sessao ficar valida (ate {}s)...\n".format(
                espera_max
            )
        )

        limite = time.time() + espera_max
        aviso = 0
        while time.time() < limite:
            cookies = driver.get_cookies()
            try:
                ua = driver.execute_script("return navigator.userAgent")
            except Exception:
                ua = None
            s = _sessao_de_cookies(cookies, ua)

            if _sessao_viva(s, empresa):
                print(">>> Sessao capturada.\n")
                try:
                    salvar_sessao(
                        empresa, cookies, s.headers.get("User-Agent")
                    )
                except Exception as e:
                    print("[aviso] nao salvei a sessao: {}".format(e))
                return s

            if time.time() > aviso:
                restante = int(limite - time.time())
                print("    ...ainda esperando ({}s)".format(restante))
                aviso = time.time() + 15
            time.sleep(3)

        raise RuntimeError(
            "Login nao concluido em {}s - abortando.".format(espera_max)
        )
    finally:
        driver.quit()


def descobrir_molde_associacao(cli, limite=40):
    """
    Acha uma associacao equipamento<->grupo ja existente NESTE tenant para
    usar como molde. Assim o script se adapta a qualquer cliente em vez de
    depender de um formato fixo que eu tenha inferido do bundle.
    """
    try:
        registros = cli.equipamentos_com_grupo()
    except Exception:
        return None

    tentativas = 0
    for e in registros:
        if not e.get("equipmentGroupIds"):
            continue
        tentativas += 1
        if tentativas > limite:
            break
        try:
            assocs = cli.associacoes(e["id"])
        except Exception:
            continue
        if assocs:
            return dict(assocs[0])
    return None


def chave_real(d, *nomes):
    """
    A API do TracKit mistura casing: o equipamento usa 'equipmentGroupId' e a
    associacao usa 'equipmentGroupID'. Acha a chave como ela realmente vem,
    em vez de assumir uma grafia.
    """
    minusculas = {k.lower(): k for k in d}
    for n in nomes:
        if n.lower() in minusculas:
            return minusculas[n.lower()]
    return nomes[0]


def valor_por_chave(d, *nomes):
    return d.get(chave_real(d, *nomes))


def montar_associacao(molde, equipment_id, group_id, inicio):
    """
    Clona o molde do tenant e troca so o que muda.

    Cuidado com casing: se eu adicionar 'equipmentGroupId' sem remover o
    'equipmentGroupID' que veio do molde, o JSON vai com as duas chaves e o
    grupo gravado depende de qual o binder do ASP.NET resolve por ultimo -
    ou seja, o equipamento poderia herdar o grupo do molde.
    """
    base = dict(molde) if molde else {}

    # fora tudo que nao pode ser herdado (inclui os objetos aninhados
    # 'equipment' e 'equipmentGroup', que trazem o cadastro inteiro do molde)
    descartar = {"id", "equipmentgroup", "equipment", "enddate"}
    base = {k: v for k, v in base.items() if k.lower() not in descartar}

    # descobre a grafia usada por ESTE tenant antes de sobrescrever
    k_equip = chave_real(base, "equipmentID", "equipmentId")
    k_grupo = chave_real(base, "equipmentGroupID", "equipmentGroupId")
    k_inicio = chave_real(base, "startDate")

    # remove qualquer variante de caixa das chaves que vamos definir
    alvos = {k_equip.lower(), k_grupo.lower(), k_inicio.lower()}
    base = {k: v for k, v in base.items() if k.lower() not in alvos}

    base[k_equip] = equipment_id
    base[k_grupo] = group_id
    base[k_inicio] = inicio
    return base


def resolver(colecao, termo, rotulo):
    """
    Casa por nome (substring, case-insensitive) OU por ID numerico exato.
    Nomes duplicados existem no tenant, entao ambiguidade nunca e resolvida
    no chute: aborta mostrando os IDs para voce escolher.
    """
    termo = str(termo).strip()

    if termo.isdigit():  # escape hatch: ID exato
        alvo_id = int(termo)
        for x in colecao:
            if x.get("id") == alvo_id:
                return x
        sys.exit("[erro] {} com id {} nao existe.".format(rotulo, termo))

    alvo = termo.lower()
    achados = [x for x in colecao if alvo in (x.get("name") or "").lower()]
    if not achados:
        sys.exit(
            "[erro] {} '{}' nao encontrado. Opcoes: {}".format(
                rotulo, termo, [x.get("name") for x in colecao][:15]
            )
        )
    if len(achados) > 1:
        exatos = [x for x in achados if (x.get("name") or "").lower() == alvo]
        if len(exatos) == 1:
            return exatos[0]
        linhas = "\n".join(
            "    {}  {}".format(x.get("id"), x.get("name")) for x in achados
        )
        sys.exit(
            "[erro] {} '{}' ambiguo - passe o ID em vez do nome:\n{}".format(
                rotulo, termo, linhas
            )
        )
    return achados[0]


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--empresa", required=True)
    p.add_argument(
        "--espera",
        type=int,
        default=300,
        help="segundos para voce concluir o login no navegador",
    )
    p.add_argument(
        "--novo-login",
        dest="novo_login",
        action="store_true",
        help="ignora a sessao salva e abre o navegador de novo",
    )
    p.add_argument(
        "--inspect",
        action="store_true",
        help="so imprime o shape do payload e as opcoes disponiveis",
    )
    p.add_argument(
        "--grupos",
        help="lista grupos que casam com o termo, com quantos equipamentos "
        "cada um tem (para desempatar nomes duplicados)",
    )
    p.add_argument("--equipamento")
    p.add_argument(
        "--prefixo",
        help="prefixo do nome (padrao: --empresa em maiusculo, ex. VAP)",
    )
    p.add_argument("--ticket")
    p.add_argument("--zendesk")
    p.add_argument("--modelo")
    p.add_argument("--grupo")
    p.add_argument("--perfil")
    p.add_argument(
        "--molde",
        help="parte do nome de um equipamento existente cujas configuracoes "
        "serao replicadas (padrao: procura um com 'HWS-' no nome)",
    )
    p.add_argument(
        "--inicio",
        help="data/hora de inicio da associacao de grupo "
        "(ISO, padrao: agora em UTC)",
    )
    p.add_argument("--de", type=int, default=1)
    p.add_argument("--ate", type=int, default=1)
    p.add_argument(
        "--executar", action="store_true", help="sem isso, roda em dry-run"
    )
    p.add_argument(
        "--verificar",
        type=int,
        help="le um equipamento pelo id e imprime o JSON (conferencia)",
    )
    a = p.parse_args()

    cli = TrackitClient(
        a.empresa, obter_sessao(a.empresa, a.espera, a.novo_login)
    )

    equipamentos = cli.equipamentos()
    print("Sessao ok - {} equipamentos visiveis.".format(len(equipamentos)))

    if a.inspect:
        molde = cli.equipamento(equipamentos[0]["id"])
        print("\n=== SHAPE do payload de SaveEquipment ===")
        print(json.dumps(molde, indent=2, ensure_ascii=False)[:3000])
        for rot, fn in (
            ("MODELOS", cli.modelos),
            ("PERFIS DE REDE", cli.perfis_rede),
            ("GRUPOS", cli.grupos),
        ):
            print("\n=== {} ===".format(rot))
            for x in fn()[:20]:
                print("  {}  {}".format(x.get("id"), x.get("name")))
        return

    if a.verificar:
        print(
            json.dumps(
                cli.equipamento(a.verificar), indent=2, ensure_ascii=False
            )
        )
        return

    if a.grupos:
        alvo = a.grupos.lower()
        casando = [
            g
            for g in cli.grupos()
            if alvo in (g.get("name") or "").lower()
        ]
        registros = cli.equipamentos_com_grupo()
        uso = {}
        for e in registros:
            # o endpoint devolve equipmentGroupIds (LISTA) - um equipamento
            # pode ter passado por varios grupos ao longo do tempo
            gids = e.get("equipmentGroupIds")
            if gids is None:
                gid = e.get("equipmentGroupId")
                gids = [gid] if gid else []
            for gid in gids or []:
                if gid:
                    uso[gid] = uso.get(gid, 0) + 1

        if not uso:
            # nao da para distinguir "grupos vazios" de "meu parser falhou"
            print(
                "\n[aviso] li {} registros mas nao extrai nenhum "
                "equipmentGroupId - a contagem abaixo NAO vale.\n"
                "Chaves disponiveis no 1o registro: {}".format(
                    len(registros),
                    sorted(registros[0].keys()) if registros else "(vazio)",
                )
            )

        print("\n=== GRUPOS casando com '{}' ===".format(a.grupos))
        print("(base: {} registros, {} grupos com uso detectado)".format(
            len(registros), len(uso)))
        for g in casando:
            print(
                "  {:>12}  {:>5} equip.  {}".format(
                    g.get("id"), uso.get(g.get("id"), 0), g.get("name")
                )
            )
        return

    modelo = resolver(cli.modelos(), a.modelo, "modelo")
    grupo = resolver(cli.grupos(), a.grupo, "grupo")
    perfil = resolver(cli.perfis_rede(), a.perfil, "perfil de rede")
    print(
        "modelo={} | grupo={} | perfil={}".format(
            modelo["name"], grupo["name"], perfil["name"]
        )
    )

    # Molde: um equipamento JA cadastrado com as configuracoes que voce quer
    # replicar (flags de horimetro, odometro, tarefa, etc). Por padrao procura
    # um do seu padrao de nomenclatura "HWS-".
    termo_molde = a.molde or "hws-"
    cands = [
        e
        for e in equipamentos
        if termo_molde.lower() in (e.get("name") or "").lower()
    ]
    if not cands:
        sys.exit(
            "[erro] nenhum equipamento casa com o molde '{}'. "
            "Use --molde <parte do nome>.".format(termo_molde)
        )
    molde = cli.equipamento(cands[0]["id"])
    print("molde: {}".format(molde.get("name")))

    # campos que NUNCA podem ser herdados de outro equipamento
    for campo in ("id", "equipmentModel", "plate", "capacity"):
        molde.pop(campo, None)

    existentes = set((e.get("name") or "").strip() for e in equipamentos)
    prefixo = a.prefixo or a.empresa.upper()

    # Molde da ASSOCIACAO de grupo. Como SaveEquipment ignora
    # equipmentGroupId, o grupo entra por Forms/EquipmentGroupAssociation.
    # Em vez de adivinhar o formato (que pode variar por tenant), copio uma
    # associacao real deste cliente e troco so o que muda.
    inicio = a.inicio or datetime.now(timezone.utc).replace(
        microsecond=0
    ).isoformat().replace("+00:00", "Z")

    molde_assoc = descobrir_molde_associacao(cli)
    if molde_assoc is None:
        print(
            "[aviso] nenhuma associacao existente encontrada neste tenant - "
            "vou usar o formato minimo {equipmentId, equipmentGroupId, "
            "startDate}."
        )
    else:
        print(
            "molde de associacao: campos {}".format(sorted(molde_assoc.keys()))
        )
    zendesk_txt = (
        " | #{}".format(a.zendesk)
        if a.zendesk and a.zendesk.lower() != "n"
        else ""
    )

    criados = 0
    for k in range(a.de, a.ate + 1):
        nome = "{} | {} | HWS-{}{} | {:02d}".format(
            prefixo, a.equipamento, a.ticket, zendesk_txt, k
        )

        if nome in existentes:  # idempotencia
            print("[{}] JA EXISTE, pulando: {}".format(k, nome))
            continue

        payload = dict(molde)  # ja veio sem id/plate/capacity
        payload.update(
            {
                "name": nome,
                "equipmentModelId": modelo["id"],
                "equipmentGroupId": grupo["id"],
                "networkProfileId": perfil["id"],
                "manuallySetLastPosition": False,
            }
        )

        if not a.executar:
            print("[{}] DRY-RUN {}".format(k, nome))
            if k == a.de:  # mostra os payloads completos do primeiro
                print("  equipamento ->")
                print(json.dumps(payload, indent=2, ensure_ascii=False))
                print("  associacao de grupo ->")
                print(
                    json.dumps(
                        montar_associacao(
                            molde_assoc, "<id-do-novo>", grupo["id"], inicio
                        ),
                        indent=2,
                        ensure_ascii=False,
                    )
                )
            continue

        try:
            r = cli.salvar_equipamento(payload)
            novo_id = r.get("id") if r else None
            criados += 1
            print("[{}] criado id={} :: {}".format(k, novo_id, nome))

            if not novo_id:
                print("      [ATENCAO] sem id na resposta - nao dá p/ associar")
                continue

            # 2o passo: o grupo (SaveEquipment sozinho nao associa)
            try:
                cli.salvar_associacao(
                    montar_associacao(
                        molde_assoc, novo_id, grupo["id"], inicio
                    )
                )
            except Exception as e:
                print("      [ATENCAO] falhou ao associar grupo: {}".format(e))

            # le de volta e confere o que o servidor REALMENTE gravou
            try:
                gravado = cli.equipamento(novo_id)
                divergencias = [
                    "{}: pedi {!r}, gravou {!r}".format(
                        campo, payload.get(campo), gravado.get(campo)
                    )
                    for campo in (
                        "name",
                        "equipmentModelId",
                        "networkProfileId",
                    )
                    if gravado.get(campo) != payload.get(campo)
                ]

                grupos_gravados = [
                    valor_por_chave(x, "equipmentGroupID", "equipmentGroupId")
                    for x in cli.associacoes(novo_id)
                ]
                if grupo["id"] not in grupos_gravados:
                    divergencias.append(
                        "grupo: pedi {}, associacoes gravadas {}".format(
                            grupo["id"], grupos_gravados or "(nenhuma)"
                        )
                    )

                if divergencias:
                    print("      [ATENCAO] o servidor gravou diferente:")
                    for dv in divergencias:
                        print("        - {}".format(dv))
                else:
                    print("      verificado OK (equipamento + grupo)")
            except Exception as e:
                print("      [aviso] nao consegui verificar: {}".format(e))
        except Exception as e:
            print("[{}] FALHOU :: {} :: {}".format(k, nome, e))

    if a.executar:
        print("\nFim. Criados: {}.".format(criados))
    else:
        print("\nDry-run concluido. Rode com --executar para gravar.")


if __name__ == "__main__":
    main()
