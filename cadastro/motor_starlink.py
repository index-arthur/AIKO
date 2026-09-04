"""
Cadastro de Starlink no TracKit.

Um kit Starlink e DOIS registros: um equipamento chamado "<PREFIXO> - <S/N>"
e um device com o proprio S/N, vinculado a ele.

A receita foi levantada dos 41 kits que ja existem em ABE, BRC, CNB, FTM e
VAP - sem excecao de modelo, grupo ou perfil:

    nome do equipamento : "SM - M1HT03247099RPZ"
    modelo              : Starlink
    grupo               : Starlink Aiko
    perfil de rede      : nenhum
    device              : o S/N, type 3 (Globalstar)
    flags               : todas falsas, sem placa e sem capacidade

Prefixo: SM = Starlink Mini, SV = Starlink V4.

Por isso aqui o payload e montado explicitamente, e nao clonado de um molde
como no cadastro comum: a receita e conhecida e igual para todo kit.
"""
from trackit_api_client import TrackitClient, obter_sessao, valor_por_chave

TIPO_GLOBALSTAR = 3
MODELO = "Starlink"
GRUPO = "Starlink Aiko"
PREFIXOS = {"SM": "Starlink Mini", "SV": "Starlink V4"}


def _chave(texto):
    return (texto or "").strip().lower()


def montar_nome(prefixo, serial):
    return "{} - {}".format(prefixo.strip().upper(), serial.strip())


def planejar(equipamentos, mdts, seriais, prefixo):
    """
    Funcao pura: nao toca na rede. Devolve (planos, problemas).

    planos - lista de dicts {serial, nome} na ordem de gravacao
    """
    problemas = []
    seriais = [s.strip() for s in seriais if s and s.strip()]

    if prefixo not in PREFIXOS:
        problemas.append(
            "Prefixo '{}' invalido - use SM (Mini) ou SV (V4).".format(prefixo)
        )
    if not seriais:
        problemas.append("Nenhum S/N informado.")
    if problemas:
        return [], problemas

    # 1. S/N repetido na lista colada
    vistos, repetidos = set(), []
    for s in seriais:
        if _chave(s) in vistos:
            repetidos.append(s)
        vistos.add(_chave(s))
    if repetidos:
        problemas.append(
            "S/N repetido na lista: {}".format(", ".join(sorted(set(repetidos))))
        )

    # 2. equipamento com esse nome ja existe
    nomes = {_chave(e.get("name")) for e in equipamentos}
    ja_equip = [s for s in seriais if _chave(montar_nome(prefixo, s)) in nomes]
    if ja_equip:
        problemas.append(
            "Equipamento ja existe para: {}".format(", ".join(ja_equip[:10]))
        )

    # 3. device com esse S/N ja cadastrado (mesmo com espaco sobrando)
    devices = {_chave(m.get("deviceID")) for m in mdts}
    ja_device = [s for s in seriais if _chave(s) in devices]
    if ja_device:
        problemas.append(
            "Device ja cadastrado para: {}".format(", ".join(ja_device[:10]))
        )

    planos = [{"serial": s.strip(), "nome": montar_nome(prefixo, s)}
              for s in seriais]
    return planos, problemas


def _payload_equipamento(nome, modelo_id, grupo_id):
    """
    A receita do Starlink, copiada campo a campo de um kit ja gravado.

    Duas coisas que o SaveEquipment nao perdoa, ambas descobertas pelo corpo
    da resposta 500 (que traz o motivo em "detail"):

    - equipmentGroupId e OBRIGATORIO na validacao, mesmo o endpoint gravando
      "O Grupo de Equipamento e Obrigatorio" se vier 0. E enviar o grupo
      real aqui JA CRIA a associacao - nao chamar
      EquipmentGroupAssociation/Save depois, sob pena de duplicar.
    - networkProfileId nao deve ser enviado. Starlink nao usa perfil e a
      chave nem existe nos registros gravados; mandar null quebra o binder.
    """
    return {
        "name": nome,
        "equipmentModelId": modelo_id,
        "equipmentGroupId": grupo_id,
        "canCreateTask": False,
        "forceAssignment": False,
        "canPerformQuickMaintenance": False,
        "hasPrecisePosition": False,
        "assignOnlyOnProductiveState": False,
        "tags": [],
        "useHourmeter": False,
        "useOdometer": False,
        "hourmeterRule": 0,
        "odometerRule": 0,
        "isOverrideMeters": False,
        "useHydrometer": False,
        "manuallySetLastPosition": False,
    }


def executar_starlink(dados, log, dry_run=False, progresso=None):
    """dados: empresa, seriais (lista), prefixo (SM/SV), usuario/senha."""
    empresa = dados["empresa"]
    cli = TrackitClient(
        empresa.lower(),
        obter_sessao(
            empresa.lower(),
            usuario=dados.get("usuario"),
            senha=dados.get("senha"),
            log=log,
        ),
    )
    prefixo = (dados.get("prefixo") or "SM").strip().upper()

    log("Lendo cadastros do tenant...")
    equipamentos = cli.equipamentos()
    mdts = cli._get("Forms/MobileDataTerminal/GetAllMobileDataTerminal")

    # modelo e grupo por nome: os IDs mudam de cliente para cliente
    modelo = next(
        (m for m in cli.modelos() if _chave(m.get("name")) == _chave(MODELO)), None)
    grupo = next(
        (g for g in cli.grupos() if _chave(g.get("name")) == _chave(GRUPO)), None)

    faltando = []
    if not modelo:
        faltando.append("modelo '{}'".format(MODELO))
    if not grupo:
        faltando.append("grupo '{}'".format(GRUPO))
    if faltando:
        msg = ("{} nao existe(m) em {}. Crie no TracKit antes - nao vou "
               "escolher outro por conta.".format(
                   " e ".join(faltando), empresa.upper()))
        log(msg, "erro")
        return dict(criados=0, falhas=0, total=0, dry_run=dry_run,
                    problemas=[msg])

    log("Modelo: {} (id {}) | Grupo: {} (id {})".format(
        modelo["name"], modelo["id"], grupo["name"].strip(), grupo["id"]))

    planos, problemas = planejar(
        equipamentos, mdts, dados.get("seriais") or [], prefixo)

    if problemas:
        log("Nao vou gravar - resolva antes:", "erro")
        for p in problemas:
            log("  - " + p, "erro")
        return dict(criados=0, falhas=0, total=len(planos),
                    dry_run=dry_run, problemas=problemas)

    log("{} Starlink {} a cadastrar:".format(
        len(planos), PREFIXOS.get(prefixo, prefixo)))
    for p in planos:
        log("  {}  (device type {})".format(p["nome"], TIPO_GLOBALSTAR))

    if dry_run:
        log("Simulacao: nada foi gravado.")
        return dict(criados=0, falhas=0, total=len(planos),
                    dry_run=True, problemas=[])

    criados = falhas = 0
    for pos, p in enumerate(planos, start=1):
        etiqueta = "{}/{}".format(pos, len(planos))
        try:
            r = cli.salvar_equipamento(
                _payload_equipamento(p["nome"], modelo["id"], grupo["id"]))
            eq_id = r.get("id") if r else None
            if not eq_id:
                raise RuntimeError("equipamento criado sem id na resposta")

            # O SaveEquipment ja cria a associacao de grupo a partir do
            # equipmentGroupId enviado - ele nao ignora o campo, apenas zera
            # a copia guardada no proprio equipamento. Chamar
            # EquipmentGroupAssociation/Save aqui criaria uma SEGUNDA
            # associacao para o mesmo grupo.

            cli.salvar_mdt({
                "deviceID": p["serial"],
                "type": TIPO_GLOBALSTAR,
                "equipmentID": eq_id,
            })

            problemas_item = _conferir(cli, eq_id, p, modelo["id"])
            if problemas_item:
                falhas += 1
                log("[{}] {} -> GRAVOU DIFERENTE".format(etiqueta, p["nome"]),
                    "erro")
                for x in problemas_item:
                    log("        - " + x, "erro")
            else:
                criados += 1
                log("[{}] OK id={} :: {}".format(etiqueta, eq_id, p["nome"]))
        except Exception as e:
            falhas += 1
            log("[{}] FALHOU {} :: {}".format(etiqueta, p["nome"], e), "erro")

        if progresso:
            progresso(pos, len(planos))

    log("Fim. {} Starlink cadastrados, {} com problema.".format(criados, falhas))
    return dict(criados=criados, falhas=falhas, total=len(planos),
                dry_run=False, problemas=[])


def _conferir(cli, eq_id, plano, modelo_id):
    """Le de volta equipamento e device - escrever sem conferir apodrece."""
    problemas = []
    try:
        eq = cli.equipamento(eq_id)
        if (eq.get("name") or "").strip() != plano["nome"]:
            problemas.append("nome: pedi {!r}, gravou {!r}".format(
                plano["nome"], eq.get("name")))
        if eq.get("equipmentModelId") != modelo_id:
            problemas.append("modelo: pedi {}, gravou {}".format(
                modelo_id, eq.get("equipmentModelId")))

        grupos = [valor_por_chave(a, "equipmentGroupID", "equipmentGroupId")
                  for a in cli.associacoes(eq_id)]
        if len(grupos) != 1:
            problemas.append(
                "grupo: esperava 1 associacao, achei {} {}".format(
                    len(grupos), grupos))

        mdts = cli._get("Forms/MobileDataTerminal/GetAllMobileDataTerminal")
        dev = next((m for m in mdts
                    if _chave(m.get("deviceID")) == _chave(plano["serial"])), None)
        if not dev:
            problemas.append("device {} nao apareceu na base".format(
                plano["serial"]))
        elif dev.get("equipmentID") != eq_id:
            problemas.append("device ficou no equipamento {}".format(
                dev.get("equipmentID")))
        elif dev.get("type") != TIPO_GLOBALSTAR:
            problemas.append("device gravou type {}, esperado {}".format(
                dev.get("type"), TIPO_GLOBALSTAR))
    except Exception as e:
        problemas.append("nao consegui verificar: {}".format(e))
    return problemas
