"""
Devolucao de equipamento: desvincula bordos em lote.

Quando o equipamento volta para o estoque, o A9 e retirado e o bordo precisa
ficar solto para ser usado em outro lugar. Na tela do TracKit isso e abrir o
bordo um por um e trocar o Equipamento para "Nao associado" - e para achar
cada bordo alguem digita os 16 caracteres do deviceID a mao. Era so isso que
o estoque pedia para parar de fazer.

O desvinculo e o mesmo caminho do formulario: SaveMobileDataTerminal com
equipmentID nulo. Foi lido do bundle do TracKit, onde a opcao "Nao associado"
e literalmente {id: null, name: ...} jogada no mesmo POST do vinculo.

Nao existe endpoint de exclusao de bordo na API - so Save. Entao desvincular
nunca apaga nada: o registro continua na base, solto, exatamente como os
milhares que ja estao assim (so a VAP tem 1.532).

Desvincular o bordo errado tira um veiculo de producao em silencio. Por isso
planejar() levanta TODAS as pendencias antes de qualquer escrita e o executor
se recusa a gravar enquanto houver alguma - mesma regra da vinculacao.
"""
from trackit_api_client import TrackitClient, obter_sessao


def _chave(serial):
    """Ha serial cadastrado com espaco sobrando - normaliza os dois lados."""
    return (serial or "").strip().lower()


def planejar(equipamentos, mdts, seriais):
    """
    Funcao pura: nao toca na rede. Devolve (planos, problemas).

    planos - lista de dicts na ordem de gravacao:
        mdt    - o registro do bordo como esta na base
        serial - como o usuario digitou
        equipamento - nome do equipamento de onde ele vai sair
        ja_solto    - True quando nao ha nada a fazer (so registra e pula)
    """
    problemas = []
    seriais = [s.strip() for s in seriais if s and s.strip()]

    if not seriais:
        return [], ["Nenhum device informado."]

    # 1. device repetido na lista colada
    vistos, repetidos = set(), []
    for s in seriais:
        if _chave(s) in vistos:
            repetidos.append(s)
        vistos.add(_chave(s))
    if repetidos:
        problemas.append(
            "Device repetido na lista: {}".format(", ".join(sorted(set(repetidos))))
        )

    # 2. indice dos bordos. Duplicata na base so atrapalha se estiver NA LISTA:
    #    ai nao da para saber qual registro desvincular.
    indice, duplicados = {}, set()
    for m in mdts:
        k = _chave(m.get("deviceID"))
        if not k:
            continue
        if k in indice:
            duplicados.add(k)
        indice[k] = m

    ambiguos = sorted({_chave(s) for s in seriais} & duplicados)
    if ambiguos:
        problemas.append(
            "Device cadastrado mais de uma vez no TracKit, nao da para saber "
            "qual desvincular: {} (resolva no TracKit antes)".format(
                ", ".join(ambiguos)
            )
        )

    # 3. device que nao existe. Aqui, ao contrario da vinculacao, NAO se cria
    #    nada: devolucao e de aparelho que ja estava em campo. Device que nao
    #    existe e digitacao errada, e criar um registro novo so sujaria a base.
    inexistentes = [s for s in seriais if _chave(s) not in indice]
    if inexistentes:
        problemas.append(
            "Device nao encontrado no TracKit: {}{}\n"
            "  (confira a digitacao - devolucao nao cria cadastro novo)".format(
                ", ".join(inexistentes[:10]),
                " ..." if len(inexistentes) > 10 else "",
            )
        )

    nomes = {e["id"]: (e.get("name") or "").strip() for e in equipamentos}
    planos = []
    for s in seriais:
        m = indice.get(_chave(s))
        if not m:
            continue
        eq_id = m.get("equipmentID")
        planos.append({
            "mdt": m,
            "serial": s,
            "equipamento": nomes.get(eq_id, str(eq_id)) if eq_id else "",
            "ja_solto": not eq_id,
        })

    return planos, problemas


def executar_devolucao(dados, log, dry_run=False, progresso=None):
    """dados: empresa, seriais (lista), e opcionalmente usuario/senha."""
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

    log("Lendo equipamentos e bordos...")
    equipamentos = cli.equipamentos()
    mdts = cli._get("Forms/MobileDataTerminal/GetAllMobileDataTerminal")
    log("{} equipamentos, {} bordos.".format(len(equipamentos), len(mdts)))

    planos, problemas = planejar(equipamentos, mdts, dados.get("seriais") or [])

    if problemas:
        log("Nao vou gravar - resolva antes:", "erro")
        for p in problemas:
            log("  - " + p, "erro")
        return dict(desvinculados=0, ja_soltos=0, falhas=0, total=len(planos),
                    dry_run=dry_run, problemas=problemas)

    a_fazer = [p for p in planos if not p["ja_solto"]]
    ja_soltos = len(planos) - len(a_fazer)

    log("O que vai sair de onde ({} de {}):".format(len(a_fazer), len(planos)))
    for p in planos:
        if p["ja_solto"]:
            log("  {}  ->  (ja estava sem equipamento)".format(p["serial"]),
                "aviso")
        else:
            log("  {}  ->  sai de {}".format(p["serial"], p["equipamento"]))

    if dry_run:
        log("Simulacao: nada foi gravado.")
        return dict(desvinculados=0, ja_soltos=ja_soltos, falhas=0,
                    total=len(a_fazer), dry_run=True, problemas=[])

    desvinculados = falhas = 0
    for pos, p in enumerate(a_fazer, start=1):
        etiqueta = "{}/{}".format(pos, len(a_fazer))
        m, serial = p["mdt"], p["serial"]
        try:
            # equipmentID nulo e o "Nao associado" da tela. Mandar 0 nao serve:
            # 0 e um id, nulo e a ausencia dele.
            cli.salvar_mdt({
                "id": m["id"],
                "deviceID": m.get("deviceID"),
                "type": m.get("type", 0),
                "equipmentID": None,
            })

            # le de volta: escrever sem conferir e como nao ter escrito
            atuais = cli._get(
                "Forms/MobileDataTerminal/GetAllMobileDataTerminal")
            atual = next((x for x in atuais if x.get("id") == m["id"]), None)
            if atual is not None and not atual.get("equipmentID"):
                desvinculados += 1
                log("[{}] OK {} saiu de {}".format(
                    etiqueta, serial, p["equipamento"]))
            else:
                falhas += 1
                log("[{}] AINDA VINCULADO {} -> equipmentID {}".format(
                    etiqueta, serial,
                    atual.get("equipmentID") if atual else "(sumiu da base)"),
                    "erro")
        except Exception as e:
            falhas += 1
            log("[{}] FALHOU {} :: {}".format(etiqueta, serial, e), "erro")

        if progresso:
            progresso(pos, len(a_fazer))

    log("Fim. {} desvinculados, {} ja estavam soltos, {} com problema.".format(
        desvinculados, ja_soltos, falhas))
    return dict(desvinculados=desvinculados, ja_soltos=ja_soltos, falhas=falhas,
                total=len(a_fazer), dry_run=False, problemas=[])
