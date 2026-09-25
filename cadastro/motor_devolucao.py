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
from datetime import datetime, timezone

from trackit_api_client import TrackitClient, obter_sessao

# Quantos dias de silencio fazem um bordo ser considerado "voltou do campo".
#
# O numero saiu dos dados, nao do chute: numa amostra de 150 equipamentos COM
# bordo na ABE, 110 tinham comunicado nas ultimas 24h e apenas 1 caiu na
# faixa de 30 a 90 dias. Ou a maquina falou faz pouco, ou esta muda ha meses
# - 30 dias cai no vale entre os dois mundos, sem separar casos parecidos.
DIAS_SILENCIO = 30


def _chave(serial):
    """Ha serial cadastrado com espaco sobrando - normaliza os dois lados."""
    return (serial or "").strip().lower()


def _dias_desde(iso):
    """Dias desde a data ISO do TracKit. None se nunca comunicou."""
    if not iso:
        return None
    texto = str(iso).strip().strip('"')
    if not texto or texto.lower() == "null":
        return None
    if texto.endswith("Z"):
        texto = texto[:-1] + "+00:00"
    quando = datetime.fromisoformat(texto)
    if quando.tzinfo is None:
        quando = quando.replace(tzinfo=timezone.utc)
    return (datetime.now(timezone.utc) - quando).total_seconds() / 86400.0


def checar_atividade(cli, planos, log=None):
    """
    Descobre, para cada bordo a desvincular, ha quanto tempo o equipamento
    dele falou com o TracKit. Devolve (situacoes, problemas).

    Esta e a trava que importa. Desvincular um bordo em campo faz a maquina
    parar de guardar dados sem ninguem perceber - o operador segue
    trabalhando e o cliente so descobre o buraco depois.

    Tres desfechos, e a diferenca entre os dois ultimos e o ponto:
      silencioso   - mudo ha mais de DIAS_SILENCIO, ou nunca comunicou
      ativo        - falou faz pouco: BARRA
      nao verificado - a pergunta falhou: BARRA tambem. Nao saber se esta em
                       campo tem que pesar igual a saber que esta.
    """
    situacoes, ativos, nao_verificados = [], [], []

    for p in planos:
        # o que importa e ter um equipamento para perguntar - tanto faz se
        # veio de um bordo a desvincular ou de uma linha EQUIP:
        eq_id = p.get("equip_id")
        if not eq_id:
            continue
        rotulo = p["serial"] or "(sem bordo)"
        try:
            dias = _dias_desde(cli.ultima_comunicacao(eq_id))
        except Exception as e:
            nao_verificados.append("{} ({}): {}".format(
                rotulo, p["equipamento"], e))
            situacoes.append(dict(p, dias=None, estado="nao verificado"))
            continue

        estado = "ativo" if dias is not None and dias < DIAS_SILENCIO \
            else "silencioso"
        if estado == "ativo":
            ativos.append("{} em {} (falou ha {:.0f} dia(s))".format(
                rotulo, p["equipamento"], dias))
        situacoes.append(dict(p, dias=dias, estado=estado))
        if log:
            log("  {}  {}  {}".format(
                rotulo,
                "nunca comunicou" if dias is None
                else "ultima comunicacao ha {:.0f} dia(s)".format(dias),
                "<< EM CAMPO" if estado == "ativo" else ""))

    problemas = []
    if ativos:
        problemas.append(
            "Bordo ainda COMUNICANDO (menos de {} dias) - a maquina esta em "
            "campo e pararia de guardar dados:\n    {}\n"
            "  Se a devolucao for real mesmo assim, marque 'Desvincular "
            "mesmo os que ainda comunicam'.".format(
                DIAS_SILENCIO, "\n    ".join(ativos[:10]))
        )
    if nao_verificados:
        problemas.append(
            "Nao consegui checar a ultima comunicacao destes - sem essa "
            "resposta eu nao desvinculo:\n    {}".format(
                "\n    ".join(nao_verificados[:10]))
        )
    return situacoes, problemas


MARCA_EQUIP = "EQUIP:"


def _separar(linhas):
    """
    A lista aceita dois tipos de linha:

        <deviceID>            - o bordo a desvincular
        EQUIP: <nome>         - um equipamento, sem bordo nenhum

    O segundo existe porque equipamento vazio nao aparece na lista de
    bordos: e exatamente o que sobra DEPOIS de uma devolucao, e era
    inalcancavel pela aba. Quem poe essas linhas e o botao da busca, nao a
    digitacao.
    """
    devices, equipamentos = [], []
    for ln in linhas:
        texto = (ln or "").strip()
        if not texto:
            continue
        if texto.upper().startswith(MARCA_EQUIP):
            nome = texto[len(MARCA_EQUIP):].strip()
            if nome:
                equipamentos.append(nome)
        else:
            devices.append(texto)
    return devices, equipamentos


def planejar(equipamentos, mdts, seriais):
    """
    Funcao pura: nao toca na rede. Devolve (planos, problemas).

    planos - lista de dicts na ordem de gravacao, de dois tipos:

      alvo="bordo"        desvincula o bordo do equipamento dele
        mdt         - o registro do bordo como esta na base
        serial      - como o usuario digitou
        equipamento - nome do equipamento de onde ele vai sair
        equip_id    - id desse equipamento
        ja_solto    - True quando nao ha nada a fazer (so registra e pula)

      alvo="equipamento"  nao ha o que desvincular; so serve para excluir
        equipamento, equip_id
    """
    problemas = []
    seriais, nomes_equip = _separar(seriais)

    if not seriais and not nomes_equip:
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
            "alvo": "bordo",
            "mdt": m,
            "serial": s,
            "equip_id": eq_id,
            "equipamento": nomes.get(eq_id, str(eq_id)) if eq_id else "",
            "ja_solto": not eq_id,
        })

    # 4. linhas EQUIP: resolvidas pelo nome exato.
    #    Nome nao e chave no TracKit - a VAP tem equipamentos homonimos.
    #    Escolher um dos dois para EXCLUIR seria imperdoavel, entao aqui
    #    ambiguidade barra em vez de desempatar.
    por_nome = {}
    for e in equipamentos:
        por_nome.setdefault(_chave(e.get("name")), []).append(e)

    ja_na_lista = {p["equip_id"] for p in planos if p["equip_id"]}
    for nome in nomes_equip:
        achados = por_nome.get(_chave(nome), [])
        if not achados:
            problemas.append(
                "Equipamento nao encontrado: {!r}".format(nome))
            continue
        if len(achados) > 1:
            problemas.append(
                "Existe mais de um equipamento chamado {!r} (ids {}) - "
                "renomeie no TracKit antes.".format(
                    nome, ", ".join(str(e["id"]) for e in achados)))
            continue
        e = achados[0]
        if e["id"] in ja_na_lista:
            # ja vai ser tratado pelo bordo que sai dele; repetir criaria
            # duas tentativas de exclusao do mesmo equipamento
            continue
        presos = [m for m in mdts if m.get("equipmentID") == e["id"]]
        if presos:
            problemas.append(
                "{} ainda tem {} bordo(s) vinculado(s) ({}) - use o deviceID "
                "para desvincular primeiro.".format(
                    (e.get("name") or "").strip(), len(presos),
                    ", ".join(str(m.get("deviceID")) for m in presos[:3])))
            continue
        planos.append({
            "alvo": "equipamento",
            "mdt": None,
            "serial": None,
            "equip_id": e["id"],
            "equipamento": (e.get("name") or "").strip(),
            "ja_solto": True,   # nao ha bordo: nada a desvincular
        })

    return planos, problemas


def _pode_excluir(cli, eq_id, nome, log):
    """
    O equipamento so pode sair se ficou realmente vazio.

    Le os bordos de novo em vez de confiar na lista do inicio: o desvinculo
    acabou de mudar a base, e outra pessoa pode ter vinculado algo enquanto
    o lote rodava.
    """
    atuais = cli._get("Forms/MobileDataTerminal/GetAllMobileDataTerminal")
    presos = [m.get("deviceID") for m in atuais
              if m.get("equipmentID") == eq_id]
    if presos:
        log("        nao excluo {}: ainda tem {} bordo(s) - {}".format(
            nome, len(presos), ", ".join(str(p) for p in presos[:5])), "aviso")
        return False
    return True


def executar_devolucao(dados, log, dry_run=False, progresso=None):
    """
    dados: empresa, seriais (lista), usuario/senha, e:
      excluir_equipamento - True para excluir o equipamento que ficar vazio
      forcar_ativos       - True para ignorar a trava de comunicacao
    """
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

    excluir = bool(dados.get("excluir_equipamento"))

    a_fazer = [p for p in planos
               if p["alvo"] == "bordo" and not p["ja_solto"]]
    so_equip = [p for p in planos if p["alvo"] == "equipamento"]
    ja_soltos = len([p for p in planos
                     if p["alvo"] == "bordo" and p["ja_solto"]])

    if so_equip and not excluir:
        # linha EQUIP: so faz sentido para excluir - sem bordo nao ha o que
        # desvincular, e seguir calado nao faria nada e pareceria sucesso
        msg = ("{} equipamento(s) na lista nao tem bordo nenhum: nao ha o "
               "que desvincular. Marque 'Desvincular e EXCLUIR o "
               "equipamento' para remove-los.".format(len(so_equip)))
        log("Nao vou gravar - resolva antes:", "erro")
        log("  - " + msg, "erro")
        return dict(desvinculados=0, ja_soltos=ja_soltos, falhas=0,
                    excluidos=0, total=0, dry_run=dry_run, problemas=[msg])

    log("O que vai sair de onde ({} de {}):".format(len(a_fazer), len(planos)))
    for p in planos:
        if p["alvo"] == "equipamento":
            log("  (sem bordo)  ->  {}".format(p["equipamento"]))
        elif p["ja_solto"]:
            log("  {}  ->  (ja estava sem equipamento)".format(p["serial"]),
                "aviso")
        else:
            log("  {}  ->  sai de {}".format(p["serial"], p["equipamento"]))

    # A trava roda ANTES de qualquer escrita, e tambem na simulacao: quem
    # simula precisa ver o impedimento ali, nao descobrir na hora de gravar.
    #
    # Vale para os EQUIP: tambem. Equipamento sem bordo nao deveria estar
    # comunicando; se estiver, ou a lista de bordos que eu li esta velha, ou
    # alguem acabou de tirar o device de uma maquina em campo. Nos dois
    # casos, parar e o certo.
    checar = a_fazer + so_equip
    if checar:
        log("Checando quem ainda esta em campo...")
        _situacoes, travas = checar_atividade(cli, checar, log)
        if travas and not dados.get("forcar_ativos"):
            log("Nao vou gravar - resolva antes:", "erro")
            for t in travas:
                log("  - " + t, "erro")
            return dict(desvinculados=0, ja_soltos=ja_soltos, falhas=0,
                        excluidos=0, total=len(checar), dry_run=dry_run,
                        problemas=travas)
        if travas and dados.get("forcar_ativos"):
            log("ATENCAO: a trava de comunicacao foi ignorada por escolha "
                "sua. Ha bordo em campo nesta lista.", "aviso")
            for t in travas:
                log("  - " + t, "aviso")

    if excluir:
        log("Os equipamentos abaixo serao EXCLUIDOS (irreversivel):", "aviso")
        for p in a_fazer + so_equip:
            log("  x {}".format(p["equipamento"]), "aviso")

    if dry_run:
        log("Simulacao: nada foi gravado.")
        return dict(desvinculados=0, ja_soltos=ja_soltos, falhas=0,
                    excluidos=0, total=len(a_fazer), dry_run=True,
                    problemas=[])

    desvinculados = falhas = excluidos = 0
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

                # A exclusao so acontece AQUI: depois do desvinculo lido de
                # volta. Se o passo acima falhar, nem se tenta - e a ordem
                # que o TracKit exige e que o estoque ja segue na mao.
                eq_id = m.get("equipmentID")
                if excluir and eq_id:
                    try:
                        if _pode_excluir(cli, eq_id, p["equipamento"], log):
                            cli.excluir_equipamento(eq_id)
                            sobrou = any(
                                e.get("id") == eq_id
                                for e in cli.equipamentos()
                            )
                            if sobrou:
                                falhas += 1
                                log("        NAO EXCLUIU {} - continua na "
                                    "base".format(p["equipamento"]), "erro")
                            else:
                                excluidos += 1
                                log("        equipamento {} excluido".format(
                                    p["equipamento"]))
                    except Exception as e:
                        falhas += 1
                        log("        FALHOU ao excluir {} :: {}".format(
                            p["equipamento"], e), "erro")
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

    # Equipamentos que entraram sozinhos (linha EQUIP:): nao ha desvinculo
    # antes, entao vao direto - mas passam pela mesma conferencia.
    for p in so_equip:
        eq_id, nome = p["equip_id"], p["equipamento"]
        try:
            if not _pode_excluir(cli, eq_id, nome, log):
                falhas += 1
                continue
            cli.excluir_equipamento(eq_id)
            if any(e.get("id") == eq_id for e in cli.equipamentos()):
                falhas += 1
                log("NAO EXCLUIU {} - continua na base".format(nome), "erro")
            else:
                excluidos += 1
                log("equipamento {} excluido".format(nome))
        except Exception as e:
            falhas += 1
            log("FALHOU ao excluir {} :: {}".format(nome, e), "erro")

    log("Fim. {} desvinculados, {} equipamento(s) excluido(s), {} ja estavam "
        "soltos, {} com problema.".format(
            desvinculados, excluidos, ja_soltos, falhas))
    return dict(desvinculados=desvinculados, ja_soltos=ja_soltos,
                falhas=falhas, excluidos=excluidos, total=len(a_fazer),
                dry_run=False, problemas=[])
