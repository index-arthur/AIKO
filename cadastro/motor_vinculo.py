"""
Vinculacao de computadores de bordo (MDT) aos equipamentos.

O vinculo e o campo equipmentID no proprio bordo: basta reenviar o registro
por Forms/MobileDataTerminal/SaveMobileDataTerminal com o equipmentID
preenchido.

O pareamento e POSICIONAL: o 1o serial da lista vai para o 1o equipamento
que casa com o filtro (ordenado por nome), o 2o para o 2o, e assim por
diante. Como os nomes terminam em 01, 02, 03..., a ordem alfabetica e a
ordem certa.

Vincular o serial errado e pior do que nao vincular: o equipamento passa a
reportar dados de outra maquina e ninguem percebe. Por isso montar_pares
levanta TODAS as pendencias antes de qualquer escrita, e o executor se
recusa a gravar enquanto houver alguma.
"""
from trackit_api_client import TrackitClient, obter_sessao

TIPO_MDT = 0  # o tenant so usa MDT; os outros tipos existem mas nao sao usados


def _chave(serial):
    """Ha serial cadastrado com espaco sobrando - normaliza os dois lados."""
    return (serial or "").strip().lower()


def montar_pares(equipamentos, mdts, filtro, seriais, criar_faltantes=True):
    """
    Funcao pura: nao toca na rede. Devolve (pares, problemas).

    pares     - lista de (mdt_ou_None, equipamento, serial) na ordem de
                gravacao. mdt None significa "bordo ainda nao existe".
    problemas - lista de strings; se houver qualquer uma, nao se grava nada

    criar_faltantes=True (padrao): serial que ainda nao existe vira cadastro
    novo. E o fluxo da Aiko - o device chega novo e nunca esta pre-cadastrado;
    quando volta ao estoque ele e desvinculado, mas continua na base.

    O preco disso: um IMEI digitado errado nao da erro, vira um device novo
    com o numero torto (a base ja tem casos assim, ex: 3544'891242'21' com um
    digito a menos). A defesa e a simulacao, que marca [NOVO] linha a linha
    para conferencia antes de gravar.

    criar_faltantes=False: serial inexistente e erro. Util para revincular
    devices que voltaram do estoque, quando todos deveriam existir.
    """
    problemas = []

    alvo = (filtro or "").strip().lower()
    if not alvo:
        return [], ["Informe o filtro do equipamento (ex: HWS-6848)."]

    equipamentos_alvo = sorted(
        (e for e in equipamentos if alvo in (e.get("name") or "").lower()),
        key=lambda e: (e.get("name") or "").strip(),
    )
    seriais = [s.strip() for s in seriais if s and s.strip()]

    if not equipamentos_alvo:
        problemas.append(
            "Nenhum equipamento casa com '{}'.".format(filtro)
        )
    if not seriais:
        problemas.append("Nenhum serial informado.")
    if problemas:
        return [], problemas

    # 1. serial repetido na lista colada
    vistos, repetidos = set(), []
    for s in seriais:
        if _chave(s) in vistos:
            repetidos.append(s)
        vistos.add(_chave(s))
    if repetidos:
        problemas.append(
            "Serial repetido na lista: {}".format(", ".join(sorted(set(repetidos))))
        )

    # 2. indice dos bordos existentes.
    #    A base tem seriais cadastrados em duplicidade (na VAP sao 15). Isso
    #    so importa se um deles estiver NA SUA LISTA: ai nao da para saber
    #    qual registro vincular, e vincular o errado seria pior que parar.
    #    Duplicata que nao entra no lote nao e problema nosso.
    indice, duplicados_base = {}, set()
    for m in mdts:
        k = _chave(m.get("deviceID"))
        if not k:
            continue
        if k in indice:
            duplicados_base.add(k)
        indice[k] = m

    ambiguos = sorted({_chave(s) for s in seriais} & duplicados_base)
    if ambiguos:
        problemas.append(
            "Serial cadastrado mais de uma vez no TracKit, nao da para saber "
            "qual vincular: {} (resolva no TracKit antes)".format(
                ", ".join(ambiguos)
            )
        )

    # 3. serial que nao existe: erro, ou candidato a cadastro novo
    inexistentes = [s for s in seriais if _chave(s) not in indice]
    if inexistentes and not criar_faltantes:
        problemas.append(
            "Serial nao encontrado no TracKit: {}{}\n"
            "  (se sao bordos novos, marque 'Criar bordos que nao existem')"
            .format(
                ", ".join(inexistentes[:10]),
                " ..." if len(inexistentes) > 10 else "",
            )
        )

    # 4. serial que ja esta vinculado a algum equipamento
    nomes = {e["id"]: (e.get("name") or "").strip() for e in equipamentos}
    ja_vinculados = []
    for s in seriais:
        m = indice.get(_chave(s))
        if m and m.get("equipmentID"):
            ja_vinculados.append(
                "{} -> {}".format(
                    s, nomes.get(m["equipmentID"], m["equipmentID"])
                )
            )
    if ja_vinculados:
        problemas.append(
            "Serial ja vinculado: {}".format("; ".join(ja_vinculados[:10]))
        )

    # 5. equipamento que ja tem bordo
    ocupados = {m.get("equipmentID") for m in mdts if m.get("equipmentID")}
    eq_ocupados = [
        (e.get("name") or "").strip()
        for e in equipamentos_alvo
        if e["id"] in ocupados
    ]
    if eq_ocupados:
        problemas.append(
            "Equipamento ja tem bordo: {}".format("; ".join(eq_ocupados[:10]))
        )

    # 6. contagens diferentes - o pareamento posicional sairia torto
    if len(seriais) != len(equipamentos_alvo):
        problemas.append(
            "{} serial(is) para {} equipamento(s) que casam com '{}'. "
            "O pareamento e por posicao, entao as contagens precisam bater."
            .format(len(seriais), len(equipamentos_alvo), filtro)
        )

    # o serial vai junto no par: quando o bordo nao existe ainda, e dele que
    # sai o deviceID do cadastro novo (com os apostrofos, exatamente como
    # veio - o padrao da casa e esse e nao cabe a mim normalizar)
    pares = [
        (indice.get(_chave(s)), e, s)
        for s, e in zip(seriais, equipamentos_alvo)
    ]
    return pares, problemas


def executar_vinculacao(dados, log, dry_run=False, progresso=None):
    """
    dados: empresa, filtro, seriais (lista), e opcionalmente usuario/senha.
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

    criar = dados.get("criar_faltantes", True)
    pares, problemas = montar_pares(
        equipamentos, mdts, dados.get("filtro"), dados.get("seriais") or [],
        criar_faltantes=criar,
    )

    if problemas:
        log("Nao vou gravar - resolva antes:", "erro")
        for p in problemas:
            log("  - " + p, "erro")
        return dict(vinculados=0, criados=0, falhas=0, total=len(pares),
                    dry_run=dry_run, problemas=problemas)

    novos = sum(1 for mdt, _, _ in pares if mdt is None)
    log("De-para ({} vinculo(s), {} bordo(s) a criar):".format(
        len(pares), novos))
    for mdt, eq, serial in pares:
        log("  {} {}  ->  {}".format(
            "[NOVO]" if mdt is None else "      ",
            serial, (eq.get("name") or "").strip()))

    if dry_run:
        log("Simulacao: nada foi gravado.")
        return dict(vinculados=0, criados=0, falhas=0, total=len(pares),
                    dry_run=True, problemas=[])

    vinculados = criados = falhas = 0
    for pos, (mdt, eq, serial) in enumerate(pares, start=1):
        etiqueta = "{}/{}".format(pos, len(pares))
        try:
            if mdt is None:
                # cadastro novo: sem id, o TracKit cria (mesmo caminho do
                # botao "+" da tela de Computador de bordo)
                cli.salvar_mdt({
                    "deviceID": serial,
                    "type": TIPO_MDT,
                    "equipmentID": eq["id"],
                })
                acao = "CRIADO+VINCULADO"
            else:
                cli.salvar_mdt({
                    "id": mdt["id"],
                    "deviceID": mdt.get("deviceID"),
                    "type": mdt.get("type", TIPO_MDT),
                    "equipmentID": eq["id"],
                })
                acao = "OK"

            # le de volta: escrever sem conferir e como nao ter escrito
            atuais = cli._get(
                "Forms/MobileDataTerminal/GetAllMobileDataTerminal")
            atual = next(
                (m for m in atuais
                 if _chave(m.get("deviceID")) == _chave(serial)),
                None,
            )
            if atual and atual.get("equipmentID") == eq["id"]:
                if mdt is None:
                    criados += 1
                vinculados += 1
                log("[{}] {} {} -> {}".format(
                    etiqueta, acao, serial, (eq.get("name") or "").strip()))
            else:
                falhas += 1
                log("[{}] GRAVOU DIFERENTE {} -> equipmentID {}".format(
                    etiqueta, serial,
                    atual.get("equipmentID") if atual else "(nao encontrei)"),
                    "erro")
        except Exception as e:
            falhas += 1
            log("[{}] FALHOU {} :: {}".format(etiqueta, serial, e), "erro")

        if progresso:
            progresso(pos, len(pares))

    log("Fim. {} vinculados ({} criados), {} com problema.".format(
        vinculados, criados, falhas))
    return dict(vinculados=vinculados, criados=criados, falhas=falhas,
                total=len(pares), dry_run=False, problemas=[])
