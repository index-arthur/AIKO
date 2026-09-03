"""
Motor de cadastro por API do TracKit - substitui o Selenium na v6.0.

O Selenium sobrevive so no login (uma vez, e a sessao fica em cache). Todo o
cadastro vira POST JSON, o que elimina XPath quebrado, modal de emergencia e
espera de loading.

Contrato mantido igual ao motor antigo:
    executar_automacao_api(dados, log) com as mesmas chaves de `dados`.
"""
import re
from datetime import datetime, timezone

from trackit_api_client import (
    TrackitClient,
    descobrir_molde_associacao,
    montar_associacao,
    obter_sessao,
    valor_por_chave,
)


def _resolver(colecao, termos, rotulo, escolher=None, confirmar=False):
    """
    Recebe a LISTA de termos da HUD (campo "padrao" aceita varios separados
    por virgula) e devolve UM item. Aceita ID numerico exato tambem.

    Nunca desempata no chute: se mais de um casar, aborta mostrando os IDs.
    O tenant da VAP tem grupos com nome identico - escolher em silencio
    colocaria metade do lote no grupo errado.
    """
    if not termos:
        raise ValueError("Nenhum {} informado.".format(rotulo))

    # ID exato tem prioridade
    for t in termos:
        t = str(t).strip()
        if t.isdigit():
            for x in colecao:
                if x.get("id") == int(t):
                    return x
            raise ValueError("{} com id {} nao existe.".format(rotulo, t))

    achados = []
    for x in colecao:
        nome = (x.get("name") or "").lower()
        if any(str(t).strip().lower() in nome for t in termos if str(t).strip()):
            achados.append(x)

    # Nada casou: em vez de beco sem saida, deixa escolher da lista do tenant.
    # (O default 'equipamentos' nao existe na BRC, por exemplo.)
    if not achados:
        if escolher:
            selecionado = escolher(rotulo, [], colecao)
            if selecionado is None:
                raise ValueError("{}: selecao cancelada.".format(rotulo))
            return selecionado
        disponiveis = [x.get("name") for x in colecao][:15]
        raise ValueError(
            "{} nao encontrado para {}.\nDisponiveis (primeiros 15): {}".format(
                rotulo, termos, disponiveis
            )
        )

    # Um unico match nao garante que seja o certo: "reserva" casa com
    # "Reservas_Manutencao" e com qualquer outro grupo parecido. Com
    # confirmar=True o usuario ve o que foi escolhido antes de rodar.
    if len(achados) == 1 and confirmar and escolher:
        selecionado = escolher(rotulo, achados, colecao)
        if selecionado is None:
            raise ValueError("{}: selecao cancelada.".format(rotulo))
        return selecionado

    if len(achados) > 1:
        # tenta match exato antes de desistir
        exatos = [
            x
            for x in achados
            if (x.get("name") or "").strip().lower()
            in [str(t).strip().lower() for t in termos]
        ]
        if len(exatos) == 1:
            return exatos[0]

        # Quem chama pode oferecer uma escolha ao usuario (a HUD faz isso).
        # Sem callback (uso por CLI), aborta - nunca escolhe sozinho.
        if escolher:
            selecionado = escolher(rotulo, achados, colecao)
            if selecionado is None:
                raise ValueError("{}: selecao cancelada.".format(rotulo))
            return selecionado

        linhas = "\n".join(
            "    {}  {}".format(x.get("id"), x.get("name")) for x in achados
        )
        raise ValueError(
            "{} ambiguo - informe o ID em vez do nome:\n{}".format(
                rotulo, linhas
            )
        )

    return achados[0]


def normalizar_ticket(valor):
    """
    Aceita o ticket como ele vem do ClickUp e devolve so o numero.

    'HWS-12312', 'hws-12312', 'HWS 12312', 'hws12312' e '12312' dao todos
    '12312'. O prefixo e reposto na hora de montar o nome, entao guardar so
    o numero evita o "HWS-HWS-12312" de quem cola o ticket inteiro.
    """
    texto = str(valor or "").strip()
    return re.sub(r"^hws[\s\-_]*", "", texto, flags=re.IGNORECASE).strip()


def montar_nome(dados, numero):
    zendesk = str(dados.get("zendesk") or "").strip()
    sufixo = " | #{}".format(zendesk) if zendesk and zendesk.upper() != "N" else ""
    # normaliza de novo aqui: e o unico ponto por onde todo nome passa, e
    # assim nem um chamador distraido consegue produzir "HWS-HWS-...".
    return "{} | {} | HWS-{}{} | {:02d}".format(
        dados["empresa"], dados["equipamento"],
        normalizar_ticket(dados["ticket"]), sufixo, numero
    )


def executar_automacao_api(
    dados, log, dry_run=False, progresso=None, escolher=None,
    confirmar=False,
):
    """
    dry_run=True nao escreve nada: so resolve os IDs e mostra o que faria.
    progresso: callback opcional (feitos, total) para a barra da HUD.
    escolher: callback (rotulo, candidatos) -> item, chamado quando os termos
              casam com mais de um cadastro. Sem ele, ambiguidade e erro.
    """
    empresa = dados["empresa"]
    sessao = obter_sessao(
        empresa.lower(),
        usuario=dados.get("usuario"),
        senha=dados.get("senha"),
        log=log,
    )
    cli = TrackitClient(empresa.lower(), sessao)

    log("Lendo cadastros do tenant...")
    equipamentos = cli.equipamentos()
    log("{} equipamentos na base.".format(len(equipamentos)))

    modelo = _resolver(
        cli.modelos(), dados["modelos"], "Modelo", escolher, confirmar
    )
    grupo = _resolver(
        cli.grupos(), dados["grupos"], "Grupo", escolher, confirmar
    )
    perfil = _resolver(
        cli.perfis_rede(), dados["perfil"], "Perfil de rede", escolher, confirmar
    )
    log(
        "Modelo: {} | Grupo: {} | Perfil: {}".format(
            modelo["name"], grupo["name"], perfil["name"]
        )
    )

    # molde do equipamento: replica as flags de um cadastro ja validado
    termo_molde = dados.get("molde") or "hws-"
    cands = [
        e
        for e in equipamentos
        if termo_molde.lower() in (e.get("name") or "").lower()
    ]
    if not cands:
        raise ValueError(
            "Nenhum equipamento existente casa com '{}' para servir de molde.\n"
            "Sem molde eu teria que chutar as flags (horimetro, odometro, "
            "forceAssignment...). Informe um molde valido.".format(termo_molde)
        )
    molde = cli.equipamento(cands[0]["id"])
    log("Molde: {}".format(molde.get("name")))
    for campo in ("id", "equipmentModel", "plate", "capacity"):
        molde.pop(campo, None)

    molde_assoc = descobrir_molde_associacao(cli)
    if molde_assoc is None:
        log("[aviso] sem associacao de molde neste tenant - usando formato minimo.")

    inicio_iso = (
        datetime.now(timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    )

    existentes = set((e.get("name") or "").strip() for e in equipamentos)
    primeiro = int(dados["parou"]) + 1
    ultimo = int(dados["limite"])
    total = max(0, ultimo - primeiro + 1)

    criados = falhas = pulados = 0

    for pos, k in enumerate(range(primeiro, ultimo + 1), start=1):
        nome = montar_nome(dados, k)

        if nome in existentes:
            pulados += 1
            log("[{}/{}] JA EXISTE, pulando: {}".format(pos, total, nome))
            if progresso:
                progresso(pos, total)
            continue

        payload = dict(molde)
        payload.update(
            {
                "name": nome,
                "equipmentModelId": modelo["id"],
                "equipmentGroupId": grupo["id"],
                "networkProfileId": perfil["id"],
                "manuallySetLastPosition": False,
            }
        )

        if dry_run:
            log("[{}/{}] SIMULACAO: {}".format(pos, total, nome))
            if progresso:
                progresso(pos, total)
            continue

        try:
            r = cli.salvar_equipamento(payload)
            novo_id = r.get("id") if r else None
            if not novo_id:
                raise RuntimeError("resposta sem id")

            # SaveEquipment ignora equipmentGroupId: o grupo e um 2o passo
            cli.salvar_associacao(
                montar_associacao(molde_assoc, novo_id, grupo["id"], inicio_iso)
            )

            problemas = _conferir(cli, novo_id, payload, grupo["id"])
            if problemas:
                falhas += 1
                log("[{}/{}] {} -> GRAVOU DIFERENTE".format(pos, total, nome))
                for p in problemas:
                    log("        - {}".format(p))
            else:
                criados += 1
                log("[{}/{}] OK id={} :: {}".format(pos, total, novo_id, nome))
        except Exception as e:
            falhas += 1
            log("[{}/{}] FALHOU :: {} :: {}".format(pos, total, nome, e))

        if progresso:
            progresso(pos, total)

    resumo = dict(
        criados=criados, falhas=falhas, pulados=pulados, total=total,
        dry_run=dry_run,
    )
    log(
        "Fim. {} criados, {} pulados, {} com problema.".format(
            criados, pulados, falhas
        )
        if not dry_run
        else "Simulacao concluida: {} seriam criados, {} ja existem.".format(
            total - pulados, pulados
        )
    )
    return resumo


def _conferir(cli, novo_id, payload, group_id):
    """Le de volta o que o servidor gravou. Escrever sem conferir apodrece."""
    problemas = []
    try:
        gravado = cli.equipamento(novo_id)
        for campo in ("name", "equipmentModelId", "networkProfileId"):
            if gravado.get(campo) != payload.get(campo):
                problemas.append(
                    "{}: pedi {!r}, gravou {!r}".format(
                        campo, payload.get(campo), gravado.get(campo)
                    )
                )
        # a associacao usa 'equipmentGroupID' (ID maiusculo); o equipamento
        # usa 'equipmentGroupId'. Ler a grafia errada da falso negativo.
        grupos = [
            valor_por_chave(x, "equipmentGroupID", "equipmentGroupId")
            for x in cli.associacoes(novo_id)
        ]
        if group_id not in grupos:
            problemas.append(
                "grupo: pedi {}, gravou {}".format(group_id, grupos or "(nenhum)")
            )
    except Exception as e:
        problemas.append("nao consegui verificar: {}".format(e))
    return problemas
