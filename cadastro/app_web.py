"""
Cadastro de Bordo - Aiko : casca em WebView2 (pywebview).

Substitui a HUD em Tkinter. A interface vira HTML/CSS dentro do WebView2,
que e o proprio motor do Edge - por isso a rolagem e o redimensionamento
ficam suaves: quem compoe e a GPU, nao o Tk reposicionando 121 widgets a
cada pixel.

O QUE NAO MUDA: nada de regra de negocio. Este arquivo so traduz o que a
tela mandou para a chamada do motor correspondente, e devolve o log. Os
motores (motor_api, motor_vinculo, motor_starlink, motor_devolucao,
motor_clickup) e o trackit_api_client continuam identicos - 2.609 linhas
que nunca souberam que existia interface, e que seguem sem saber.

Rodar durante o desenvolvimento:
    python app_web.py
"""
import json
import os
import sys
import threading
import traceback
import webbrowser
from datetime import datetime

import webview

from motor_api import executar_automacao_api
from motor_vinculo import executar_vinculacao
from motor_starlink import executar_starlink
from motor_devolucao import executar_devolucao
from trackit_api_client import TrackitClient, obter_sessao
import cofre
import tutorial
from atualizador import (
    VERSION, RELEASES_URL, MODO_TESTE_UPDATE,
    checar_atualizacao, baixar_nova_versao, aplicar_atualizacao,
)
import motor_clickup as mcu


def _pasta():
    """
    Onde esta a interface, rodando do fonte ou empacotada.

    Empacotado NAO e a pasta do .exe: desde o PyInstaller 6 o --add-data do
    modo onedir cai em _internal/, e e para la que o sys._MEIPASS aponta.
    Montar o caminho a partir do sys.executable acharia uma pasta web que
    nao existe, e o app abriria em branco.
    """
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return os.path.join(base, "web")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "web")


def _diagnostico(msg):
    """
    Rastro em arquivo, so com CADASTRO_DEBUG ligado.

    Empacotado com --windowed nao ha console: se a ponte entre a tela e o
    Python nao subir, o app simplesmente diz "modo esboco" e nao ha onde
    olhar. Numa maquina de outra pessoa isso e um beco sem saida - daqui
    sai a resposta.
    """
    if not os.environ.get("CADASTRO_DEBUG"):
        return
    try:
        os.makedirs(cofre.PASTA, exist_ok=True)
        with open(os.path.join(cofre.PASTA, "diagnostico.log"), "a",
                  encoding="utf-8") as f:
            f.write("{} {}\n".format(datetime.now().isoformat(" ", "seconds"),
                                     msg))
    except Exception:
        pass


class Ponte:
    """
    Metodos chamados pelo JS via window.pywebview.api.<nome>().

    Tudo aqui devolve dict com 'ok'. Excecao que vaza para o JS vira uma
    string sem contexto do lado de la, entao e melhor capturar e mandar o
    motivo junto.
    """

    def __init__(self):
        # Underscore NAO e estilo: o pywebview varre os atributos publicos
        # deste objeto para expor ao JavaScript, e entrava janela adentro
        # tocando propriedades COM do WebView2 fora da thread de UI. Isso
        # travava a janela ("nao esta respondendo") e atrasava a publicacao
        # da ponte, fazendo a tela achar que nao havia Python por tras.
        self._janela = None
        self._sessao = None
        self._empresa = None

    # ---------- canal de volta (Python -> JS) ----------
    def _js(self, codigo):
        try:
            if self._janela:
                self._janela.evaluate_js(codigo)
        except Exception:
            pass        # janela fechando no meio do lote nao e erro

    def log(self, msg, tipo=None):
        self._js("window.appLog({}, {})".format(
            json.dumps(str(msg)), json.dumps(tipo)))

    def progresso(self, feito, total):
        self._js("window.appProgresso({}, {})".format(int(feito), int(total)))

    # ---------- abertura ----------
    def inicio(self):
        """
        Primeira coisa que a tela pergunta, ja carregada.

        A versao e o aviso de atualizacao vinham empurrados pelo Python no
        callback de abertura, e isso TRAVAVA a janela: o evaluate_js chega
        antes de a pagina existir, o WebView2 marshala para a thread de UI
        e ela fica esperando uma resposta que nunca vem. Quem puxa agora e
        o JS, que so pergunta depois de estar pronto.
        """
        _diagnostico("inicio(): a tela alcancou a ponte")
        return {"ok": True, "versao": VERSION}

    # ---------- login guardado ----------
    def login_guardado(self):
        """O que a tela mostra ao abrir. Nunca devolve a senha para o JS."""
        par = cofre.ler_login()
        if not par:
            return {"ok": True, "tem": False}
        usuario, _senha = par
        return {"ok": True, "tem": True, "usuario": usuario}

    def esquecer_login(self):
        cofre.esquecer_login()
        return {"ok": True}

    # ---------- cadastros do cliente ----------
    @staticmethod
    def _listar_bordos(cli):
        """O que alimenta a busca da aba Devolucao."""
        eqs = cli.equipamentos()
        nomes = {e["id"]: (e.get("name") or "").strip() for e in eqs}
        mdts = cli._get("Forms/MobileDataTerminal/GetAllMobileDataTerminal")

        bordos = [{"device": (m.get("deviceID") or "").strip(),
                   "equipamento": nomes.get(m.get("equipmentID"), "")}
                  for m in mdts if (m.get("deviceID") or "").strip()]
        # Equipamento SEM bordo entra na busca: e o que sobra depois de uma
        # devolucao, e sem ele nao da para alcanca-lo para excluir.
        ocupados = {m.get("equipmentID") for m in mdts if m.get("equipmentID")}
        bordos += [{"device": "", "equipamento": nomes[e["id"]]}
                   for e in eqs
                   if e["id"] not in ocupados and nomes[e["id"]]]
        bordos.sort(key=lambda b: (b["device"].lower(),
                                   b["equipamento"].lower()))
        return bordos

    def recarregar(self):
        """
        Rele a base do cliente sem refazer o login.

        A busca da Devolucao trabalhava sobre a foto tirada no Conectar:
        equipamento criado DEPOIS nao aparecia ali, mesmo existindo no
        TracKit. Chamado sozinho ao fim de cada lote, e no botao.
        """
        try:
            if not self._sessao or not self._empresa:
                return {"ok": False, "erro": "Conecte primeiro."}
            cli = TrackitClient(self._empresa, self._sessao)
            return {"ok": True, "bordos": self._listar_bordos(cli)}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    # ---------- conexao ----------
    def conectar(self, empresa, usuario, senha, lembrar=False):
        try:
            empresa = (empresa or "").strip().lower()
            if usuario and not usuario.endswith("@aiko.digital"):
                usuario = usuario + "@aiko.digital"

            # Senha em branco com login guardado = usar o guardado. E o que
            # faz o campo poder ficar vazio depois de marcar "usar sempre".
            if not senha:
                par = cofre.ler_login()
                if par and (not usuario or par[0] == usuario):
                    usuario, senha = par

            sessao = obter_sessao(empresa, usuario=usuario or None,
                                  senha=senha or None, log=self.log)
            cli = TrackitClient(empresa, sessao)
            self._sessao, self._empresa = sessao, empresa

            bordos = self._listar_bordos(cli)

            # So guarda DEPOIS de a credencial provar que funciona. Guardar
            # antes deixaria uma senha errada salva, e o proximo Conectar
            # falharia sozinho sem ninguem entender por que.
            if lembrar and senha:
                cofre.salvar_login(usuario or "", senha)
                self.log("Senha guardada neste PC (cifrada pelo Windows).",
                         "info")
            elif not lembrar:
                cofre.esquecer_login()

            def enxuto(itens):
                return [{"id": x.get("id"), "name": (x.get("name") or "").strip()}
                        for x in itens]

            return {"ok": True,
                    "listas": {"modelo": enxuto(cli.modelos()),
                               "grupo": enxuto(cli.grupos()),
                               "perfil": enxuto(cli.perfis_rede())},
                    "bordos": bordos}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    # ---------- ClickUp ----------
    def clickup_conectar(self):
        try:
            token = mcu.ler_token()
            if not token:
                return {"ok": False,
                        "erro": "Nenhum token guardado. Use 'Trocar token'."}
            cu = mcu.ClickUp(token)
            usuario = cu.usuario().get("username", "?")
            defs = {f["id"]: f for f in cu.campos()}

            campos = []
            for c in mcu.CAMPOS:
                d = defs.get(c["id"], {})
                opcoes = ((d.get("type_config") or {}).get("options") or [])
                campos.append({
                    "id": c["id"], "rotulo": c["rotulo"], "tipo": c["tipo"],
                    "padrao": c.get("padrao", ""),
                    "opcoes": [{"id": o.get("id"),
                                "name": o.get("name") or o.get("label") or ""}
                               for o in opcoes],
                })
            return {"ok": True, "usuario": usuario,
                    "statuses": cu.statuses(), "campos": campos}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def clickup_prefs(self):
        """
        O que ficou marcado e preenchido da ultima vez.

        Mesmo arquivo e mesmo formato da HUD em Tk ({id: {marcado, valor}}),
        para as duas versoes lerem o trabalho uma da outra enquanto convivem.
        """
        u = mcu.ler_ultimos()
        campos = {}
        for c in mcu.CAMPOS:
            d = u.get(c["id"])
            if not isinstance(d, dict):
                continue
            # Data de vinculacao nasce sempre hoje: e o dia em que o kit
            # sai, nao o do lote passado.
            valor = "" if c.get("hoje") else (d.get("valor") or "")
            campos[c["id"]] = {"marcado": bool(d.get("marcado")),
                               "valor": valor}

        st = u.get("_status_card")
        if isinstance(st, dict):
            campos["status_card"] = {"marcado": bool(st.get("marcado")),
                                     "valor": st.get("valor") or ""}

        return {"ok": True,
                "clickup": bool(u.get("_clickup_ligado")),
                "trackit": u.get("_trackit_ligado", True),
                "campos": campos,
                "tem_token": bool(mcu.ler_token())}

    def clickup_salvar_prefs(self, prefs):
        try:
            prefs = prefs or {}
            u = mcu.ler_ultimos()
            for cid, d in (prefs.get("campos") or {}).items():
                chave = "_status_card" if cid == "status_card" else cid
                u[chave] = {"marcado": bool(d.get("marcado")),
                            "valor": d.get("valor") or ""}
            u["_clickup_ligado"] = bool(prefs.get("clickup"))
            u["_trackit_ligado"] = bool(prefs.get("trackit"))
            mcu.salvar_ultimos(u)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def clickup_token(self, token):
        try:
            mcu.salvar_token(token)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    # ---------- utilidades ----------
    def copiar(self, texto):
        """Reserva: o JS tenta a area de transferencia dele primeiro."""
        try:
            self._js("navigator.clipboard.writeText({})".format(
                json.dumps(str(texto))))
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def tutorial(self):
        return {"ok": True, "texto": tutorial.TEXTO}

    # ---------- atualizacao ----------
    def checar_update(self):
        """
        Roda em thread e empurra o resultado quando chega. Em rede lenta a
        consulta demora ate 5s, e travar a abertura do app por causa dela
        seria pior do que o aviso chegar um instante depois.
        """
        def worker():
            info = checar_atualizacao()
            self._js("window.appUpdate({})".format(json.dumps(info)))
        threading.Thread(target=worker, daemon=True).start()
        return {"ok": True}

    def atualizar(self):
        """Baixa o instalador da release e roda - o mesmo caminho da HUD Tk."""
        def worker():
            try:
                info = checar_atualizacao()
                url = info.get("exe_url")
                if not url:
                    self.log("A release nao tem instalador anexado. "
                             "Abrindo a pagina.", "aviso")
                    webbrowser.open(info.get("release_url") or RELEASES_URL)
                    return
                if MODO_TESTE_UPDATE:
                    self.log("MODO_TESTE_UPDATE ligado: nao vou substituir "
                             "nada.", "aviso")
                    return
                self.log("Baixando a versao {}...".format(
                    info.get("versao_remota")))
                caminho = baixar_nova_versao(
                    url, lambda pct: self.progresso(pct, 100))
                self.log("Aplicando. O app fecha e reabre sozinho.")
                aplicar_atualizacao(caminho)
            except Exception as e:
                self.log("Nao consegui atualizar: {}".format(e), "erro")
                self.log("Baixe manualmente em " + RELEASES_URL, "info")
        threading.Thread(target=worker, daemon=True).start()
        return {"ok": True}

    # ---------- execucao ----------
    def executar(self, aba, dados, simulacao):
        """Dispara em thread: o lote demora e a tela nao pode congelar."""
        threading.Thread(target=self._rodar, args=(aba, dados, bool(simulacao)),
                         daemon=True).start()
        return {"ok": True}

    def _rodar(self, aba, dados, simulacao):
        try:
            d = dict(dados or {})
            usuario = (d.get("usuario") or "").strip()
            if usuario and not usuario.endswith("@aiko.digital"):
                usuario = usuario + "@aiko.digital"
            d["usuario"] = usuario or None
            d["senha"] = d.get("senha") or None

            comum = dict(log=self.log, dry_run=simulacao,
                         progresso=self.progresso)

            if aba == "cadastro":
                # O motor aceita termo OU id; mandar o id evita a classe de
                # erro em que "reserva" casava com o grupo parecido.
                d["modelos"] = [str(d.pop("modelo_id"))]
                d["grupos"] = [str(d.pop("grupo_id"))]
                d["perfil"] = [str(d.pop("perfil_id"))]
                executar_automacao_api(d, **comum)

            elif aba == "vinculacao":
                executar_vinculacao(d, **comum)

            elif aba == "starlink":
                self._starlink(d, simulacao)

            elif aba == "devolucao":
                executar_devolucao(d, **comum)

            else:
                self.log("Aba desconhecida: {}".format(aba), "erro")
        except Exception as e:
            self.log("FALHOU :: {}".format(e), "erro")
            self.log(traceback.format_exc().strip(), "erro")
        finally:
            self._js("window.appFim({})")

    def _starlink(self, d, simulacao):
        """
        TracKit primeiro, ClickUp depois - e so se o TracKit saiu limpo.
        Gravar no board um kit que nao entrou no TracKit cria um cadastro
        que ninguem consegue casar depois.
        """
        # Uma lista so para os dois destinos. O patrimonio vem colado ao
        # lado do S/N (e assim que sai do Excel), mas quem usa patrimonio e
        # o ClickUp - o TracKit leva so o serial. Mandar a linha inteira
        # criaria o equipamento "SM - M1HT...RPZ  12345", com o patrimonio
        # grudado no nome.
        itens = mcu.ler_linhas("\n".join(d.get("seriais") or []))

        resumo = None
        if d.get("trackit"):
            so_seriais = dict(d, seriais=[i["serial"] for i in itens])
            resumo = executar_starlink(so_seriais, log=self.log,
                                       dry_run=simulacao,
                                       progresso=self.progresso)
            if resumo.get("falhas") or resumo.get("problemas"):
                self.log("ClickUp nao vai rodar: o TracKit nao saiu limpo.",
                         "aviso")
                return

        if not d.get("clickup"):
            return

        token = mcu.ler_token()
        if not token:
            self.log("ClickUp: nenhum token guardado.", "erro")
            return

        campos = []
        for c in (d.get("campos_cu") or []):
            definicao = next((x for x in mcu.CAMPOS if x["id"] == c["id"]), None)
            if not definicao or c.get("valor") in (None, ""):
                continue
            valor = c["valor"]
            if definicao["tipo"] == "moeda":
                valor = mcu.ler_moeda(valor)
            elif definicao["tipo"] == "data":
                valor = mcu.data_em_ms(mcu.ler_data(valor))
            elif definicao["tipo"] == "etiqueta":
                valor = [valor]
            campos.append({"campo": definicao, "valor": valor,
                           "texto": c.get("texto") or str(c["valor"])})

        status = next((c for c in (d.get("campos_cu") or [])
                       if c["id"] == "status_card"), None)

        mcu.executar_clickup({
            "token": token,
            "prefixo": d.get("prefixo"),
            "itens": itens,
            "campos": campos,
            "status_card": (status or {}).get("valor") or None,
            # So marca "cadastrado no TracKit" quando ele REALMENTE acabou de
            # gravar nesta rodada, sem falha e sem simulacao.
            "trackit_sim": bool(
                d.get("trackit") and resumo and not simulacao
                and not resumo.get("falhas") and not resumo.get("problemas")),
        }, log=self.log, dry_run=simulacao, progresso=self.progresso)


def main():
    ponte = Ponte()
    janela = webview.create_window(
        "Cadastro de Bordo - Aiko  -  v{}".format(VERSION),
        os.path.join(_pasta(), "index.html"),
        js_api=ponte, width=1180, height=780,
        min_size=(760, 520), background_color="#1c1c1c",
    )
    ponte._janela = janela
    # Sem callback de abertura de proposito: qualquer evaluate_js disparado
    # antes de a pagina carregar trava a janela. A tela chama api.inicio()
    # quando estiver pronta.
    webview.start(debug=bool(os.environ.get("CADASTRO_DEBUG")))


if __name__ == "__main__":
    main()
