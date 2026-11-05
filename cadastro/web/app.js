/* ==========================================================================
   Cadastro de Bordo - Aiko : interface

   A regra de negocio NAO mora aqui. Esta tela so coleta o que o usuario
   digitou e entrega para o Python, que continua rodando os mesmos motores
   de sempre (motor_api, motor_vinculo, motor_starlink, motor_devolucao,
   motor_clickup). Toda trava que importa - a dos 30 dias sem comunicacao,
   a conferencia por leitura de volta, o pareamento posicional - fica la.
   Duplicar qualquer uma delas aqui criaria duas versoes da verdade.

   Sem o pywebview (ex: abrindo o index.html direto no navegador) a tela
   entra em MODO ESBOCO: dados de mentira, nada e gravado. Serve para
   olhar o visual sem subir o app.
   ========================================================================== */

const $  = (s) => document.querySelector(s);
const $$ = (s) => Array.from(document.querySelectorAll(s));

const estado = {
  aba: "cadastro",
  listas: { modelo: [], grupo: [], perfil: [] },
  escolha: { modelo: null, grupo: null, perfil: null },
  bordos: [],
  linhaDev: null,        // linha escolhida na tabela de busca
  camposCU: [],
  rodando: false,
  esboco: false,
};

const MARCA_EQUIP = "EQUIP:";
const LIMITE_BUSCA = 50;
const SEM_BORDO = "(sem bordo)";

/* ----------------------------------------------------------- ponte ----- */

function temPonte() { return !!(window.pywebview && window.pywebview.api); }

async function chamar(metodo, ...args) {
  if (!temPonte()) return esbocoResponder(metodo, args);
  try {
    return await window.pywebview.api[metodo](...args);
  } catch (e) {
    return { ok: false, erro: String(e) };
  }
}

/* O Python empurra para ca por evaluate_js. Precisam ser globais. */
window.appLog = (linha, tipo) => escreverLog(linha, tipo);
window.appProgresso = (feito, total) => {
  $("#barra").style.width = total ? (feito / total * 100) + "%" : "0";
};
window.appFim = (resumo) => {
  estado.rodando = false;
  travarAcoes(false);
  $("#barra").style.width = "0";
  if (resumo && resumo.aviso) escreverLog(resumo.aviso, "aviso");
};
window.appUpdate = (info) => {
  if (!info || !info.tem_update) return;
  $("#banner-texto").textContent =
    `Versão ${info.versao_remota} disponível (você está na ${$("#versao").textContent.replace("v","")}).`;
  $("#banner-update").hidden = false;
};

/* ------------------------------------------------------------- log ----- */

function escreverLog(linha, tipo) {
  const el = $("#log");
  const span = document.createElement("span");
  if (tipo) span.className = tipo;
  span.textContent = linha + "\n";
  el.appendChild(span);
  // so acompanha o fim se o usuario ja estava no fim; senao atrapalha quem
  // subiu para ler uma linha antiga enquanto o lote roda
  const noFim = el.scrollHeight - el.scrollTop - el.clientHeight < 40;
  if (noFim) el.scrollTop = el.scrollHeight;
}

/* ------------------------------------------------------------ tema ----- */

function aplicarTema(nome) {
  document.documentElement.dataset.tema = nome;
  try { localStorage.setItem("tema", nome); } catch (e) { /* sem storage */ }
}
$("#btn-tema").onclick = () =>
  aplicarTema(document.documentElement.dataset.tema === "claro" ? "escuro" : "claro");

/* ------------------------------------------------------------ abas ----- */

const ROTULOS = {
  cadastro: "Cadastrar", vinculacao: "Vincular",
  starlink: "Cadastrar Starlink", devolucao: "Desvincular",
};

$$(".aba").forEach((b) => {
  b.onclick = () => {
    estado.aba = b.dataset.aba;
    $$(".aba").forEach((x) => x.classList.toggle("ativa", x === b));
    $$(".conteudo-aba").forEach((x) =>
      x.classList.toggle("ativa", x.dataset.aba === estado.aba));
    $("#btn-executar").textContent = ROTULOS[estado.aba];
  };
});

/* -------------------------------------------------- previa do nome ----- */

function prefixoTicket() { return $("#prefixo-ticket .ativa").dataset.valor; }

function normalizarTicket(v) {
  return String(v || "").trim().replace(/^(HWS|SP)[\s\-_]*/i, "").trim();
}

const modoNome = () => $("#modo-nome .ativa").dataset.valor;

function montarNome(n) {
  // Nome livre: a pessoa escreve a base e a automacao so numera. A
  // numeracao de dois digitos fica: e ela que faz o "parou no bordo" e o
  // pareamento posicional da Vinculacao funcionarem.
  if (modoNome() === "livre") {
    const base = $("#nome-livre").value.trim() || "?";
    return `${base} | ${String(n).padStart(2, "0")}`;
  }
  const zen = $("#zendesk").value.trim();
  const sufixo = (zen && zen.toUpperCase() !== "N") ? ` | #${zen}` : "";
  const emp = $("#empresa").value.trim().toUpperCase() || "?";
  const eq  = $("#equipamento").value.trim().toUpperCase() || "?";
  const tk  = normalizarTicket($("#ticket").value) || "?";
  return `${emp} | ${eq} | ${prefixoTicket()}-${tk}${sufixo} | ${String(n).padStart(2, "0")}`;
}

function faixaLote() {
  const parou = parseInt($("#parou").value, 10) || 0;
  const limite = parseInt($("#limite").value, 10) || 0;
  const primeiro = parou + 1;
  return [primeiro, Math.max(limite, primeiro)];
}

function nomesDoLote() {
  const [a, b] = faixaLote();
  const out = [];
  for (let n = a; n <= b; n++) out.push(montarNome(n));
  return out;
}

function atualizarPrevia() {
  const livre = modoNome() === "livre";
  $("#grade-livre").hidden = !livre;
  $("#grade-montado").hidden = livre;

  if (!livre) {
    // Colou "SP-4471" com o seletor no HWS? O seletor segue o que foi colado -
    // confiar no seletor faria um chamado do SGI nascer como HWS-4471.
    const escrito = (String($("#ticket").value).trim().match(/^(HWS|SP)/i) || [])[1];
    if (escrito) marcarSegmento("#prefixo-ticket", escrito.toUpperCase());
  }
  const [a, b] = faixaLote();
  const qtd = b - a + 1;
  $("#previa-nome").textContent =
    montarNome(a) + (qtd > 1 ? `   (${qtd} nomes)` : "");
}

["empresa", "equipamento", "ticket", "zendesk", "parou", "limite", "nome-livre"]
  .forEach((id) => $("#" + id).addEventListener("input", atualizarPrevia));

function marcarSegmento(sel, valor) {
  $$(`${sel} button`).forEach((b) => b.classList.toggle("ativa", b.dataset.valor === valor));
}
$$(".segmentado button").forEach((b) => {
  b.onclick = () => {
    marcarSegmento("#" + b.parentElement.id, b.dataset.valor);
    atualizarPrevia(); previaStarlink();
  };
});

$("#btn-copiar-nomes").onclick = async () => {
  const exigidos = modoNome() === "livre"
    ? [["nome-livre", "Texto do nome"]]
    : [["empresa","Empresa"],["equipamento","Equipamento"],["ticket","Ticket"]];
  const faltando = exigidos
    .filter(([id]) => !$("#" + id).value.trim()).map(([, r]) => r);
  if (faltando.length) return avisar("Preencha antes: " + faltando.join(", "));
  const nomes = nomesDoLote();
  await copiar(nomes.join("\n"));
  escreverLog(`Copiado(s) ${nomes.length} nome(s) para a área de transferência.`);
};

async function copiar(texto) {
  try { await navigator.clipboard.writeText(texto); }
  catch (e) { await chamar("copiar", texto); }   // fallback pelo Python
}

/* --------------------------------------------------------- conexao ----- */

$("#btn-olho").onclick = () => {
  const i = $("#senha");
  const vendo = i.type === "text";
  i.type = vendo ? "password" : "text";
  $("#btn-olho").textContent = vendo ? "ver" : "ocultar";
};

$("#btn-conectar").onclick = async () => {
  const empresa = $("#empresa").value.trim();
  if (!empresa) return avisar("Informe a sigla da empresa primeiro.");

  chip("indo", "conectando...");
  $("#btn-conectar").disabled = true;
  const r = await chamar("conectar", empresa,
                         $("#usuario").value.trim(), $("#senha").value);
  $("#btn-conectar").disabled = false;

  if (!r || !r.ok) {
    chip("erro", "falhou");
    escreverLog("Conexão falhou: " + ((r && r.erro) || "?"), "erro");
    return avisar((r && r.erro) || "Não conectou.");
  }
  chip("on", "conectado a " + empresa.toUpperCase());
  estado.listas = r.listas || { modelo: [], grupo: [], perfil: [] };
  estado.bordos = r.bordos || [];
  $$(".seletor").forEach((b) => {
    b.disabled = false;
    if (!estado.escolha[b.dataset.chave]) b.textContent = "Selecionar...";
  });
  escreverLog(`Conectado: ${estado.listas.modelo.length} modelos, ` +
              `${estado.listas.grupo.length} grupos, ` +
              `${estado.listas.perfil.length} perfis, ` +
              `${estado.bordos.length} bordos.`, "ok");
  buscarBordo();
};

function chip(estadoChip, texto) {
  const c = $("#chip-conexao");
  c.dataset.estado = estadoChip; c.textContent = texto;
}

/* ------------------------------------------------- seletor em janela --- */

$$(".seletor").forEach((b) => { b.onclick = () => abrirEscolha(b); });

function abrirEscolha(botao) {
  const chave = botao.dataset.chave;
  const itens = estado.listas[chave] || [];
  if (!itens.length) return avisar("Conecte primeiro.");

  const dlg = $("#dlg-escolha");
  const lista = $("#dlg-escolha-lista");
  const busca = $("#dlg-escolha-busca");
  $("#dlg-escolha-titulo").textContent =
    { modelo: "Modelo", grupo: "Grupo", perfil: "Perfil de rede" }[chave];
  busca.value = "";
  let marcado = null;

  function desenhar() {
    const t = busca.value.trim().toLowerCase();
    lista.innerHTML = "";
    itens.filter((x) => !t || (x.name || "").toLowerCase().includes(t))
      .forEach((x) => {
        const li = document.createElement("li");
        li.innerHTML = `<span>${escapar(x.name)}</span> <span class="id">${x.id}</span>`;
        li.onclick = () => {
          marcado = x;
          Array.from(lista.children).forEach((c) => c.classList.toggle("escolhida", c === li));
        };
        li.ondblclick = () => { marcado = x; dlg.close("ok"); };
        lista.appendChild(li);
      });
    if (!lista.children.length)
      lista.innerHTML = '<li class="vazio">nada encontrado</li>';
  }
  busca.oninput = desenhar;
  desenhar();

  dlg.onclose = () => {
    if (dlg.returnValue !== "ok" || !marcado) return;
    estado.escolha[chave] = marcado;
    botao.textContent = marcado.name;
  };
  dlg.showModal();
  busca.focus();
}

const escapar = (s) => String(s == null ? "" : s)
  .replace(/[&<>"]/g, (c) => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;" }[c]));

/* -------------------------------------------------------- vinculacao -- */

const linhasDe = (id) => $("#" + id).value.split("\n").map((s) => s.trim()).filter(Boolean);

$("#seriais").addEventListener("input", () => {
  const n = linhasDe("seriais").length;
  $("#conta-seriais").textContent = n ? `${n} serial(is) na lista` : "";
});

/* ---------------------------------------------------------- starlink -- */

function previaStarlink() {
  const s = linhasDe("sl-seriais");
  const pref = $("#prefixo-sl .ativa").dataset.valor;
  $("#previa-sl").textContent = s.length
    ? `${s.length} kit(s) · o 1º ficará: ${pref} - ${s[0].split(/[\s;,]+/)[0]}`
    : "";
}
$("#sl-seriais").addEventListener("input", previaStarlink);

$("#sl-clickup").onchange = (e) => {
  $("#cartao-clickup").hidden = !e.target.checked;
};

$("#btn-cu-conectar").onclick = async () => {
  $("#dica-cu").textContent = "conectando...";
  const r = await chamar("clickup_conectar");
  if (!r || !r.ok) { $("#dica-cu").textContent = "falhou"; return avisar((r && r.erro) || "?"); }
  $("#dica-cu").textContent = "conectado como " + r.usuario;
  estado.camposCU = r.campos || [];
  desenharCamposCU(r.statuses || []);
};

$("#btn-cu-token").onclick = async () => {
  const t = prompt("Cole o token pessoal do ClickUp (pk_...):");
  if (!t) return;
  const r = await chamar("clickup_token", t.trim());
  $("#dica-cu").textContent = (r && r.ok) ? "token guardado" : "falhou";
};

function desenharCamposCU(statuses) {
  const alvo = $("#campos-cu");
  alvo.innerHTML = "";

  linhaCU(alvo, "status_card", "Status do card",
          selectDe(statuses.map((s) => ({ id: s, name: s }))), true);

  estado.camposCU.forEach((c) => {
    const w = (c.tipo === "lista" || c.tipo === "etiqueta")
      ? selectDe(c.opcoes || [])
      : Object.assign(document.createElement("input"), { value: c.padrao || "" });
    linhaCU(alvo, c.id, c.rotulo, w, false);
  });
}

function selectDe(opcoes) {
  const s = document.createElement("select");
  s.appendChild(new Option("—", ""));
  opcoes.forEach((o) => s.appendChild(new Option(o.name, o.id)));
  return s;
}

function linhaCU(alvo, id, rotulo, widget, marcado) {
  const lbl = document.createElement("label");
  lbl.className = "caixa";
  const cx = Object.assign(document.createElement("input"),
                           { type: "checkbox", checked: marcado });
  lbl.append(cx, document.createTextNode(" " + rotulo));
  widget.disabled = !marcado;
  widget.dataset.campo = id;
  cx.onchange = () => { widget.disabled = !cx.checked; };
  cx.dataset.para = id;
  alvo.append(lbl, widget);
}

function camposMarcadosCU() {
  return $$("#campos-cu [data-campo]")
    .filter((w) => !w.disabled)
    .map((w) => ({ id: w.dataset.campo,
                   valor: w.value,
                   texto: w.tagName === "SELECT"
                          ? (w.selectedOptions[0] || {}).text || "" : w.value }));
}

/* --------------------------------------------------------- devolucao -- */

$("#busca-dev").addEventListener("input", buscarBordo);

function buscarBordo() {
  const termo = $("#busca-dev").value.trim().toLowerCase();
  const corpo = $("#tabela-dev tbody");
  corpo.innerHTML = "";
  estado.linhaDev = null;

  if (!estado.bordos.length) return dicaBusca("conecte para carregar os bordos do cliente");
  if (termo.length < 2) return dicaBusca(`${estado.bordos.length} bordos carregados — digite ao menos 2 caracteres`);

  const achados = estado.bordos.filter((b) =>
    (b.device || "").toLowerCase().includes(termo) ||
    (b.equipamento || "").toLowerCase().includes(termo));

  achados.slice(0, LIMITE_BUSCA).forEach((b) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escapar(b.device || SEM_BORDO)}</td>` +
                   `<td>${escapar(b.equipamento || "(sem equipamento)")}</td>`;
    tr.onclick = () => {
      estado.linhaDev = b;
      Array.from(corpo.children).forEach((c) => c.classList.toggle("escolhida", c === tr));
    };
    tr.ondblclick = () => { estado.linhaDev = b; addNaLista(); };
    corpo.appendChild(tr);
  });

  if (!achados.length) dicaBusca(`nada encontrado para "${termo}"`);
  else if (achados.length > LIMITE_BUSCA)
    dicaBusca(`${achados.length} encontrados — mostrando os ${LIMITE_BUSCA} primeiros, refine a busca`);
  else dicaBusca(`${achados.length} encontrado(s)`);
}

const dicaBusca = (t) => { $("#dica-busca").textContent = t; };

function exigeLinha() {
  if (!estado.linhaDev) { avisar("Clique em um bordo da lista."); return null; }
  return estado.linhaDev;
}

$("#btn-copiar-device").onclick = async () => {
  const b = exigeLinha(); if (!b) return;
  const t = b.device || b.equipamento;
  await copiar(t); dicaBusca("copiado: " + t);
};
$("#btn-copiar-reserva").onclick = async () => {
  const b = exigeLinha(); if (!b) return;
  const t = b.device ? `Reserva - ${b.device}` : b.equipamento;
  await copiar(t); dicaBusca("copiado: " + t);
};
$("#btn-add-lista").onclick = addNaLista;

function addNaLista() {
  const b = exigeLinha(); if (!b) return;
  // Equipamento sem bordo entra marcado: nao ha device para digitar, e o
  // motor precisa saber que a linha e um equipamento a excluir.
  const linha = b.device ? b.device : `${MARCA_EQUIP} ${b.equipamento}`;
  if (linhasDe("dev-lista").includes(linha)) return dicaBusca(`${linha} já está na lista`);
  const atual = $("#dev-lista").value.replace(/\s+$/, "");
  $("#dev-lista").value = (atual ? atual + "\n" : "") + linha + "\n";
  contarDev(); dicaBusca(`${linha} adicionado à lista`);
}

function contarDev() {
  const n = linhasDe("dev-lista").length;
  $("#conta-dev").textContent = n ? `${n} item(ns) na lista` : "";
}
$("#dev-lista").addEventListener("input", contarDev);

const acaoDev = () => ($$('input[name="acao-dev"]').find((r) => r.checked) || {}).value;

/* ------------------------------------------------------- confirmacao -- */

function avisar(msg) { escreverLog(msg, "aviso"); alert(msg); }

function confirmarExclusao(quantos, empresa) {
  return new Promise((resolve) => {
    const dlg = $("#dlg-excluir"), campo = $("#dlg-excluir-campo"), ok = $("#dlg-excluir-ok");
    $("#dlg-excluir-texto").textContent =
      `Até ${quantos} equipamento(s) de ${empresa} serão EXCLUÍDOS.`;
    campo.value = ""; ok.disabled = true;
    campo.oninput = () => { ok.disabled = campo.value.trim().toUpperCase() !== "EXCLUIR"; };
    dlg.onclose = () => resolve(dlg.returnValue === "ok");
    dlg.showModal(); campo.focus();
  });
}

/* ---------------------------------------------------------- executar -- */

function travarAcoes(travado) {
  $("#btn-executar").disabled = travado;
  $("#btn-simular").disabled = travado;
}

async function executar(simulacao) {
  if (estado.rodando) return;
  const empresa = $("#empresa").value.trim().toUpperCase();
  if (!empresa) return avisar("Informe a sigla da empresa.");

  const dados = { empresa, usuario: $("#usuario").value.trim(), senha: $("#senha").value };
  let quantos = 0;

  if (estado.aba === "cadastro") {
    for (const k of ["modelo", "grupo", "perfil"])
      if (!estado.escolha[k]) return avisar("Escolha o " + k + ".");
    const livre = modoNome() === "livre";
    const tk = normalizarTicket($("#ticket").value);
    if (livre && !$("#nome-livre").value.trim())
      return avisar("Escreva o texto do nome.");
    if (!livre && !/^\d+$/.test(tk))
      return avisar("Ticket inválido. Use o número (12312) ou o ticket inteiro (HWS-12312, SP-4471).");
    Object.assign(dados, {
      nome_livre: livre ? $("#nome-livre").value.trim() : "",
      equipamento: $("#equipamento").value.trim().toUpperCase(),
      ticket: tk, prefixo_ticket: prefixoTicket(),
      zendesk: $("#zendesk").value.trim() || "N",
      parou: parseInt($("#parou").value, 10) || 0,
      limite: parseInt($("#limite").value, 10) || 0,
      modelo_id: estado.escolha.modelo.id,
      grupo_id: estado.escolha.grupo.id,
      perfil_id: estado.escolha.perfil.id,
    });
    quantos = dados.limite - dados.parou;
    if (quantos < 1) return avisar("A quantidade total precisa ser maior que 'Parou no bordo'.");

  } else if (estado.aba === "vinculacao") {
    dados.filtro = $("#filtro").value.trim();
    dados.seriais = linhasDe("seriais");
    if (!dados.filtro) return avisar("Informe o filtro do equipamento (ex: HWS-6848).");
    if (!dados.seriais.length) return avisar("Cole os seriais, um por linha.");
    quantos = dados.seriais.length;

  } else if (estado.aba === "starlink") {
    dados.seriais = linhasDe("sl-seriais");
    dados.prefixo = $("#prefixo-sl .ativa").dataset.valor;
    dados.trackit = $("#sl-trackit").checked;
    dados.clickup = $("#sl-clickup").checked;
    dados.campos_cu = dados.clickup ? camposMarcadosCU() : [];
    if (!dados.seriais.length) return avisar("Cole os S/N, um por linha.");
    if (!dados.trackit && !dados.clickup) return avisar("Escolha ao menos um destino.");
    quantos = dados.seriais.length;

  } else {
    dados.seriais = linhasDe("dev-lista");
    dados.excluir_equipamento = acaoDev() === "excluir";
    dados.forcar_ativos = $("#forcar-ativos").checked;
    if (!dados.seriais.length) return avisar("Cole os deviceID, um por linha.");
    quantos = dados.seriais.length;
  }

  if (!simulacao) {
    const texto = {
      cadastro:   `Isto vai CRIAR ${quantos} equipamento(s) em ${empresa} — de verdade, na produção.\n\nO 1º ficará: ${montarNome(faixaLote()[0])}`,
      vinculacao: `Isto vai VINCULAR ${quantos} bordo(s) em ${empresa} — de verdade, na produção.\n\nFiltro: ${dados.filtro}`,
      starlink:   `Isto vai CADASTRAR ${quantos} Starlink em ${empresa} — de verdade, na produção.`,
      devolucao:  `Isto vai DESVINCULAR ${quantos} bordo(s) em ${empresa} — de verdade, na produção.\n\n` +
                  `Os equipamentos onde eles estão param de reportar posição.` +
                  (dados.forcar_ativos ? "\n\n⚠ A trava de comunicação está DESLIGADA." : ""),
    }[estado.aba];
    if (!confirm(texto + "\n\nContinuar?")) return;

    // Exclusao e irreversivel e o botao e o mesmo do desvinculo, que da
    // para refazer. Sem uma barreira diferente, as duas custam o mesmo
    // clique distraido.
    if (estado.aba === "devolucao" && dados.excluir_equipamento &&
        !(await confirmarExclusao(quantos, empresa))) return;
  }

  estado.rodando = true; travarAcoes(true);
  $("#barra").style.width = "0";
  escreverLog((simulacao ? "[SIMULAÇÃO] " : "") +
              `${ROTULOS[estado.aba]}: ${quantos} item(ns) em ${empresa}...`);
  const r = await chamar("executar", estado.aba, dados, simulacao);
  if (r && r.erro) escreverLog("Erro: " + r.erro, "erro");
  window.appFim(r);
}

$("#btn-executar").onclick = () => executar(false);
$("#btn-simular").onclick  = () => executar(true);
$("#btn-limpar-log").onclick = () => { $("#log").innerHTML = ""; };
$("#btn-tutorial").onclick = () => chamar("tutorial");
$("#btn-depois").onclick = () => { $("#banner-update").hidden = true; };
$("#btn-atualizar").onclick = () => chamar("atualizar");

/* ------------------------------------------------------- modo esboco -- */

function esbocoResponder(metodo, args) {
  const mentira = (n, f) => Array.from({ length: n }, (_, i) => f(i));
  if (metodo === "conectar") {
    return Promise.resolve({ ok: true,
      listas: {
        modelo: mentira(14, (i) => ({ id: 100 + i, name: ["Feller","Skidder","Caminhão Pipa","Pá Carregadeira","Escavadeira","Comboio","Motoniveladora","Starlink","Veículo","Garra Traçadora","Trator de Pneu","Rolo Compactador","Caminhão Oficina","Mecânico"][i] })),
        grupo:  mentira(6, (i) => ({ id: 200 + i, name: ["Feller","Skidder","Reservas/Inativos","Devoluções","Starlink Aiko","Caminhão Pipa"][i] })),
        perfil: mentira(3, (i) => ({ id: 300 + i, name: ["Equipamentos","Courier Hotspot","Bullet"][i] })),
      },
      bordos: [
        { device: "03742185a370f861", equipamento: "Reserva - 03742185a370f861" },
        { device: "77a8f86b9c501643", equipamento: "GARRA PICADORA 02" },
        { device: "3544'8914257'65'", equipamento: "ABE | COMODATO | HWS-13041 | #54877 | 01" },
        { device: "24861d35aec12b81", equipamento: "DEVOLUÇÃO | ANTIGO TC6172 | TICKET 55908" },
        { device: "", equipamento: "ABE | TESTEARTHUR | HWS-1111111 | 01" },
        ...mentira(40, (i) => ({ device: "aaaa" + String(i).padStart(12, "0"),
                                 equipamento: "Pipa " + String(i + 1).padStart(2, "0") })),
      ] });
  }
  if (metodo === "clickup_conectar") {
    return Promise.resolve({ ok: true, usuario: "Arthur (esboço)",
      statuses: ["ESTOQUE", "EM USO", "MANUTENÇÃO"],
      campos: [
        { id: "nf", rotulo: "NF", tipo: "texto" },
        { id: "valor", rotulo: "Valor de compra (R$)", tipo: "moeda" },
        { id: "aquisicao", rotulo: "Data de aquisição", tipo: "data" },
        { id: "condicao", rotulo: "Condição", tipo: "lista",
          opcoes: [{ id: "n", name: "Novo" }, { id: "u", name: "Usado" }] },
        { id: "cliente", rotulo: "Cliente", tipo: "etiqueta",
          opcoes: [{ id: "abe", name: "ABE" }, { id: "vaf", name: "VAF" }, { id: "vap", name: "VAP" }] },
        { id: "plano", rotulo: "Plano de dados", tipo: "lista",
          opcoes: [{ id: "1", name: "50GB" }, { id: "2", name: "Ilimitado" }] },
        { id: "rede", rotulo: "Nome da rede", tipo: "texto", padrao: "Coletor Aiko" },
        { id: "senha", rotulo: "Senha da rede", tipo: "texto", padrao: "Aiko@2024" },
      ] });
  }
  if (metodo === "executar") {
    const [aba, dados, simulacao] = args;
    const itens = (dados.seriais || []).length ||
                  ((dados.limite || 0) - (dados.parou || 0)) || 3;
    escreverLog("MODO ESBOÇO — nada é gravado; isto é só para ver a tela.", "aviso");
    let i = 0;
    const t = setInterval(() => {
      i++;
      window.appProgresso(i, itens);
      escreverLog(`[${i}/${itens}] ${simulacao ? "SIMULAÇÃO" : "OK"} :: exemplo ${i}`,
                  i % 5 === 0 ? "aviso" : "ok");
      if (i >= itens) { clearInterval(t); escreverLog("Fim (esboço)."); window.appFim({}); }
    }, 220);
    return Promise.resolve({ ok: true });
  }
  return Promise.resolve({ ok: true });
}

/* ------------------------------------------------------------ inicio -- */

function iniciar() {
  let tema = "escuro";
  try { tema = localStorage.getItem("tema") || "escuro"; } catch (e) { /* ignora */ }
  aplicarTema(tema);
  atualizarPrevia(); previaStarlink(); contarDev(); buscarBordo();

  if (!temPonte()) {
    estado.esboco = true;
    chip("indo", "modo esboço");
    escreverLog("Modo esboço: sem Python por trás. Os dados são de mentira e " +
                "nada é gravado — serve para avaliar a tela.", "aviso");
    escreverLog('Clique em "Conectar" para popular os seletores e a busca.', "info");
  }
}

if (temPonte()) iniciar();
else {
  window.addEventListener("pywebviewready", iniciar);
  // sem pywebview o evento nunca vem; segue em modo esboco
  setTimeout(() => { if (!document.documentElement.dataset.tema) iniciar(); }, 120);
}
