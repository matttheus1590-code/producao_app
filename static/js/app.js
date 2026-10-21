// Fecha alertas automaticamente após alguns segundos
document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll(".alert").forEach((el) => {
    setTimeout(() => {
      const alerta = bootstrap.Alert.getOrCreateInstance(el);
      alerta.close();
    }, 6000);
  });
});

// Menu lateral: abre/fecha em telas pequenas (celular/tablet)
document.addEventListener("DOMContentLoaded", () => {
  const sidebar = document.getElementById("appSidebar");
  const backdrop = document.getElementById("appSidebarBackdrop");
  const botaoAbrir = document.getElementById("btnToggleSidebar");
  if (!sidebar || !backdrop || !botaoAbrir) return;

  function fechar() {
    sidebar.classList.remove("app-sidebar-open");
    backdrop.classList.remove("show");
  }

  botaoAbrir.addEventListener("click", () => {
    sidebar.classList.add("app-sidebar-open");
    backdrop.classList.add("show");
  });
  backdrop.addEventListener("click", fechar);
  sidebar.querySelectorAll(".app-nav-link").forEach((link) => link.addEventListener("click", fechar));
});

// Menu lateral: mantém a posição de rolagem dos grupos entre navegações —
// pedido do Bruno (02/09/2026): como cada clique num link recarrega a
// página inteira (não é SPA), sem isso o menu sempre voltava pro topo,
// mesmo que estivesse rolado lá embaixo (ex: em Cadastros/P&D). Guarda o
// scrollTop em sessionStorage (só desta aba) a cada rolagem e restaura
// assim que a próxima página carrega.
document.addEventListener("DOMContentLoaded", () => {
  const nav = document.querySelector(".app-sidebar-nav");
  if (!nav) return;
  const CHAVE_SCROLL_MENU = "menuLateralScrollTop";

  const salvo = sessionStorage.getItem(CHAVE_SCROLL_MENU);
  if (salvo !== null) {
    nav.scrollTop = parseInt(salvo, 10) || 0;
  }

  let salvamentoPendente = false;
  nav.addEventListener("scroll", () => {
    if (salvamentoPendente) return;
    salvamentoPendente = true;
    requestAnimationFrame(() => {
      sessionStorage.setItem(CHAVE_SCROLL_MENU, String(nav.scrollTop));
      salvamentoPendente = false;
    });
  });
});

/* ---------------------------------------------------------------------
   Tela cheia (pedido do Bruno, 09/10/2026): "deixar essa área full screen/
   tela cheia, pensando numa melhor visualização da listagem... mesma coisa
   nas estações". Componente genérico: qualquer botão com
   data-tela-cheia-alvo="#id" (e, opcional, data-tela-cheia-titulo="..." e
   data-tela-cheia-extra="#id1,#id2" pra levar junto, como a caixa de busca)
   leva aquele elemento pra uma camada que cobre a janela inteira — e tenta
   também a tela cheia de verdade do navegador (F11-like, via Fullscreen API
   na página toda, pra os modais continuarem funcionando). O elemento é
   MOVIDO (não clonado), então botões, formulários, busca e redimensionar
   coluna seguem funcionando; ao sair, tudo volta exatamente pro lugar.
   Esc ou o botão "Sair" fecham. Se a página recarregar logo em seguida (ex.:
   "Avançar" no Kanban, que envia formulário), a tela cheia é retomada.
--------------------------------------------------------------------- */
(function () {
  const CHAVE = "telaCheiaRetomar";
  let estado = null;

  function mover(el, destino) {
    const marcador = document.createComment("tela-cheia");
    el.parentNode.insertBefore(marcador, el);
    destino.appendChild(el);
    return { el: el, marcador: marcador };
  }

  function devolver(item) {
    if (item.marcador.parentNode) {
      item.marcador.parentNode.insertBefore(item.el, item.marcador);
      item.marcador.parentNode.removeChild(item.marcador);
    }
  }

  function entrar(botao, nativo) {
    if (estado) return;
    const seletor = botao.getAttribute("data-tela-cheia-alvo");
    const alvo = seletor ? document.querySelector(seletor) : null;
    if (!alvo) return;

    const overlay = document.createElement("div");
    overlay.className = "tela-cheia-overlay";
    const barra = document.createElement("div");
    barra.className = "tela-cheia-barra";
    const titulo = document.createElement("span");
    titulo.className = "tela-cheia-titulo";
    titulo.textContent = botao.getAttribute("data-tela-cheia-titulo") || "Tela cheia";
    const extra = document.createElement("div");
    extra.className = "tela-cheia-extra";
    const sair = document.createElement("button");
    sair.type = "button";
    sair.className = "btn btn-light btn-sm tela-cheia-sair";
    sair.innerHTML = '<i class="bi bi-fullscreen-exit"></i> Sair da tela cheia (Esc)';
    sair.addEventListener("click", fechar);
    barra.appendChild(titulo);
    barra.appendChild(extra);
    barra.appendChild(sair);
    const corpo = document.createElement("div");
    corpo.className = "tela-cheia-corpo";
    overlay.appendChild(barra);
    overlay.appendChild(corpo);
    document.body.appendChild(overlay);

    const extras = (botao.getAttribute("data-tela-cheia-extra") || "")
      .split(",").map(function (s) { return s.trim(); }).filter(Boolean)
      .map(function (s) { return document.querySelector(s); }).filter(Boolean)
      .map(function (el) { return mover(el, extra); });
    // Modais que vivem dentro do alvo (ex.: matéria-prima no Kanban) vão pro
    // <body> enquanto a tela cheia está ligada — senão ficariam presos atrás
    // da camada (empilhamento) e não abririam.
    const modais = Array.prototype.slice.call(alvo.querySelectorAll(".modal"))
      .map(function (el) { return mover(el, document.body); });
    const item = mover(alvo, corpo);
    alvo.classList.add("tela-cheia-ativa");
    document.body.classList.add("tela-cheia-body");

    estado = { botao: botao, seletor: seletor, overlay: overlay, extras: extras, modais: modais, item: item, nativo: false };

    if (nativo && document.documentElement.requestFullscreen) {
      document.documentElement.requestFullscreen().then(function () {
        if (estado) estado.nativo = true;
      }).catch(function () { /* sem tela cheia nativa: fica só a camada */ });
    }
  }

  function fechar() {
    if (!estado) return;
    const e = estado;
    estado = null;
    e.item.el.classList.remove("tela-cheia-ativa");
    devolver(e.item);
    e.modais.forEach(devolver);
    e.extras.forEach(devolver);
    if (e.overlay.parentNode) e.overlay.parentNode.removeChild(e.overlay);
    document.body.classList.remove("tela-cheia-body");
    if (document.fullscreenElement && document.exitFullscreen) {
      document.exitFullscreen().catch(function () {});
    }
    try { sessionStorage.removeItem(CHAVE); } catch (err) { /* ignora */ }
  }

  document.addEventListener("click", function (ev) {
    const botao = ev.target.closest("[data-tela-cheia-alvo]");
    if (!botao) return;
    ev.preventDefault();
    if (estado) fechar(); else entrar(botao, true);
  });

  // Esc: na tela cheia nativa o navegador consome a tecla e dispara fullscreenchange; sem a nativa, vale o keydown.
  document.addEventListener("fullscreenchange", function () {
    if (estado && !document.fullscreenElement) fechar();
  });
  document.addEventListener("keydown", function (ev) {
    if (ev.key === "Escape" && estado && !document.querySelector(".modal.show")) {
      fechar();
    }
  });

  window.addEventListener("beforeunload", function () {
    if (!estado) return;
    try {
      sessionStorage.setItem(CHAVE, JSON.stringify({ seletor: estado.seletor, t: Date.now() }));
    } catch (err) { /* ignora */ }
  });

  document.addEventListener("DOMContentLoaded", function () {
    let salvo = null;
    try {
      salvo = JSON.parse(sessionStorage.getItem(CHAVE) || "null");
      sessionStorage.removeItem(CHAVE);
    } catch (err) { salvo = null; }
    if (!salvo || Date.now() - salvo.t > 20000) return;
    const botao = Array.prototype.find.call(
      document.querySelectorAll("[data-tela-cheia-alvo]"),
      function (b) { return b.getAttribute("data-tela-cheia-alvo") === salvo.seletor; }
    );
    if (botao) entrar(botao, false);
  });
})();


/* ------------------------------------------------------------------------
 * Máscara de moeda / número no formato brasileiro (pedido do Bruno,
 * 21/10/2026: "deixe fácil e correto o formato da moeda REAL, está confuso
 * os pontos, sinais etc").
 *
 *   <input type="text" inputmode="decimal" data-br="moeda">  -> 5.273,91
 *   <input type="text" inputmode="decimal" data-br="qtd">    -> 2  /  2,5
 *
 * Regra de leitura (a MESMA do servidor, _parse_valor_br em app.py):
 *   - tem vírgula  -> vírgula é o decimal, pontos são milhar ("5.273,91");
 *   - só pontos no padrão de milhar ("100.000", "1.250.000") -> milhar;
 *   - senão o ponto é decimal ("12.5", "5273.91").
 * Ao sair do campo o texto é reescrito já formatado ("100.000" -> "100.000,00")
 * e uma linha logo abaixo mostra como o sistema entendeu ("= R$ 100.000,00").
 * ---------------------------------------------------------------------- */
(function () {
  function parse(texto) {
    if (texto === null || texto === undefined) return null;
    var t = String(texto).replace(/R\$/g, "").replace(/[\s ]/g, "");
    if (t === "") return null;
    if (t.indexOf(",") >= 0) t = t.replace(/\./g, "").replace(",", ".");
    else if (/^\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, "");
    if (!/^(\d+(\.\d+)?|\.\d+)$/.test(t)) return NaN;
    return parseFloat(t);
  }

  function fmt(n, minD, maxD) {
    return n.toLocaleString("pt-BR", { minimumFractionDigits: minD, maximumFractionDigits: maxD });
  }

  // número (0 se vazio/inválido) — pros cálculos de total na tela
  function num(texto) {
    var n = parse(texto);
    return n === null || isNaN(n) ? 0 : n;
  }

  function casas(input) {
    return input.getAttribute("data-br") === "qtd" ? [0, 3] : [2, 5];
  }

  function previewDe(input) {
    if (input.getAttribute("data-br") !== "moeda") return null;
    var ancora = input.closest(".input-group") || input;
    var el = ancora.nextElementSibling;
    if (!el || !el.classList.contains("br-preview")) {
      el = document.createElement("div");
      el.className = "br-preview small";
      ancora.insertAdjacentElement("afterend", el);
    }
    return el;
  }

  function atualizar(input, formatar) {
    var n = parse(input.value);
    var c = casas(input);
    var prev = previewDe(input);
    if (n === null) {
      input.classList.remove("is-invalid");
      if (prev) prev.textContent = "";
      return;
    }
    if (isNaN(n)) {
      input.classList.add("is-invalid");
      if (prev) { prev.textContent = "Valor inválido — use o formato 1.234,56"; prev.className = "br-preview small text-danger"; }
      return;
    }
    input.classList.remove("is-invalid");
    var texto = fmt(n, c[0], c[1]);
    if (formatar) input.value = texto;
    if (prev) { prev.textContent = "= R$ " + texto; prev.className = "br-preview small text-success"; }
  }

  document.addEventListener("input", function (ev) {
    var el = ev.target;
    if (!el.matches || !el.matches("input[data-br]")) return;
    var limpo = el.value.replace(/[^0-9.,R$\s-]/g, "");
    if (limpo !== el.value) el.value = limpo;
    atualizar(el, false);
  });

  document.addEventListener("focusout", function (ev) {
    var el = ev.target;
    if (el.matches && el.matches("input[data-br]")) atualizar(el, true);
  });

  document.addEventListener("DOMContentLoaded", function () {
    document.querySelectorAll("input[data-br]").forEach(function (el) { atualizar(el, false); });
  });

  window.BR = { parse: parse, fmt: fmt, num: num };
})();


/* Linha inteira clicável (Listagem Geral — pedido do Bruno, 21/10/2026): um
 * clique em qualquer área da linha abre o link em data-href (visualização do
 * pedido, sem edição). Não dispara quando o clique foi num link/botão/campo
 * (esses seguem o comportamento próprio), nem ao selecionar texto. Ctrl/Cmd ou
 * botão do meio abrem em nova aba. */
(function () {
  function destino(ev) {
    var tr = ev.target.closest && ev.target.closest("tr[data-href]");
    if (!tr) return null;
    if (ev.target.closest("a, button, input, select, textarea, label, summary, [data-no-row-click]")) return null;
    var sel = window.getSelection && window.getSelection().toString();
    if (sel) return null;
    return tr.getAttribute("data-href");
  }
  document.addEventListener("click", function (ev) {
    if (ev.button !== 0) return;
    var href = destino(ev);
    if (!href) return;
    if (ev.ctrlKey || ev.metaKey) window.open(href, "_blank");
    else window.location.href = href;
  });
  document.addEventListener("auxclick", function (ev) {
    if (ev.button !== 1) return;
    var href = destino(ev);
    if (href) window.open(href, "_blank");
  });
})();
