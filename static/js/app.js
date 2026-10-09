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
