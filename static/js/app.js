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
