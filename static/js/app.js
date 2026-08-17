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
