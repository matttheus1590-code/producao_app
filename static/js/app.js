// Fecha alertas automaticamente após alguns segundos
document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll(".alert").forEach((el) => {
    setTimeout(() => {
      const alerta = bootstrap.Alert.getOrCreateInstance(el);
      alerta.close();
    }, 6000);
  });
});
