// Menú móvil
(function () {
  var abrir = document.getElementById("abrir-menu");
  var velo = document.getElementById("velo");
  if (abrir) abrir.addEventListener("click", function () { document.body.classList.add("menu-abierto"); });
  if (velo) velo.addEventListener("click", function () { document.body.classList.remove("menu-abierto"); });
  document.querySelectorAll("[data-cerrar]").forEach(function (b) {
    b.addEventListener("click", function () { b.closest(".mensaje").remove(); });
  });
  document.querySelectorAll("form[data-confirmar]").forEach(function (f) {
    f.addEventListener("submit", function (e) { if (!confirm(f.dataset.confirmar)) e.preventDefault(); });
  });
})();
