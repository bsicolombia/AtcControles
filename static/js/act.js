
// ======================================================
// ATC CONTROLES - FUNCIONES GENERALES
// ======================================================

document.addEventListener("DOMContentLoaded", function () {

    const abrir = document.getElementById("abrir-menu");
    const lateral = document.getElementById("menu-lateral");
    const velo = document.getElementById("velo");

    // ABRIR MENÚ MÓVIL
    function abrirMenu() {
        if (!lateral || !velo) return;

        lateral.classList.add("abierto");
        velo.classList.add("activo");

        if (abrir) {
            abrir.setAttribute("aria-expanded", "true");
        }

        document.body.classList.add("menu-abierto");
    }

    // CERRAR MENÚ MÓVIL
    function cerrarMenu() {
        if (!lateral || !velo) return;

        lateral.classList.remove("abierto");
        velo.classList.remove("activo");

        if (abrir) {
            abrir.setAttribute("aria-expanded", "false");
        }

        document.body.classList.remove("menu-abierto");
    }

    // BOTÓN ABRIR
    if (abrir) {
        abrir.addEventListener("click", function () {

            if (lateral.classList.contains("abierto")) {
                cerrarMenu();
            } else {
                abrirMenu();
            }

        });
    }

    // CERRAR AL HACER CLIC EN EL VELO
    if (velo) {
        velo.addEventListener("click", cerrarMenu);
    }

    // CERRAR CON ESCAPE
    document.addEventListener("keydown", function (e) {
        if (e.key === "Escape") {
            cerrarMenu();
        }
    });

    // CERRAR MENÚ AL SELECCIONAR UNA OPCIÓN
    document.querySelectorAll(".nav > a").forEach(function (enlace) {
        enlace.addEventListener("click", function () {
            if (window.innerWidth <= 760) {
                cerrarMenu();
            }
        });
    });

    // CERRAR MENSAJES
    document.querySelectorAll("[data-cerrar]").forEach(function (boton) {

        boton.addEventListener("click", function () {
            const mensaje = boton.closest(".mensaje");

            if (mensaje) {
                mensaje.remove();
            }
        });

    });

    // CONFIRMAR FORMULARIOS
    document.querySelectorAll("form[data-confirmar]").forEach(function (formulario) {

        formulario.addEventListener("submit", function (e) {

            if (!confirm(formulario.dataset.confirmar)) {
                e.preventDefault();
            }

        });

    });

    // RESTABLECER MENÚ AL CAMBIAR A ESCRITORIO
    window.addEventListener("resize", function () {

        if (window.innerWidth > 760) {
            cerrarMenu();
        }

    });

});
