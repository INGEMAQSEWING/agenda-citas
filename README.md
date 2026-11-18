# Agenda de citas

Agenda gratis para salones de belleza, barberías y uñas. La clienta escoge el servicio, el día y la hora; el salón acepta o rechaza desde su celular. Sin precios, sin WhatsApp y sin instalar nada.

- **Página:** `index.html`. Con `?s=<salón>` funciona en línea; sin eso queda en modo de prueba, guardando todo en el navegador.
- **Programa del salón:** `apps-script/Codigo.gs`. Va dentro de una Hoja de cálculo de Google del salón, guarda las citas y envía los correos.
- **Lista de salones:** `salones.js`. Relaciona el nombre corto del link con la dirección del programa de cada salón.
- **Guías:** `tutorial/tutorial-salon.html` y `tutorial/tutorial-clienta.html`.

## Cómo agregar un salón

1. En el Gmail del salón, crea una **Hoja de cálculo de Google** nueva, llamada por ejemplo "Agenda Salón Bella".
2. En la hoja, ve a **Extensiones → Apps Script**. Borra lo que aparece y pega todo el contenido de `apps-script/Codigo.gs`.
3. Arriba del código, cambia `CAMBIA-ESTA-CLAVE` por la clave del salón y guarda (ícono de disquete).
4. Arriba, escoge la función **prepararHoja** y toca **Ejecutar**. Acepta los permisos: *Revisar permisos → la cuenta → Configuración avanzada → Ir a … → Permitir*.
5. Toca **Implementar → Nueva implementación**. En el engranaje escoge **Aplicación web** y llena:
   - *Ejecutar como*: **Yo**.
   - *Quién tiene acceso*: **Cualquier usuario**.

   Luego toca **Implementar** y copia la **URL de la aplicación web**, la que termina en `/exec`.
6. Agrega una línea en `salones.js`, por ejemplo `"salonbella": "https://script.google.com/macros/s/.../exec"`, y súbelo.
7. Links que quedan:
   - Clientas: `https://ingemaqsewing.github.io/agenda-citas/?s=salonbella`
   - Panel del salón: `https://ingemaqsewing.github.io/agenda-citas/?s=salonbella#/salon`

> Si cambias el código después, usa **Implementar → Gestionar implementaciones → editar → Versión nueva**, para que la URL no cambie.
