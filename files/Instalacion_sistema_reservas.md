# Tu sistema de reservas — instalación

Todo corre dentro de tu cuenta de Google, pegado a tu propia hoja de cálculo. No hay plataforma externa, no hay comisiones, no hay tope de citas, y los datos de tus clientes son tuyos.

**Tiempo de montaje: 20 minutos.** Se hace una sola vez.

---

## Antes de empezar

Ya debes tener el libro `Control_Barberia_RIIF` importado como Hoja de cálculo de Google, con los nombres de tus barberos puestos en Config (celdas A5 y A6). Si todavía no, hazlo primero: nada de esto funciona sin eso.

---

## Paso 1 — Abrir el editor de código

En tu hoja de cálculo de Google: menú **Extensiones** → **Apps Script**.

Se abre una pestaña nueva. Arriba a la izquierda dice "Proyecto sin título": renómbralo a **Reservas RIIF**.

---

## Paso 2 — Pegar el código

Vas a ver un archivo llamado `Código.gs` con unas líneas de ejemplo.

1. Selecciona todo lo que hay adentro y bórralo.
2. Abre el archivo `Codigo.gs` que te pasé, copia todo su contenido y pégalo ahí.
3. Guarda con **Ctrl+S** (o el icono del disquete).

---

## Paso 3 — Ajustar tus datos

Arriba del todo, en el bloque `AJUSTES`, cambia estas líneas:

```javascript
var WHATSAPP    = '573000000000';   // tu WhatsApp real: 57 + número, sin + ni espacios
var EMAIL_AVISO = '';               // tu correo, para recibir aviso de cada cita
var DIRECCION   = 'Cartagena';      // tu dirección real, aparece en la confirmación
```

Ejemplo: si tu WhatsApp es 300 555 1234, ahí va `'573005551234'`.

Los otros ajustes (anticipación mínima de 2 horas, cupos cada 15 minutos, reservas hasta 21 días adelante) están puestos en valores sensatos. Cámbialos después si quieres.

Guarda otra vez.

---

## Paso 4 — Agregar las tres páginas

En el panel izquierdo, junto a **Archivos**, dale al **+** → **HTML**. Vas a hacer esto tres veces.

1. Nómbralo exactamente `Reservar` (sin .html, el editor se lo pone solo). Borra lo que traiga, pega todo el contenido de `Reservar.html`. Guarda.
2. **+** → **HTML**, nómbralo exactamente `Cancelar`, pega el contenido de `Cancelar.html`. Guarda.
3. **+** → **HTML**, nómbralo exactamente `Logo`, pega el contenido de `Logo.html`. Guarda.

Los nombres tienen que ser exactos: `Reservar`, `Cancelar` y `Logo`. Si les pones otro nombre, la página no abre o sale sin logo.

`Logo` es tu logo convertido a vector. Al ser código y no imagen, se ve nítido en cualquier pantalla, carga instantáneo y no depende de que ningún archivo siga existiendo en tu Drive.

---

## Paso 5 — Crear las hojas del sistema

Todavía en el editor, arriba hay un selector de función que dice `doGet`. Cámbialo a **`configurarHojas`** y dale **Ejecutar**.

La primera vez Google te va a pedir permisos:

1. **Revisar permisos** → escoge tu cuenta
2. Aparece "Google no ha verificado esta aplicación". Es normal: la aplicación eres tú. Dale **Configuración avanzada** → **Ir a Reservas RIIF (no seguro)**
3. **Permitir**

Vuelve a la hoja de cálculo. Van a aparecer tres pestañas nuevas:

| Hoja | Para qué |
|---|---|
| **Reservas** | Cada cita que entra por tu link. No la edites a mano. |
| **Horarios** | Qué días y a qué horas trabaja cada barbero |
| **Bloqueos** | Ratos que quieres cerrar: almuerzo, vacaciones, un evento |

---

## Paso 6 — Ajustar los horarios

Abre la hoja **Horarios**. Quedó puesto lunes a sábado, 9:00 a 20:00, domingo cerrado.

- Si abres domingo, agrega esas filas.
- Si un barbero entra más tarde, cámbiale la hora.
- Si un barbero no trabaja los lunes, **borra esa fila**. Sin fila, ese día no aparece disponible para él.

Formato de las horas: `09:00`, `20:00`. Con cero adelante y dos puntos.

---

## Paso 7 — Publicar

En el editor de Apps Script, arriba a la derecha: **Implementar** → **Nueva implementación**.

1. Al lado de "Selecciona el tipo", el icono del engranaje → **Aplicación web**
2. Descripción: `Reservas v1`
3. **Ejecutar como:** Yo (tu correo)
4. **Quién tiene acceso:** **Cualquier usuario** ← esto es lo importante, sin esto tus clientes no pueden entrar
5. **Implementar**

Te da una URL larga que termina en `/exec`. **Esa es tu página de reservas.** Cópiala y guárdala.

---

## Paso 8 — Probarlo tú mismo

Abre la URL en tu celular, en modo incógnito, y reserva una cita de prueba de principio a fin.

Verifica que:
- Aparezcan tus 17 servicios con precio y duración
- Al escoger un servicio largo, los cupos disponibles se reduzcan (un tinte de 2h10 no cabe a las 7 pm si cierras a las 8)
- La cita aparezca en la hoja **Reservas**
- Si pusiste tu correo, te llegue la confirmación

Después borra esa fila de prueba de la hoja Reservas.

---

## Paso 9 — Ponerle un link bonito

La URL de Google es larga y fea. Acórtala gratis para que quepa en la bio de Instagram:

- **Bitly** (bit.ly) o **acortar.link**: pegas la larga, te da `bit.ly/citas-riif`
- Si algún día compras un dominio (`riif.co`), lo rediriges a esa URL y queda `riif.co/citas`

Ese link corto va en: bio de Instagram, mensaje de bienvenida de WhatsApp Business, mensaje de ausencia, botón de Facebook, y en la vitrina del local con un código QR.

---

## Cómo se usa en el día a día

**Los clientes** entran al link, escogen servicio, barbero, día y hora, dejan sus datos y listo. La cita cae sola a la hoja Reservas. Si dejaron correo, les llega la confirmación con un enlace para cancelar solos.

**Tú** ves todo en la hoja Reservas en tiempo real, desde el celular.

**Para bloquear un rato** (almuerzo, cita médica, vacaciones): agrega una fila en la hoja **Bloqueos**. Ejemplo:

| Barbero | Fecha | Desde | Hasta | Motivo |
|---|---|---|---|---|
| Todos | 2026-08-20 | 12:00 | 13:00 | Almuerzo |
| Juan | 2026-08-22 | 09:00 | 20:00 | Día libre |

Escribe `Todos` para cerrar el local completo ese rato.

**Al cerrar el día:** en tu hoja de cálculo aparece un menú nuevo llamado **RIIF** → **Pasar citas de hoy al Registro**. Eso trae las citas del día a la hoja Registro con fecha, hora, barbero, cliente y servicio ya llenos. Tú (o el barbero) solo completa tres cosas: estado, método de pago y propina. De ahí en adelante el Cierre diario, las Comisiones y el Análisis se calculan solos.

---

## Lo que este sistema hace y lo que no

**Hace:** reservas 24/7 con tu propio link, cupos reales según la duración de cada servicio, sin doble reserva (tiene candado para dos clientes reservando al mismo tiempo), asignación automática al barbero con menos citas, confirmación por correo, cancelación por parte del cliente, bloqueos de agenda, y todo cae a tu hoja.

**No hace todavía, y se puede agregar después:**
- Recordatorio automático 24 horas antes (se hace con un disparador de tiempo en Apps Script)
- Pago o depósito anticipado
- Recordatorio por WhatsApp automático

**Sus límites reales:**
- Google permite unos 100 correos al día desde una cuenta gratuita. Muy por encima de lo que vas a usar.
- Si un cliente no deja correo, no recibe confirmación escrita. Por eso el celular sí es obligatorio: con eso le escribes por WhatsApp.
- El mantenimiento es tuyo. Es la contraparte de no pagar comisión ni depender de nadie.

---

## Si algo falla

| Síntoma | Causa casi siempre |
|---|---|
| No cargan los servicios | El nombre del archivo HTML no es exactamente `Reservar`, o Config no tiene precio y duración |
| No aparecen cupos ningún día | La hoja Horarios está vacía o las horas no están como `09:00` |
| Los clientes ven "necesitas permiso" | En la implementación quedó "Solo yo" en vez de "Cualquier usuario" |
| Cambié el código y no pasa nada | Hay que volver a implementar: **Implementar** → **Gestionar implementaciones** → editar → **Versión: Nueva** → Implementar |
