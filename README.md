# Sistema de Gestion de Reservas y Sincronizacion con Google Calendar

Modulo de automatizacion desarrollado en Google Apps Script (JavaScript) para la gestion operativa y sincronizacion de reservas entre Google Sheets y Google Calendar.

El sistema elimina la duplicidad en el registro de datos, previene solapamientos de fechas mediante validacion previa de disponibilidad y proporciona control contable automatizado para cobros, abonos y convenios especiales.

---

## Contexto y Solucion

* **Problema:** El control manual de fechas, capacidad y estados de pago presentaba riesgos recurrentes de doble agendamiento en un mismo dia, inconsistencia en el registro de saldos y falta de trazabilidad en descuentos autorizados por administracion.
* **Solucion:** Arquitectura basada en un script en Google Apps Script conectado a la API de Calendar, complementado con formulas de matriz dinamicas (ARRAYFORMULA) en Sheets para sincronizar automaticamente el ciclo de vida de los eventos (creacion, actualizacion y cancelacion).

---

## Caracteristicas Tecnicas

* **Persistencia y Control de Estado:** Almacenamiento persistente del identificador unico de evento (UID devuelto por Calendar API) en una columna de control, permitiendo actualizar o eliminar eventos existentes sin generar registros duplicados.
* **Prevencion de Colisiones de Agenda:** Consulta preventiva de disponibilidad en el calendario mediante `CalendarApp.getEventsForDay()` antes de crear un nuevo evento. En caso de conflicto, el script detiene la insercion de esa fila y despliega una alerta modal interactiva en la interfaz.
* **Logica Financiera Reactiva:**
  * Calculo dinamico de costos segun categorias de asistentes (adultos, ninos y cobros extra).
  * Soporte para rebajas monetarias fijas o exencion total mediante token de convenio (`-`).
  * Estado de liquidacion condicional: asignacion automatica del estado `PAGADO` al liquidar el saldo total o registrar una exencion, evitando falsos positivos cuando los costos aun no han sido determinados.
* **Tolerancia a Datos Parciales:** Construccion dinamica de metadatos (titulo y cuerpo del evento), permitiendo registrar reservas que cuenten con fecha y entidad validas aun cuando no exista un responsable nominal asignado.

---

## Tecnologias Utilizadas

* **Lenguaje:** JavaScript (Apps Script runtime).
* **APIs de Google Workspace:**
  * `SpreadsheetApp` (Lectura por lotes en matrices bidimensionales y manipulacion de interfaz).
  * `CalendarApp` (Integracion y control de eventos en Google Calendar).
* **Motor de Hoja de Calculo:** Google Sheets (ARRAYFORMULA y formato condicional).

---

## Estructura de Datos (Google Sheets)

| Columna | Campo | Tipo | Descripcion |
| :---: | :--- | :--- | :--- |
| **A** | Responsable | Texto | Encargado de la reserva |
| **B** | Evento | Texto | Tipo de actividad programada |
| **C** | Organizacion / Iglesia | Texto | Entidad solicitante |
| **D** | Telefono | Texto | Contacto de referencia |
| **E** | Fecha | Date | Fecha agendada (`DD/MM/YYYY`) |
| **F** | Horario | Texto | Tramo horario reservado |
| **G - I** | Adultos / Ninos / Menores 4 | Numerico | Asistentes desglosados |
| **J** | Extras | Numerico | Servicios adicionales contratados |
| **K** | Descuento | Numerico / Texto | Monto a rebajar o `-` para exencion |
| **L** | Valor total | Formula | Costo bruto calculado |
| **M** | Abonado | Numerico | Monto ingresado |
| **N** | Por pagar | Formula | Saldo restante o indicador `PAGADO` |
| **O** | Estado | Texto | Estado de proceso (`Agendado`, `Cancelado`, `Conflicto`) |
| **P** | ID_Evento | String | UID generado por Calendar API (columna oculta) |

---

## Implementacion

1. Abrir el libro de trabajo en Google Sheets.
2. Acceder a **Extensiones > Apps Script**.
3. Incorporar el codigo en el editor (`sincronizar_reservas.js`).
4. Guardar el proyecto y autorizar los permisos requeridos de Google Workspace en la primera ejecucion.
5. Asignar la funcion `sincronizarPeriquillo` a un control visual (boton) en la hoja de calculo.
