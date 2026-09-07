//PARA CONEXION CON LA HOJA Y RECOLECCION DE DATOS
function sincronizarPeriquillo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 3; 
  var lastRow = sheet.getLastRow();
  
  if (lastRow < startRow) return;

  var range = sheet.getRange(startRow, 1, lastRow - startRow + 1, 16);
  var data = range.getValues();
  var calendar = CalendarApp.getDefaultCalendar();
  var conflictos = [];

  //RECORRIDO DE LAS FILAS
  for (var i = 0; i < data.length; i++) {
    var fila = data[i];
    var filaIndex = startRow + i;

    var cliente     = fila[0] ? fila[0].toString().trim() : "";
    var tipoEvento  = fila[1] ? fila[1].toString().trim() : "Evento";
    var iglesia     = fila[2] ? fila[2].toString().trim() : "";
    var telefono    = fila[3] ? fila[3].toString().trim() : "Por confirmar";
    var fecha       = new Date(fila[4]);
    var horario     = fila[5] ? fila[5].toString().trim() : "Por definir";
    var adultos     = fila[6] !== "" ? fila[6] : "-";
    var ninos       = fila[7] !== "" ? fila[7] : "-";
    var menores     = fila[8] !== "" ? fila[8] : "-";
    var extras      = fila[9] !== "" ? fila[9].toString().trim() : "0";
    var descuento   = fila[10] !== "" ? fila[10].toString().trim() : "0";
    var total       = fila[11] !== "" ? fila[11] : 0;
    var abono       = fila[12] !== "" ? fila[12] : 0;
    var saldo       = fila[13] !== "" ? fila[13].toString().trim() : ""; 
    var estado      = fila[14] ? fila[14].toString().trim() : "";
    var eventoId    = fila[15] ? fila[15].toString().trim() : ""; 

    //CONTEO DE ASISTENTES 
    var numAdultos = Number(fila[6]) || 0;
    var numNinos = Number(fila[7]) || 0;
    var numMenores = Number(fila[8]) || 0;
    var totalPersonas = numAdultos + numNinos + numMenores;
    var detallePersonas = totalPersonas > 0 ? totalPersonas + " pers." : "Asistencia por confirmar";

    //LÓGICA PARA EL ESTADO DE PAGO:
    //Solo es "PAGADO" si la columna "Por pagar" dice explícitamente "PAGADO" o si se marcó "-"
    var saldoUpper = saldo.toUpperCase();
    var textoEstadoPago = "PENDIENTE";
    if (saldoUpper === "PAGADO" || descuento === "-") {
      textoEstadoPago = "PAGADO";
    } else if (Number(total) === 0) {
      textoEstadoPago = "POR DEFINIR";
    }

    //CANCELAR / ELIMINAR
    var est = estado.toLowerCase();
    if ((est === "cancelar" || est === "cancelado" || est === "eliminar" || est === "borrar") && eventoId) {
      try {
        var evento = calendar.getEventById(eventoId);
        if (evento) evento.deleteEvent();
      } catch (e) {}
      sheet.getRange(filaIndex, 15).setValue("Cancelado");
      sheet.getRange(filaIndex, 16).setValue("");
      continue;
    }

    //PREVENCION DE COLISIONES EN FECHAS! 
    //CREAR / EDITAR
    if (!isNaN(fecha.getTime()) && (cliente !== "" || tipoEvento !== "" || iglesia !== "") && est !== "cancelado") {
      
      if (!eventoId) {
        var eventosDelDia = calendar.getEventsForDay(fecha);
        if (eventosDelDia.length > 0) {
          sheet.getRange(filaIndex, 15).setValue("Conflicto: Fecha ocupada en Calendar");
          conflictos.push("Fila " + filaIndex + " (" + (tipoEvento || cliente) + ") el " + fecha.toLocaleDateString("es-CL"));
          continue;
        }
      }

    //FORMATO PARA DESCRIPCION EN CALENDAR 
      var partesTitulo = [tipoEvento];
      if (cliente) partesTitulo.push(cliente);
      if (iglesia) partesTitulo.push("(" + iglesia + ")");
      partesTitulo.push("(" + detallePersonas + ")");
      
      var titulo = partesTitulo.join(" - ");

      var textoDescuento = (descuento !== "0" && descuento !== "") ? "- Descuento / Convenio: " + (isNaN(Number(descuento)) ? descuento : "$" + Number(descuento).toLocaleString("es-CL")) + "\n" : "";

      var saldoMostrar = "$0";
      if (saldoUpper === "PAGADO") {
        saldoMostrar = "$0 (PAGADO)";
      } else if (!isNaN(Number(saldo)) && Number(saldo) > 0) {
        saldoMostrar = "$" + Number(saldo).toLocaleString("es-CL");
      }

      var descripcion = 
        "DETALLES DE LA RESERVA\n" +
        "-----------------------------------------\n" +
        "Tipo de Evento: " + tipoEvento + "\n" +
        "Iglesia: " + (iglesia || "No especificada") + "\n" +
        "Responsable: " + (cliente || "Por confirmar") + "\n" +
        "Telefono: " + telefono + "\n" +
        "Horario: " + horario + "\n\n" +
        "ASISTENTES:\n" +
        "- Adultos: " + adultos + "\n" +
        "- Ninos: " + ninos + "\n" +
        "- Menores de 4: " + menores + "\n" +
        "- Servicios extra: " + extras + "\n\n" +
        "ESTADO DE PAGO (" + textoEstadoPago + "):\n" +
        textoDescuento +
        "- Total: " + (Number(total) > 0 ? "$" + Number(total).toLocaleString("es-CL") : "Por definir") + "\n" +
        "- Abonado: " + (Number(abono) > 0 ? "$" + Number(abono).toLocaleString("es-CL") : "$0") + "\n" +
        "- Saldo Pendiente: " + saldoMostrar;

      //SINCRONIZACION CON CALENDAR (EDITAR / CREAR)
      if (eventoId) {
        try {
          var ev = calendar.getEventById(eventoId);
          if (ev) {
            ev.setTitle(titulo);
            ev.setDescription(descripcion);
            ev.setAllDayDate(fecha);
          }
        } catch (e) {
          eventoId = "";
        }
      }

      if (!eventoId) {
        var nuevoEv = calendar.createAllDayEvent(titulo, fecha, {
          description: descripcion
        });
        sheet.getRange(filaIndex, 15).setValue("Agendado");
        sheet.getRange(filaIndex, 16).setValue(nuevoEv.getId());
      }
    }
  }

  //ALERTA VISUAL DE COLISIONES 
  if (conflictos.length > 0) {
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      "Atencion: Tope de fechas detectado",
      "No se agendaron las siguientes filas porque ya existe una actividad en Google Calendar para ese dia:\n\n" +
      conflictos.join("\n") +
      "\n\nPor favor revisa el calendario antes de continuar.",
      ui.ButtonSet.OK
    );
  }
}
