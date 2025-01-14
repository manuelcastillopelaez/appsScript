//Constantes
var ID = 1,
  FECHAT = 2,
  SDATOOL = 3,
  PROYECTO = 4,
  LINKTICKETJIRA = 5,
  DESCRIPCION = 6,
  FASE_DE_PROYECTO = 7,
  RELEASETENTATIVA = 8,
  REGISTRANTE = 9,
  ESTADO = 10,
  COUNTEROBSERVED = 11;
OBSERVACION = 12,
  DATEACEPTED = 13;
HISTORICO = 14;


const teamQe = [
  "mcastillop@bbva.com",
  "julio.arteaga.carhua@bbva.com",
  "msalazarh@bbva.com",
  "rmiranda@bbva.com",
]
var emails = [
  "jose.espinoza.delgado@bbva.com",
  "jesus.vasquez@bbva.com",
  "ealarcon@bbva.com",
  "glomo-channel-management-pe.group@bbva.com"
].concat(teamQe);

const faseDeProyecto = {
  consulta: {
    value: 'En proceso, consulta de requisitos',
    estadosAdmitidos: ["PENDIENTE", "AGENDADO", "EN PROCESO", "ATENDIDO"],
  },
  terminado: {
    value: 'Terminado validación de requisitos',
    estadosAdmitidos: ["PENDIENTE", "AGENDADO", "EN PROCESO", "OBSERVADO", "ACEPTADO"],
  }
}

function Hora() {
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth() + 1;
  var day = now.getUTCDate();
  var hora = now.getHours();
  var minutos = now.getMinutes();

  var string = day + "/" + month + "/" + year + "-" + hora + ":" + minutos;

  return string;
}
function envioEmail(e) {
  //Detectar consecutivo
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("formulario1");
  //Obtener la ultima fila
  var ultimaFila = hoja.getLastRow();
  var idConsecutivo = ultimaFila - 1;
  //colocar valor
  hoja.getRange(ultimaFila, ID).setValue("MB100" + idConsecutivo);
  hoja.getRange(ultimaFila, ESTADO).setValue("PENDIENTE");
  hoja.getRange(ultimaFila, COUNTEROBSERVED).setValue(0);
  var requester = hoja.getRange(ultimaFila, REGISTRANTE).getValue();
  var subject = "Solicitud comite Release Glomo Id: " + "MB100" + idConsecutivo;
  var cc = emails.join(",");
  var bcc = "";


  //Get Metadata
  var fecha = Hora();
  var id = "MB100" + idConsecutivo;
  var registrante = hoja.getRange(ultimaFila, REGISTRANTE).getValue();
  var sdaTool = hoja.getRange(ultimaFila, SDATOOL).getValue();
  var linkTicketJira = hoja.getRange(ultimaFila, LINKTICKETJIRA).getValue();
  var proyecto = hoja.getRange(ultimaFila, PROYECTO).getValue();
  var descripcion = hoja.getRange(ultimaFila, DESCRIPCION).getValue();
  var releaseTentativa = hoja.getRange(ultimaFila, RELEASETENTATIVA).getValue();
  var estadoAsignado = hoja.getRange(ultimaFila, ESTADO).getValue();
  var observacion = hoja.getRange(ultimaFila, OBSERVACION).getValue();
  const ticket = GetTicketGlomo(linkTicketJira);
  if (!ticket.includes('GLOMOPE')) {
    estadoAsignado = 'RECHAZADO';
    observacion = `El link: ${linkTicketJira} no tiene el formato establecido. Los tickets tienen que venir del WorkSpace GLOMOPE.`;
    hoja.deleteRow(ultimaFila);
  }
  /* else {
    const url = CreateSheet(ticket, sdaTool, proyecto, releaseTentativa, registrante, [registrante], teamQe);
    hoja.getRange(ultimaFila, OBSERVACION).setValue(`El checklist es: ${url}`);
  }*/
  var messageCompleteHtml = formatHtmlMessage(fecha, id, registrante, proyecto, descripcion, estadoAsignado, observacion);
  GmailApp.sendEmail(requester, subject, 'html body', { cc: cc, bcc: bcc, htmlBody: messageCompleteHtml });
}

function formatHtmlMessage(fecha, id, registrante, proyecto, descripcion, estadoAsignado, observacion, mensajeAgenda) {

  var template = HtmlService.createTemplateFromFile("mailTemplate");

  var colorStatus = '';

  if (estadoAsignado == "") {
    estadoAsignado = 'PENDIENTE';
  }

  switch (estadoAsignado) {
    case 'PENDIENTE': colorStatus = '#cb3234'; break;
    case 'AGENDADO': colorStatus = '#fbc72e'; break;
    case 'EN PROCESO': colorStatus = '#fbc72e'; break;
    case 'OBSERVADO': case 'RECHAZADO': colorStatus = '#cb3234'; break;
    case 'ACEPTADO': case 'ATENDIDO': colorStatus = '#1daf5e'; break;
    default: colorStatus = '#fbc72e'; break;
  }

  template.fecha = fecha;
  template.id = id;
  template.registrante = registrante;
  template.proyecto = proyecto;
  template.descripcion = descripcion;
  template.estadoAsignado = estadoAsignado;
  template.colorStatus = colorStatus;
  template.observacion = observacion;
  template.mensajeAgenda = mensajeAgenda;
  return template.evaluate().getContent();
}

function generaId() {
  //Detectar hoja
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("formulario1");
  //Obtener la ultima fila
  var ultimaFila = hoja.getLastRow();
  var idConsecutivo = ultimaFila - 1;
  //colocar valor
  hoja.getRange(ultimaFila, 1).setValue("MB100" + idConsecutivo);
  return "MB100" + idConsecutivo;
}

function validaEstado() {
  //Detectar hoja
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("formulario1");
  //Obtener la fila activa
  var rangoActivo = hoja.getActiveRange();
  var ultimaFila = rangoActivo.getRow();
  var columna = rangoActivo.getColumn();
  var active_user = Session.getActiveUser().getEmail();
  var requester = hoja.getRange(ultimaFila, REGISTRANTE).getValue();
  var valorAnterior = hoja.getRange(ultimaFila, HISTORICO).getValue();
  var counterObserved = hoja.getRange(ultimaFila, COUNTEROBSERVED).getValue();
  var subject = "Cambio de estado - solicitud de comite Release Glomo " + hoja.getRange(ultimaFila, ID).getValue();

  var cc = emails.join(',');
  var bcc = "";
  var mensajeAgenda = '';

  if (columna == ESTADO) {
    if (hoja.getRange(ultimaFila, FASE_DE_PROYECTO).getValue() === faseDeProyecto.consulta.value && !(faseDeProyecto.consulta.estadosAdmitidos.includes(hoja.getRange(ultimaFila, ESTADO).getValue()))) {
      hoja.getRange(ultimaFila, ESTADO).setValue("PENDIENTE");
      Browser.msgBox(`Este es un estado no valido para esta fase del proyecto, solo se admite: ` + faseDeProyecto.consulta.estadosAdmitidos.join(', '));
    }
    else if (hoja.getRange(ultimaFila, FASE_DE_PROYECTO).getValue() === faseDeProyecto.terminado.value && !(faseDeProyecto.terminado.estadosAdmitidos.includes(hoja.getRange(ultimaFila, ESTADO).getValue()))) {
      hoja.getRange(ultimaFila, ESTADO).setValue("PENDIENTE");
      Browser.msgBox(`Este es un estado no valido para esta fase del proyecto, solo se admite: ` + faseDeProyecto.terminado.estadosAdmitidos.join(', '));
    }
    else if (hoja.getRange(ultimaFila, ESTADO).getValue() == "OBSERVADO" && hoja.getRange(ultimaFila, OBSERVACION).getValue() == "") {
      hoja.getRange(ultimaFila, ESTADO).setValue("EN PROCESO");
      hoja.getRange(ultimaFila, HISTORICO).setValue(valorAnterior + "/" + active_user + " intento cambiar el estado OBSERVADO - " + Hora());
      Browser.msgBox("Debe ingresar una observación al colocar el estado OBSERVADO, se devuelve el estado a EN PROCESO");
    }
    //else if (["EN PROCESO","AGENDADO"].includes(hoja.getRange(ultimaFila, ESTADO).getValue()) && active_user != requester && !teamQe.includes(active_user)) {
    else if (["EN PROCESO", "AGENDADO"].includes(hoja.getRange(ultimaFila, ESTADO).getValue()) && !teamQe.includes(active_user)) {
      const status = hoja.getRange(ultimaFila, ESTADO).getValue();
      hoja.getRange(ultimaFila, ESTADO).setValue("PENDIENTE");
      hoja.getRange(ultimaFila, HISTORICO).setValue(valorAnterior + "/" + active_user + " intento cambiar el estado a EN PROCESO - " + Hora());
      Browser.msgBox(`Solo QE puede cambiar el estado a ${status}, se devuelve el estado a pendiente`);
    }
    else if (hoja.getRange(ultimaFila, ESTADO).getValue() == "ACEPTADO" && !teamQe.includes(active_user)) {
      hoja.getRange(ultimaFila, ESTADO).setValue("PENDIENTE");
      hoja.getRange(ultimaFila, HISTORICO).setValue(valorAnterior + "/" + active_user + " intento cambiar el estado a ACEPTADO - " + Hora());
      Browser.msgBox("Solo QE puede cambiar el estado a ACEPTADO, se devuelve el estado a pendiente");
    }

    else {
      if (hoja.getRange(ultimaFila, ESTADO).getValue() == "ACEPTADO" && teamQe.includes(active_user)) {
        hoja.getRange(ultimaFila, DATEACEPTED).setValue(Hora() + "- Aceptado por: " + active_user);
      }
      if (hoja.getRange(ultimaFila, ESTADO).getValue() == "OBSERVADO") {
        hoja.getRange(ultimaFila, COUNTEROBSERVED).setValue(counterObserved + 1);
      }
      if (hoja.getRange(ultimaFila, ESTADO).getValue() == "AGENDADO") {
        mensajeAgenda = 'Su solicitud fue agendada para la próxima sesión de comité. ';
        if (hoja.getRange(ultimaFila, FASE_DE_PROYECTO).getValue() == faseDeProyecto.terminado.value) {
          mensajeAgenda.concat('Es obligatoria la presencia del PO, SM, Developer y QA.');
        }
      }
      //Get Metadata
      var fecha = Hora();
      var id = hoja.getRange(ultimaFila, ID).getValue();
      var registrante = hoja.getRange(ultimaFila, REGISTRANTE).getValue();
      var proyecto = hoja.getRange(ultimaFila, PROYECTO).getValue();
      var descripcion = hoja.getRange(ultimaFila, DESCRIPCION).getValue();
      var estadoAsignado = hoja.getRange(ultimaFila, ESTADO).getValue();
      var observacion = hoja.getRange(ultimaFila, OBSERVACION).getValue();
      var messageCompleteHtml = formatHtmlMessage(fecha, id, registrante, proyecto, descripcion, estadoAsignado, observacion, mensajeAgenda);
      GmailApp.sendEmail(requester, subject, 'html body', { cc: cc, bcc: bcc, htmlBody: messageCompleteHtml });
    }
  }
}

