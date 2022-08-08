// The main function for preparing sending emails
function sendEmails() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Responsables");

  var date = dateInSpanish();

  // Start of range of our data
  var startRow = 3; // Row 3
  var startColumn = 2; // Column B

  // End of range of our data
  var numRows = sheet.getLastRow() - startRow + 1; // Expression to obtain the number of rows with actual data
  var numColumns = 4; // Number of columns in which our data is contained.

  // Define boundaries of the data table
  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns);
  var data = dataRange.getValues();

  var tl_names = []; // This list is going to be filled within the loop with the names of the people that were sent an automatic email

  for (var i in data) {
    var curRow = data[i]; // Current working row in the loop

    // Fetch the information 
    var area = curRow[0]; // Column B
    var name = curRow[1]; // Column C
    var emailAdress = curRow[2]; // Column D
    var link = curRow[3]; // Column E    

    Logger.log(name + " " + area + " " + emailAdress + " " + link);

    if (!tl_names.includes(name)) tl_names.push(name); // If the name is not already present, add the name to the list

    // Actual code to send the email. We want to perform this action only if there's text on the var, otherwise, won't execute
    if (emailAdress.length > 0) {

      // Function to create the body of our email.
      var htmlBody = forgeBodyTLs(name, area, link, date);

      MailApp.sendEmail({
        to: emailAdress,
        subject: "Recordatorio: actualizar mapa de cubículos - " + date,
        htmlBody: htmlBody
      });
    }
  }

  // Last step is to notify the PMs of the email sent.
  sendEmailPMs(tl_names, date);
}

function dateInSpanish() {
  var date = Utilities.formatDate(new Date(), "GMT-7", "MMMM d, yyyy").toString(); // Date that this email is being sent, in America/Tijuana Timezone.
  var dateEs = LanguageApp.translate(date, "en", "es");
  return dateEs;
}

function sendEmailPMs(tl_names, date) {
  // Define PMs emails to send email after
  var recipientsPM =
  "felipe.puentes@trisourcebpo.com," +
  "francisco.arredondo@trisourcebpo.com,"+
  "erick.duran@trisourcebpo.com," + 
  "gilberto.gamino@trisourcebpo.com," +
  "carlos.hernandez@trisourcebpo.com," +
  "jesus.leyva@trisourcebpo.com";

  // People that should be CCed in the after-email
  var emailCC =
  "emilio@trisourcebpo.com," +
  "alfredo.ruiz@trisourcebpo.com," +
  "pamela.martinez@trisourcebpo.com," +
  "julian.lamadrid@trisourcebpo.com," +
  "fabian.gonzalez@trisourcebpo.com";

  var htmlBody2 = forgeBodyPMs(tl_names, date);

  MailApp.sendEmail({
    to: recipientsPM,
    cc: emailCC,
    subject: "Recordatorio: actualizar mapa de cubículos - " + date,
    htmlBody: htmlBody2
  });
}

function forgeBodyPMs(names, date) {
  var htmlNamesList = "";
  for (var i in names) {
    htmlNamesList += "<li>" + names[i] + "</li>";
  }
  var htmlBody = "<div style=\"max-width:640px;\"><font size=\"+2\"><b>Recordatorio: actualizar mapa de cubículos - " + date + "</b></font> <br><br>Buenos dias a todos!<br><br> El propósito del presente es informarles que se envió un correo de notificación automático a los siguientes TL para actualizar el mapa de cubículos: <ul>" + htmlNamesList + "</ul> Al enviar el presente, esperamos que anime a los TL bajo su supervisión a cumplir con la tarea asignada, ya que esta es crucial para el buen funcionamiento de la empresa y la agilización al momento de planificar e incorporar nuevos elementos al equipo TRI Source. <br><br> Sin mas que agregar por el momento, quedamos abiertos a comentarios y dudas que puedan surgir al respecto.<br><br> <i>Atentamente, el equipo de IT de TRI Source TJ.</i></div>";  

  return htmlBody;
}

function forgeBodyTLs(name, area, link, date) {
  var htmlBody = "<div style=\"max-width:640px;\"><font size=\"+2\"><b>Recordatorio: actualizar mapa de cubículos - " + date + "</b></font> <br><br> Buenos días, <b>" + name + "</b>.<br><br>Este es un recordatorio para que actualice el mapa de cubículos de su área designada, es decir, área <b>" + area + "</b>. <br><br><b><a href=\"" + link + "\">Haga clic aquí para ser redirigido a su área designada (" + area + ") en el archivo del mapa de cubículos.</a></b> <br><br><i>Tenga en cuenta que esta actualización generalmente está programada para realizarse los <b>jueves</b>. Sin embargo, si tiene la oportunidad de actualizar el mapa tan pronto como se dé cuenta de que algo ha cambiado, le recomendamos que lo haga.</i> <br><br>Además, tenga en cuenta que <b>este procedimiento es de suma importancia</b>, ya que actualmente estamos experimentando un crecimiento acelerado en la empresa y es crucial tener todas nuestras áreas y equipos disponibles listos para la expansión. <br><br>Para actualizar el estado de un cubículo, simplemente ubique la etiqueta del cubículo (generalmente una letra y un número, por ejemplo, <b>E4</b>), luego ubique el código en el archivo de mapa y haga clic en la flecha hacia abajo a la derecha del celda correspondiente al cubículo.<br>Elija una de las siguientes opciones: <br><br> </div> <table cellspacing=\"0\" cellpadding=\"0\" dir=\"ltr\" border=\"1\" style=\"table-layout:fixed;;width:0px;border-collapse:collapse;border:none\"> <colgroup> <col width=\"102\"> <col width=\"333\"> </colgroup> <tbody> <tr style=\"height:21px\"> <td style=\"border:1px solid rgb(0,0,0);overflow:hidden;padding:2px 3px;vertical-align:bottom;background-color:rgb(183,225,205)\">Assigned</td> <td style=\"overflow:hidden;padding:2px 3px;vertical-align:bottom;border:1px solid rgb(204,204,204)\">Cubículo con computadora y agente.</td> </tr> <tr style=\"height:21px\"> <td style=\"border-width:1px;border-style:solid;border-color:rgb(204,204,204) rgb(0,0,0) rgb(0,0,0);overflow:hidden;padding:2px 3px;vertical-align:bottom;background-color:rgb(159,197,232)\">Unassigned</td> <td style=\"overflow:hidden;padding:2px 3px;vertical-align:bottom;border:1px solid rgb(204,204,204)\">Cubículo con computadora pero sin agente.</td> </tr> <tr style=\"height:21px\"> <td style=\"border-width:1px;border-style:solid;border-color:rgb(204,204,204) rgb(0,0,0) rgb(0,0,0);overflow:hidden;padding:2px 3px;vertical-align:bottom;background-color:rgb(204,204,204)\">Empty</td> <td style=\"overflow:hidden;padding:2px 3px;vertical-align:bottom;border:1px solid rgb(204,204,204)\">Cubículo sin computadora.</td> </tr> <tr style=\"height:21px\"> <td style=\"border-width:1px;border-style:solid;border-color:rgb(204,204,204) rgb(0,0,0) rgb(0,0,0);overflow:hidden;padding:2px 3px;vertical-align:bottom;background-color:rgb(234,153,153)\">Reserved</td> <td style=\"overflow:hidden;padding:2px 3px;vertical-align:bottom;border:1px solid rgb(204,204,204)\">Computadora reservada para nuevas clases programadas</td> </tr> <tr style=\"height:21px\"><td style=\"border-width:1px;border-style:solid;border-color:rgb(204,204,204) rgb(0,0,0) rgb(0,0,0);overflow:hidden;padding:2px 3px;vertical-align:bottom;background-color:rgb(180,167,214)\">DS</td> <td style=\"overflow:hidden;padding:2px 3px;vertical-align:bottom;border:1px solid rgb(204,204,204)\">PC especial para anuncios de TV.</td> </tr> </tbody> </table> <br><br><p><font size=\"+1\"><i>Atentamente, el equipo de IT de TRI Source TJ.</i></font></p> <br><br><br><br> <div style=\"max-width:640px;\"><i>Si tiene problemas con esta tarea, o si cree que este correo electrónico no es para usted, no dude en hacérnoslo saber enviando un ticket a nuestro servicio de asistencia, envíenos un correo electrónico indicando el problema que está enfrentando actualmente y todos los detalles adicionales que considere necesarios para ayudarnos a resolver el problema.</i> <br><br> <b><a href=\"mailto:support.mx@trisourcebpo.com\">¡Enviar un ticket! - support.mx@trisourcebpo.com</a></b></div>"
  return htmlBody;
}
