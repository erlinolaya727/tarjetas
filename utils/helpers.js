const novelty = (trabajador, registros, correoTrabajador, tablaInformativa) => {
  const subject = Params.EMAIL.Asunto;
  var body = '<html><body><table>';

  var htmlBody = `<p>Buen día,</p>
    <p>Señor(a): </p>
    <p>${Params.EMAIL.Gerente}</p>
    <p>${Params.EMAIL.Cargo}</p>
    <p>Me permito informarle que el trabajador [${trabajador}] ha cargado el Reporte de Tarjetas de observación con un total de [${registros}] registros.</p>`;

  for (var i = 0; i < tablaInformativa.length; i++) {
    var fila = tablaInformativa[i];
    body += '<tr>';

    // Establecer estilos de color para encabezados
    var colorFondo = i === 0 ? 'red' : 'pink';
    var colorLetra = i === 0 ? 'white' : 'black';

    // Generar las celdas de la fila con estilos de color
    for (var j = 0; j < fila.length; j++) {
      listaNombresEmpleadosTabla = fila[0].split(" ");
      nombreEmpleado = listaNombresEmpleadosTabla[0].toLowerCase();
      nombreTrabajador = trabajador.toLowerCase();
      if (!nombreTrabajador.includes(nombreEmpleado) && j == 13) {
        fila[j] = "NA";
        body += `<td style="background-color: ${colorFondo}; color: ${colorLetra};">${fila[j]}</td>`;
      } else {
        body += `<td style="background-color: ${colorFondo}; color: ${colorLetra};">${fila[j]}</td>`;
      }
    }
    body += '</tr>';
  }

  body += '</table></body></html>';

  htmlBody += body; // Agregar la tabla al cuerpo del correo

  htmlBody += `<p>Correo enviado automáticamente a través de App Script.</p>
    <p>Muchas gracias.</p>`;


  MailApp.sendEmail({
    to: Params.EMAIL.Destinatarios,
    cc: Params.EMAIL.GerenteGeneral,
    subject: subject,
    htmlBody: htmlBody,
  });

  var htmlBodyT = `<p>Buenos días,</p>
    <p>Señor(a): </p>    
    <p>${trabajador}</p>
    <p>Me permito informarle que las tarjetas de observación se han cargado de forma exitosa con con un total de [${registros}] registros. Su registro ha sido marcado como procesado, ver tabla siguiente:</p>`

  var body = '<html><body><table>';

  for (var i = 0; i < tablaInformativa.length; i++) {
    var fila = tablaInformativa[i];
    body += '<tr>';

    // Establecer estilos de color para encabezados
    var colorFondo = i === 0 ? 'red' : 'pink';
    var colorLetra = i === 0 ? 'white' : 'black';

    // Generar las celdas de la fila con estilos de color
    for (var j = 0; j < fila.length; j++) {
      listaNombresEmpleadosTabla = fila[0].split(" ");
      nombreEmpleado = listaNombresEmpleadosTabla[0].toLowerCase();
      nombreTrabajador = trabajador.toLowerCase();
      if (!nombreTrabajador.includes(nombreEmpleado) && j == 13) {
        fila[j] = "NA";
        body += `<td style="background-color: ${colorFondo}; color: ${colorLetra};">${fila[j]}</td>`;
      } else {
        body += `<td style="background-color: ${colorFondo}; color: ${colorLetra};">${fila[j]}</td>`;
      }
    }
    body += '</tr>';
  }

  body += '</table></body></html>';

  htmlBodyT += body; // Agregar la tabla al cuerpo del correo

  htmlBodyT += `<p>Correo enviado automáticamente a través de App Script.</p>
    <p>Correo enviado Automaticamente a través de  App Script. Desarrollado por Erlin Olaya´</p>  
    <p>Muchas gracias.´</p>`;


  MailApp.sendEmail(correoTrabajador, subject, '', {
    htmlBody: htmlBodyT,
  });
};

const criticalCheckPoint = (trabajador, correoTrabajador, error) => {
  // Se obtiene fecha de novedad

  const now = getCurrentDateTime();
  let subject = "";

  if (error == "Proceso ejecutado correctamente. EXITOSO") {
    subject = "Cargue EXITOSO de Tarjetas";
  } else {
    subject = "Cargue NO EXITOSO - FALLA";
  }

  if (error.includes("NO EXITOSO")) {
    subject = "Cargue NO EXITOSO - REGISTRO DUPLICADO";
  }

  const htmlBody = `<p>Buen día,</p>
  <p>Señor(a): </p>    
    <p>${trabajador}</p>
    <p>
    En el proceso <strong>Cargue Tarjetas Reporte</strong> se generó una 
    novedad en la fecha <strong>[${now[0]} ${now[1]}]</strong>, por tal motivo 
    no se logró procesar correctamente el archivo.
    </p>

    <p>Observaciones tecnicas (No obligatorias):</p>
    <p>${error}</p>
    `;
  MailApp.sendEmail(correoTrabajador, subject, '', {
    htmlBody: htmlBody,
  });
};

const getCurrentDateTime = () => {
  const now = new Date();
  const nowFormatted = now.toLocaleString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
  return nowFormatted.split(", ");
};



