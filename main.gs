//FormForm 3.0 Installation. Steps:
//1. Change Drive templates/folders IDs in env.gs
//2. Run function obtenerPermisos() from Apps Script
//3. Create a new activator for the onFormSubmit function, whenever a form is submitted 

function obtenerPermisos(){
  MailApp.getRemainingDailyQuota();
  DriveApp.getStorageLimit();
  FormApp.getActiveForm();
  SpreadsheetApp.flush();
  CalendarApp.getTimeZone();
}

function onFormSubmit(e) {

  // 1. Obtenemos la información del FormForm
  var respuestas = e.response.getItemResponses();

  // 2. Asignamos cada respuesta del formulario a una variable
  var [grupo, titulo, descripcion, responsables, fechaPura, tipoActividad, hora, ubicacion, precio, dl_entradas, material, info, publicacion] = respuestas.map(function(f) {return f.getResponse()});
  var email = e.response.getRespondentEmail();

  //Formateamos la fecha
  var fecha = formatearFecha(fechaPura);

  //Obtenemos las etiquetas de las imagenes para asignar una imagen ideal a la historia
  cargarEtiquetas();

  // 3. Generamos la historia de instagram y la guardamos en un archivo de imagen
  var imagen = generarImagenHistoria(titulo, grupo, hora, fecha, ubicacion, fechaPura);

  // 4. Generamos los mensajes de Whatsapp en español e inglés
  var mensajes = generarMensajesWhatsapp(titulo, descripcion, fecha, hora, ubicacion, precio, dl_entradas, material, info);

  // 5. Enviar la historia y los mensajes al correo del coordi
  enviarCorreo(email, imagen, mensajes);
}

function generarImagenHistoria(titulo, grupoTrabajo, hora, fecha, ubicacion, fechaPura, debug = false){
  var carpeta = DriveApp.getFolderById(ID_CARPETA);
  var plantilla = DriveApp.getFileById(ID_PLANTILLA);

  var copiaPlantilla = plantilla.makeCopy();
  var idCopiaPlantilla = copiaPlantilla.getId();
  var presentacion = SlidesApp.openById(idCopiaPlantilla);
  var diapositiva = presentacion.getSlides()[0];

  var cadenaFecha = fecha[1].toUpperCase() + " " + fecha[2] + "" + fecha[3].toUpperCase();
  var color = COLORES_GRUPOS[grupoTrabajo];
  var imagen = buscarImagen(titulo, ubicacion);
  
  if(imagen){
    diapositiva.getPageElements()[0].asImage().replace(imagen);
  }

  diapositiva.getPageElements()[1].asShape().getFill().setSolidFill(color,0.4);
  diapositiva.getPageElements()[3].asShape().getText().getTextStyle().setFontSize(calcularTamFuente(titulo.length + 15, 1, 72));
  diapositiva.getPageElements()[5].asShape().getText().getTextStyle().setFontSize(calcularTamFuente(cadenaFecha.length, 0.1, 60));
  diapositiva.getPageElements()[6].asShape().getText().getTextStyle().setFontSize(calcularTamFuente(ubicacion.length + 15, 1, 60));

  diapositiva.replaceAllText("{{nombre_actividad}}", titulo.toUpperCase());
  diapositiva.replaceAllText("{{hora}}", hora);
  diapositiva.replaceAllText("{{fecha}}", cadenaFecha);
  diapositiva.replaceAllText("{{ubicacion}}", ubicacion.toUpperCase());

  diapositiva.replaceAllText("\n", "");

  presentacion.saveAndClose();

  const url = Slides.Presentations.Pages.getThumbnail(idCopiaPlantilla, diapositiva.getObjectId(), {"thumbnailProperties.mimeType": "PNG"}).contentUrl;
  const blob = UrlFetchApp.fetch(url).getAs(MimeType.PNG);

  var fechaSemana = obtenerLunes(fechaPura);
  var nombreCarpetaSemanal = debug?"Prueba2.0":"Semana " + fechaSemana;
  var diaSemana = obtenerDiaSemana(fechaPura);
  var letra = "A";//tipoActividad === "Diurna"? 'A' : 'B';

  var carpetaSemanal = checkIfFolderExistElseCreate(carpeta, nombreCarpetaSemanal);

  var archivo = carpetaSemanal.createFile(blob.setName(`${diaSemana}${letra}_${fecha}_${titulo}.png`));

  var carpetaEditables = checkIfFolderExistElseCreate(carpeta, "Editables");
  copiaPlantilla.setName(`${diaSemana}${letra}_${fecha}_${titulo}`);
  copiaPlantilla.moveTo(carpetaEditables);

  return archivo;
}

function buscarImagen(titulo, ubicacion){
  var keywords = (titulo + " " + ubicacion).toLowerCase().split(/[\s,\/]+/);
  var mejoresImagenes = [];
  var maximoCoincidencias = 0; 

  for(var idImagen in ETIQUETAS_IMAGENES) {
    var labels = ETIQUETAS_IMAGENES[idImagen].split(/[\s,]+/);
    var coincidencias = keywords.filter(v => labels.includes(v)).length;

    if(coincidencias > 0){
      if(coincidencias > maximoCoincidencias){
        mejoresImagenes = [];
        maximoCoincidencias = coincidencias;
        mejoresImagenes.push(idImagen);
      }else if(coincidencias === maximoCoincidencias){
        mejoresImagenes.push(idImagen);
      }
    }
  }

  const randomIndex = Math.floor(Math.random() * mejoresImagenes.length);
  const imageFile = mejoresImagenes.length !== 0 ? DriveApp.getFileById(mejoresImagenes[randomIndex]) : null;

  return imageFile;
}

function cargarEtiquetas(){
  var files = DriveApp.getFolderById(ID_CARPETA_IMAGENES).getFiles();
  
  while(files.hasNext()){
    var image = files.next();

    var idImagen = image.getId();
    var descripcion = image.getDescription();

    ETIQUETAS_IMAGENES[idImagen] = descripcion;
  }
}

function generarMensajesWhatsapp(titulo, descripcion, fecha, hora, ubicacion, precio, dl_entradas, material, info){
  var mensajesWP = [];

  for (var i = 0; i < 2; i++){
    var mensaje = (i==0?`¡Hola a todo el mundo! Os informamos de que dentro de poco realizaremos la siguiente actividad:` : `Hello everyone! We inform you that we will carry out soon the following activity:`) + "\n\n";
    
    if(i==0){
      mensaje += `📢 *${LanguageApp.translate(fecha[0], 'en', 'es').toUpperCase() + " " + fecha[2].toUpperCase() + " DE " + LanguageApp.translate(fecha[1], 'en', 'es').toUpperCase()} ${titulo}*\n`;
    }else{
      mensaje += `📢 *${fecha[0].toUpperCase() + " " + fecha[1].toUpperCase() + " " + fecha[2] + "" + fecha[3]} ${titulo}*\n`;
    }

    if(descripcion!==""){
      mensaje += `📝 ${i == 0 ? descripcion : LanguageApp.translate(descripcion, 'es', 'en')}\n`;
    }

    mensaje += `🕛 ${hora}\n`;
    mensaje += `📍 ${ubicacion}\n`;

    if(precio!=="GRATIS"){
      mensaje += `💰 ${i == 0 ? precio : LanguageApp.translate(precio, 'es', 'en')}` + "\n";
    }

    if(material!==""){
      mensaje += `🎒 ${i == 0 ? material : LanguageApp.translate(material, 'es', 'en')}` + "\n";
    }

    if(info!==""){
      mensaje += `ℹ️ ${i == 0 ? info : LanguageApp.translate(info, 'es', 'en')}\n`;
    }

    mensaje += (i==0?`¡Nos vemos pronto!😊`: `See you soon!😊`);

    mensajesWP.push(mensaje);
  }

  return mensajesWP;
}

function formatearFecha(fecha){
  var objetoFecha = new Date(fecha);
  var timeZone = CalendarApp.getTimeZone();
  var diaSemana = Utilities.formatDate(objetoFecha, timeZone, "EEEE");
  var mes = Utilities.formatDate(objetoFecha, timeZone, "MMMM");
  var dia = Utilities.formatDate(objetoFecha, timeZone, "d");
  var sufijo;

  //Averiguar sufijo en inglés
  //Si el dia del mes está entre el 10 y el 19 sabemos que el sufijo siempre será 'th', por lo que no será necesario entrar en el switch
  if(!(dia.length === 2 && dia.charAt(0) == 1)){
    switch(dia.charAt(dia.length - 1)){
      case '1':
      sufijo = "st";
      break;

      case '2':
      sufijo = "nd";
      break;

      case '3':
      sufijo = "rd";
      break;

      default:
      sufijo = "th";
      break;
    }
  }else{
    sufijo = "th";
  }

  return [diaSemana, mes, dia, sufijo];
}

function enviarCorreo(email, imagen, mensajes){
  MailApp.sendEmail(email, "¡Listo! FormFormBot ya ha preparado tu actividad.",
  "🤖🔧Beep beep boop... aquí tienes tus recursos, úsalos sabiamente...\n\nESPAÑOL:\n\n" + mensajes[0] + "\n\nINGLÉS:\n\n" + mensajes[1] + "\n\nAdemás, encontrarás la historia generada a partir del formulario adjunta a este correo.\nRecuerda que tienes total libertad para personalizar los mensajes tanto como quieras, esta es únicamente una referencia para facilitaros el trabajo a los coordis.\n\n¡Un saludo, y feliz coordinación!\n\nAtentamente,\nEl FormFormBot de ESN Granada.",
  {
    attachments:[imagen]
  });
}

function calcularTamFuente(longitud, factor_escalado, tam_max){
  var tam;

  factor_escalado = (factor_escalado > 1? 1 : (factor_escalado < 0? 0 : factor_escalado));

  //Cuanto más se acerque a 0 el coeficiente negativo, más grande tiende a ser el texto
  tam = Math.exp(-0.08 * (2 - factor_escalado) * longitud + 6) + 40;
  tam = Math.min(tam, tam_max);

  return tam;
}

function obtenerDiaSemana(fecha){
  var curr = new Date(fecha);
  var weekDay = curr.getDay();

  //Como la cuenta empieza en domingo(0) y acaba en el sábado de la siguiente semana(6),
  //Haremos un pequeño ajuste para que empiece en el lunes como 1 y acabe en el domingo como 7
  if(weekDay == 0){
    weekDay = 7;
  }

  return weekDay;
}

function obtenerLunes(fecha){
  var curr = new Date(fecha);
  var first = (curr.getDate() - (curr.getDay() - 1)) ;
  var monday = Utilities.formatDate(new Date(curr.setDate(first)), "UTC", "dd/MM/YYYY");

  return monday;
}

function checkIfFolderExistElseCreate(parent, folderName) {
  var folder;
  try {
      folder = parent.getFoldersByName(folderName).next();
  } catch (e) {
      folder = parent.createFolder(folderName);
  }

  return folder;
}