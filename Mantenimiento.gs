function imagenesDesdeExcel(){
  var spreadsheetID = ID_HOJA_FORM;
  var ss = SpreadsheetApp.openById(spreadsheetID);
  var rowVals = ss.getRange(75 +":"+ 83).getValues();

  rowVals.forEach((entry)=>{
    var hora = Utilities.formatDate(entry[8], ss.getSpreadsheetTimeZone(), "HH:mm");
    generarImagenHistoria(entry[3], entry[2], hora, formatearFecha(entry[6]), entry[9], entry[6], true);
  });
}
