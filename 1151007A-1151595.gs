//CRISTIAN FERNANDO CAMARGO CASTELLANOS  1151595 //

function enviarDoc() {
  
  var sps = SpreadsheetApp.getActive(),
      sheet = sps.getSheetByName('3raNota'),
      data  = sheet.getDataRange().getValues(),
      rowI = 1,
      rowF = sheet.getLastRow()-1;
  
  
  const docId = '1Sv-UdEI5NsUiPvvUtgd6lgx424TKo4TTk5MZKyeRs9E';
  var doc = DocumentApp.openById(docId);
  
  doc.getBody().clear();
    
  doc.getBody().appendParagraph('TERCERA NOTA');
  doc.getBody().appendHorizontalRule();
  
  var cells = [
  ['Código', 'Nombres' , 'Apellidos','Correo']
  ];
  
  for(var i =rowI; i<=rowF; i++){
    const row = data[i];
    
    const codigo = row[1];
    const nombres = row[2];
    const apellidos = row[3];
    const nota = row[4];
    const observacion = row[5];
    const correo = row[6];
    
    var estudiante = [codigo,nombres,apellidos,correo]
    cells.push(estudiante);
    
  }
  
  doc.getBody().appendTable(cells);
}


function enviarCorreo(){
  
   var sps = SpreadsheetApp.getActive(),
      sheet = sps.getSheetByName('3raNota'),
      data  = sheet.getDataRange().getValues(),
      rowI = 1,
      rowF = sheet.getLastRow()-1;
  
   for(var i =rowI; i<=rowF; i++){
    const row = data[i];
    
    const codigo = row[1];
    const nombres = row[2];
    const apellidos = row[3];
    const nota = row[4];
    const observacion = row[5];
    const correo = row[6];
    
      MailApp.sendEmail(
    {
      to:correo,
      subject:'Google App Script',
      body: ('Estimad@ estudiante '+nombres+' '+apellidos+' Código: '+codigo+'\n'+
             'Nos permitimos informarle que su calificación definitiva en el curso es: '+nota+'\n'+
             'por la siguiente observacion '+'"'+observacion+'"'+'\n'+
             'Gracias por su atencion y buen dia.')
    }
  )
    
  }
  
}



function crearPDF(){
  
  enviarDoc();
  
  const docId = '1Sv-UdEI5NsUiPvvUtgd6lgx424TKo4TTk5MZKyeRs9E';
  const carpetaId = '1ZpKPaK5ecdd_t_sxih_hQA7_rlVFjnc5';
  
  var doc = DriveApp.getFileById(docId);
  var carpetaDrive = DriveApp.getFolderById(carpetaId);
  
  const docCopia = doc.makeCopy(carpetaDrive);
  var docPDF = DocumentApp.openById(docCopia.getId());
  
  const pdf = docCopia.getAs(MimeType.PDF);
  
  carpetaDrive.createFile(pdf).setName('1151007A-1151595');
  
  

}