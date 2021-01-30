function emailPDF(fileId, id_email_map_value) {
  
  // get data from id_email_map_value
  const donor_id = id_email_map_value[0];
  const donor_name = id_email_map_value[1];
  const donor_email = id_email_map_value[2];
  const donor_date = Utilities.formatDate(id_email_map_value[3], "COT", "dd-MM-yyyy");

  // convert file to PDF
  var file = DriveApp.getFileById(fileId);
  var new_file_name = donor_name.toString().replace(/\s/g,"_") + ".pdf";
  // Logger.log("DEBUG pdf file name: " + new_file_name);
  var pdf = file.getAs('application/pdf').getBytes();
  var attach = {fileName: new_file_name, content:pdf, mimeType:'application/pdf'};

  // send email
  var subject = "Has recibido el Certificado de Donación Nro." + donor_id +" de ADOPTAMIU PERU";
  var message = 
  "Estimado/a " + donor_name.split(" ")[0] + ":\n" +
  "Por medio de la presente, adjuntamos el Certificado de Donación emitido por Adoptamiu Perú, de acuerdo con el siguiente detalle:\n\n" + 
  "Tipo de documento: Certificado de Donación\n" + 
  "Serie y número: \t" + donor_id +"\n" + 
  "N° RUC del emisor: 20605733116\n" +
  "Nombre o Razón Social del cliente: " + donor_name + "\n" +
  "Fecha de emisión: " + donor_date + "\n\n" +
  "Se adjunta el referido Certificado de Donación en formato PDF. En caso de requerir cualquier coordinación adicional, sírvase comunicarse al 960 736 626.\n\n" +
  "Atentamente,\n\n" +
  "Adoptamiu Perú";

  MailApp.sendEmail(donor_email, subject, message, {
    attachments: [attach]
  });
}
