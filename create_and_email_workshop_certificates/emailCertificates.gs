// ToDo:
// 1. data cleaning (validation, check excel inputted correct data)

function emailCertificates() {
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();                   // excel data
  const taller_certificate_slide_url = sheet.getRange(2,10).getValue().toString();        // certificates
  const pdfFolder_url = sheet.getRange(15, 10).getValue().toString();                     // pdf_folder
  const pdfFolder_id = getIdFromUrl(pdfFolder_url);
  
  Logger.log("Running on sheet: " + sheet.getName().toString());
  Logger.log("id of pdf folder: " + pdfFolder_id);
  
  // slide and PDF folder
  var certificateSlides;
  var pdfFolder;

  try {
    certificateSlides = SlidesApp.openByUrl(taller_certificate_slide_url);
  } catch (err) {
    Logger.log('ERROR: Failed to open certificate slides file: "' + taller_certificate_slide_url + '"');
    Logger.log("Please check that the url is correct and placed in the excel sheet.")
    SpreadsheetApp.getUi().alert('ERROR: Failed to open certificate slides file: "' + taller_certificate_slide_url + '"\nPlease check that the url is correct and placed in the excel sheet.');
    // DEBUG: Logger.log(err.message);
    // DEBUG: Logger.log(err.stack);
    return null;
  }

  try {
    pdfFolder = DriveApp.getFolderById(pdfFolder_id);
  } catch (err) {
    Logger.log('ERROR: Failed to open the PDF folder: "' + pdfFolder_url + '"\t with id: ' + pdfFolder_id);
    Logger.log("Please check that the url is correct and placed in the excel sheet.")
    SpreadsheetApp.getUi().alert('ERROR: Failed to open the PDF folder: "' + pdfFolder_url + '"\t with id: "' + pdfFolder_id + '"\nPlease check that the url is correct and placed in the excel sheet.');
    // DEBUG: Logger.log(err.message);
    // DEBUG: Logger.log(err.stack);
    return null;
  }

  // get dataSheet and values
  // values = [name, email, checkbox_bool]
  const lastColumn = 3;                 // first three columns of data
  const lastRow = sheet.getLastRow();
  const row_length = lastRow - 1;       // remove one to exclude header row 
  const values = sheet.getRange(2,1,row_length, lastColumn).getValues();  // start counting after header row

  // get id_email map
  // {name : [name, email]}
  id_email_map = new Map();
  for(i = 0; i < values.length; i++){
    id_email_map[values[i][0]] = [values[i][0], values[i][1]];
  };

  // get certificates_list (list of every certificate in pdfFolder)
  var certificates_list = new Set();
  var files = pdfFolder.getFiles();
  while(files.hasNext()){
    var file = files.next();
    var name = file.getName().toString().slice(12).replace(/_/g, ' '); // file name: Certificado_NAME.pdf
    certificates_list.add(name);
  };
  // add template to list
  certificates_list.add("{{name}}");

  // for each slide
  const slides = certificateSlides.getSlides();
  for(i = 0; i < slides.length; i++){
    // if not in PDFs list: create PDF and email
    var slide = slides[i];
    var attendee_name = slide.getShapes()[0].getText().asString().trim();
    if(!certificates_list.has(attendee_name)){
      if(!(id_email_map[attendee_name] && id_email_map[attendee_name].length > 0)) {
        Logger.log("ERROR: PDF exists for " + attendee_name + " but removed from data sheet. Please delete the pdf and/or slide.");
        SpreadsheetApp.getUi().alert("ERROR: PDF exists for " + attendee_name + " but removed from data sheet. Please delete the pdf and/or slide.");
      } else {
        // create PDF
        var pdf_file_id = convertSlideToPDFCertificate(slide, pdfFolder);

        // send email if not null
        if(pdf_file_id) {
          _emailPDF(pdf_file_id, id_email_map[attendee_name])
        }
      }
    }
  }
}


function convertSlideToPDFCertificate(slide, folder) {
  // get name from slide
  var name = slide.getShapes()[0].getText().asString();
  name = "Certificado_" + name.toString().trim().replace(/\s/g,"_");

  try {
    // make temp copySlides file
    var copySlides_id = SlidesApp.create(name + "temp_pdf_printer").getId();
    var copySlides = SlidesApp.openById(copySlides_id)

    // copy slide into copySlides
    copySlides.appendSlide(slide);

    // remove first empty slide
    copySlides.getSlides()[0].remove();

    // save edits
    copySlides.saveAndClose();

    // convert copySlides to PDF
    var blob = DriveApp.getFileById(copySlides_id).getAs('application/pdf');
    var new_file = DriveApp.createFile(blob);
    new_file.moveTo(folder);
    new_file.setName(name);

    // delete copySlides file
    DriveApp.getFileById(copySlides_id).setTrashed(true);
  } catch (err) {
    Logger.log("ERROR: Failed to create PDF for file: " + name);
    Logger.log("Please try again.")
    SpreadsheetApp.getUi().alert("ERROR: Failed to create PDF for file: " + name +". Please try again.");
    // DEBUG: Logger.log(err.message);
    // DEBUG: Logger.log(err.stack);
    if(new_file) {
      new_file.setTrashed(true);
    };
    return null;

  } finally {
    // delete copySlides file
    if(copySlides_id) {
      DriveApp.getFileById(copySlides_id).setTrashed(true);
    }
  }
  return new_file.getId();
}


function _emailPDF(fileId, id_email_map_value) {

  // get data from id_email_map_value
  // { name: [name, email]}
  const donor_name = id_email_map_value[0].toString().trim();
  const donor_email = id_email_map_value[1].toString().trim();

  // convert file to PDF (just in case)
  var file = DriveApp.getFileById(fileId);
  var pdf = file.getAs('application/pdf').getBytes();
  var new_file_name = ("Certificado de Taller " + donor_name + ".pdf").replace(/\s/g, "_");
  var attach = {fileName: new_file_name, content:pdf, mimeType:'application/pdf'};

  // send email
  var subject = "Has recibido el Certificado de Taller de ADOPTAMIU PERU";
  var message = 
  "Estimado/a " + donor_name.split(" ")[0] + ":\n" +
  "Por medio de la presente, adjuntamos el Certificado de Taller emitido por Adoptamiu Perú, de acuerdo con el siguiente detalle:\n\n" + 
  "Tipo de documento: Certificado de Taller\n" + 
  "Nombre o Razón Social del cliente: " + donor_name + "\n" +
  "Se adjunta el referido Certificado de Taller en formato PDF. En caso de requerir cualquier coordinación adicional, sírvase comunicarse por whatsapp al 960 736 626 o responder a este propio correo.\n\n" +
  "Atentamente,\n\n" +
  "Adoptamiu Perú";

  MailApp.sendEmail(donor_email, subject, message, { // TODO: change email to donor_email 
    attachments: [attach]
  });

  Logger.log("PDF sent to email: " + donor_email);
}

function getIdFromUrl(url) { return url.match(/[-\w]{25,}$/); }