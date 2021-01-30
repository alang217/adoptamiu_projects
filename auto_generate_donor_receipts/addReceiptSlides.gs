function addReceiptSlides() {
  const apadrinamiento_receipt_slide_id =  "#####RECEIPT_SLIDES_FILE_ID#####";   // constancias
  const adopcion_receipt_slide_id =  "#####RECEIPT_SLIDES_FILE_ID#####";   // constancias
  const donacion_receipt_slide_id =  "#####RECEIPT_SLIDES_FILE_ID#####";   // constancias
  const data_sheet_id = "#####RECEIPT_SHEETS_FILE_ID#####";       // excel data

  // get dataSheet and values
  // values = [id, amount, name, dni, amount_spanish, date, email, apadrinamiento_bool, adopcion_bool, donacion_bool]
  const sheetFile = DriveApp.getFileById(data_sheet_id);
  const dataSheetFile = SpreadsheetApp.openById(sheetFile.getId());
  const sheet = dataSheetFile.getSheets()[0];
  const lastColumn = dataSheetFile.getLastColumn();
  const lastRow = dataSheetFile.getLastRow();
  const row_length = lastRow - 2;
  const column_length = lastColumn;
  const values = sheet.getRange(3,1,row_length, column_length).getValues();

  //  get id_used from each receiptSlides:
  var id_used = []
  //    Apadrinamiento
  const apadrinamientoReceiptSlides= SlidesApp.openById(apadrinamiento_receipt_slide_id);
  apadrinamientoReceiptSlides.getSlides().forEach(function(slide) {
    var num = Number(slide.getPageElements()[0].asShape().getText().asString());
    if (!isNaN(Number(num))){
      id_used.push(Number(num));};
  });

  //    Adopcion
  const adopcionReceiptSlides= SlidesApp.openById(adopcion_receipt_slide_id);
  adopcionReceiptSlides.getSlides().forEach(function(slide) {
    var num = Number(slide.getPageElements()[0].asShape().getText().asString());
    if (!isNaN(Number(num))){
      id_used.push(Number(num));};
  });

  //    Donacion
  const donacionReceiptSlides= SlidesApp.openById(donacion_receipt_slide_id);
  donacionReceiptSlides.getSlides().forEach(function(slide) {
    var num = Number(slide.getPageElements()[0].asShape().getText().asString());
    if (!isNaN(Number(num))){
      id_used.push(Number(num));};
  });

  // Logger.log("DEBUG [id_used]: " + id_used);

  // get id_missing:  id_given - id_used (list of ids that don't have receipts)
  var id_given = sheet.getRange(3,1,row_length).getValues().map(function(x) {return Number(x)});
  var id_missing = id_given.filter( function( value ) {
    return !id_used.includes( Number(value) );
  } );
  
  // for each id_missing, copy the template slide and fill in data
  for(i = 0; i < row_length; i++) {
    donor = values[i];
    if (id_missing.includes(Number(donor[0]))){
      
      // if missing data, skip
      if (donor.slice(0,-3).includes("") || donor.slice(0,-3).includes(undefined)) {
        Logger.log("Constancia #" + donor[0].toString() +" falta data.");
      } else {

        // get data
        var id = donor[0];
        var amount = "S/." + Number(donor[1]).toFixed(2);
        var name = donor[2];
        var dni = donor[3];
        var amount_spanish = donor[4];
        var day = Number(Utilities.formatDate(donor[5], "PET", "dd"));
        var month = Utilities.formatDate(donor[5], "PET", "MM");
        switch(month.toString()){
          case "01": month = "Enero"; break;
          case "02": month = "Febrero"; break;
          case "03": month = "Marzo"; break;
          case "04": month = "Abril"; break;
          case "05": month = "Mayo"; break;
          case "06": month = "Junio"; break;
          case "07": month = "Julio"; break;
          case "08": month = "Agosto"; break;
          case "09": month = "Septiembre"; break;
          case "10": month = "Octubre"; break;
          case "11": month = "Noviembre"; break;
          case "12": month = "Diciembre"; break;
          default:  month = "";
        };
        var year = Utilities.formatDate(donor[5], "PET", "yyyy");
        var apadrinamiento = donor[7];
        var donacion = donor[8];
        var adopcion = donor[9];

        // for donor 7,8,9, choose appropriate slide
        var receiptSlides;
        if(apadrinamiento !== "") {
          receiptSlides = SlidesApp.openById(apadrinamiento_receipt_slide_id);
        } else if(adopcion !== "") {
          receiptSlides = SlidesApp.openById(adopcion_receipt_slide_id);
        } else if(donacion !== "") {
          receiptSlides = SlidesApp.openById(donacion_receipt_slide_id);
        } else {
          Logger.log("Error: Constancia #" + donor[0].toString() +" falta data.");
          continue;
        }
        // Logger.log(receiptSlides.getName());
        // get templateSlide
        const templateSlide = receiptSlides.getSlides()[0];

        // make copySlide of templateSlide 
        templateSlide.duplicate();
        var slides = receiptSlides.getSlides();
        var copySlide = slides[1];

        // input data into copySlide
        var shapes = copySlide.getShapes();
        shapes.forEach(function(shape){
          shape.getText().replaceAllText("{{id}}", id.toString());
          shape.getText().replaceAllText("{{amount}}", amount.toString());
          shape.getText().replaceAllText("{{name}}", name.toString());
          shape.getText().replaceAllText("{{dni}}", dni.toString());
          shape.getText().replaceAllText("{{amount_spanish}}", amount_spanish.toString());
          shape.getText().replaceAllText("{{day}}", day.toString());
          shape.getText().replaceAllText("{{month}}", month.toString());
          shape.getText().replaceAllText("{{year}}", year.toString());
        });
      };
    };
  };
}

