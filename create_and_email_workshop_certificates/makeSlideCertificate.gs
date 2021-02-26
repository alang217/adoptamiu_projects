// ToDo:
// 1. data cleaning (validation, check excel inputted correct data)
function makeSlideCertificate() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();                   // excel data
  const taller_certificate_slide_url = sheet.getRange(2,10).getValue().toString();        // certificates
  Logger.log("Running on sheet: " + sheet.getName().toString());
  Logger.log("URL of certificates: " + taller_certificate_slide_url);
  
  // get dataSheet and values
  // values = [name, email, checkbox_bool]
  const lastColumn = 3;                 // first three columns of data
  const lastRow = sheet.getLastRow();
  const row_length = lastRow - 1;       // remove one to exclude header row 
  const values = sheet.getRange(2,1,row_length, lastColumn).getValues();  // start counting after header row
  
  // open certificate slides
  var certificateSlides;
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

  //  get name_used from each receiptSlides:
  var name_used = new Set();
  certificateSlides.getSlides().forEach(function(slide) {
    var name = slide.getPageElements()[0].asShape().getText().asString().trim();
    if (name){
      name_used.add(name);};
  });
  name_used.delete("{{name}}");

  // Logger.log("DEBUG [name_used]: " + name_used);

  // get name_missing:  name_given - name_used (list of names that don't have receipts)
  var name_given = sheet.getRange(2,1,row_length).getValues().map(function(x) {return x.toString().trim()});
  var name_missing = new Set(name_given.filter( function( value ) {
    return !name_used.has( value );
  } ));
  
  // for each name_missing, copy the template slide and fill in data
  for(i = 0; i < row_length; i++) {
    attendee = values[i].slice(0,3);
    if (name_missing.has((attendee[0].toString().trim()))){
      // Logger.log(attendee); // DEBUG
      
      // if missing data, skip
      if (attendee.includes("") || attendee.includes(undefined) || attendee[2] == false) {
        if(attendee[0] != "") Logger.log("Persona '" + attendee[0].toString() +"' falta data.");
      } else {

        // get data
        var name = attendee[0];
        var email = attendee[1];
        
        // get templateSlide
        const templateSlide = certificateSlides.getSlides()[0];

        // make copySlide of templateSlide 
        templateSlide.duplicate();
        var slides = certificateSlides.getSlides();
        var copySlide = slides[1];

        // input data into copySlide
        var shapes = copySlide.getShapes();
        shapes.forEach(function(shape){
          shape.getText().replaceAllText("{{name}}", name.toString());
        });
        Logger.log("Created certificate for: " + name);
      };
    };
  };
}
