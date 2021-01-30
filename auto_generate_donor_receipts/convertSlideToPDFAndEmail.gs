function convertSlideToPDFAndEmail(id_email_map, receipt_slide_id, drive_pdf_folder) {
  //  get receiptSlides
  const slideFile  = DriveApp.getFileById(receipt_slide_id);
  const receiptSlides= SlidesApp.openById(slideFile.getId());

  // get pdfFolder
  const pdfFolder = DriveApp.getFolderById(drive_pdf_folder);

  // get constancias_list (constancia of every file in pdfFolder)
  var constancias_list = new Set();
  var files = pdfFolder.getFiles();
  while(files.hasNext()){
    var file = files.next();
    var name = file.getName().toString();                     // file name: Constancia_#####_NAME.pdf
    var re = "_([0-9]+)_";
    var id = name.match(re) ? name.match(re)[1] ? name.match(re)[1] : "" : "";
    constancias_list.add(Number(id));
  };

  // for each slide
  const slides = receiptSlides.getSlides();
  for(i = 0; i < slides.length; i++){
    var slide = slides[i];
    // get slide id
    var donor_id = slide.getShapes()[0].getText().asString();
    if(!constancias_list.has(Number(donor_id)) && !isNaN(Number(donor_id)))
    {
      // get name
      var name = slide.getShapes()[2].getText().asString();
      name = "Constancia_" + donor_id.toString().replace("\n", "") + "_" + name.toString().replace(/\s/g, "_").replace(/_$/g, "");

      // copy slide into new file copySlides
      try {
        var copySlides_id = SlidesApp.create(name + "temp_pdf_printer").getId();
        SlidesApp.openById(copySlides_id).appendSlide(slide);
      
        // remove first empty slide
        SlidesApp.openById(copySlides_id).getSlides()[0].remove();
        
        // save edits
        SlidesApp.openById(copySlides_id).saveAndClose();

        // convert copySlides to PDF
        var blob = DriveApp.getFileById(copySlides_id).getAs('application/pdf');
        var new_file = DriveApp.createFile(blob);
        new_file.moveTo(pdfFolder);
        new_file.setName(name);

        // automatically email pdf to donors
        emailPDF(copySlides_id, id_email_map[Number(donor_id)]);

        // delete copySlides file
        DriveApp.getFileById(copySlides_id).setTrashed(true);

      } catch (err) {
        Logger.log("Failed to create PDF for Constancia#" + donor_id);
        Logger.log(err.message);
        Logger.log(err.stack);
        if(new_file) {
          new_file.setTrashed(true);
        };
      } finally {
        // delete copySlides file
        if(copySlides_id) {
          DriveApp.getFileById(copySlides_id).setTrashed(true);
        }
        continue;
      }
    };
  };
}

