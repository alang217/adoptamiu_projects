// ToDo:
// 1. data cleaning (validation, check excel inputted correct data)

function emailReceiptsAsPDF() {
  // spreadsheet data
  const data_sheet_id = "#####RECEIPT_SHEETS_FILE_ID#####";

  // slides and PDF folders
  const file_ids = []
  // add Apadrinamiento
  const apadrinamiento_ids = [
    "#####RECEIPT_SLIDES_FILE_ID#####",   // constancias
    "#####RECEIPT_PDF_FOLDER_ID#####"     // pdf_folder
  ]
  file_ids.push(apadrinamiento_ids);

  // add Adopcion
  const adopcion_ids = [
    "#####RECEIPT_SLIDES_FILE_ID#####",   // constancias
    "#####RECEIPT_PDF_FOLDER_ID#####"     // pdf_folder
  ]
  file_ids.push(adopcion_ids);;

  // add Donacion
const donacion_ids = [
    "#####RECEIPT_SLIDES_FILE_ID#####",   // constancias
    "#####RECEIPT_PDF_FOLDER_ID#####"     // pdf_folder
  ]
  file_ids.push(donacion_ids);
  
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

  // get id_email map
  // {id : [id, name, email, date]}
  id_email_map = new Map();
  for(i = 0; i < values.length; i++){
    id_email_map[Number(values[i][0])] = [values[i][0], values[i][2], values[i][6], values[i][5]];
  };

  file_ids.forEach(function(ids){
    convertSlideToPDFAndEmail(id_email_map, ids[0], ids[1]);
  });

}

