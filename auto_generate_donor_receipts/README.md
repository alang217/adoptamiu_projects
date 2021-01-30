# Auto-Generate Donor Receipts

## Summary
After a user has donated certain amount of money to Adoptamiu, the organization is required to deliver a "Constancia de Donacion" (aka Donation Receipt) to the user. These donations could occur randomly or periodically.
Currently, Adoptamiu uses Google Spreadsheets to keep track of their donation data and Google Slides to create these Donation Receipts.

## Workflow
This project uses Google app scripts to help automate the creation of these Donation Receipts and automatically email a PDF version to the users:
1. An edit on Donation Data Spreadsheet triggers `addReceiptSlides` script.
2. `addReceiptSlides`: read new entry in Donation Data SpreadSheet. If data validated, then create a new Donation Receipt slide in appropriate Donation Slide file.
3. At any point, a worker can verify the newly created slides and then proceed by Running `emailReceiptsAsPDF` script.
4. `emailReceiptsAsPDF`: Identifies Receipt Slides that have not been converted to PDFs. These Slides are then converted to PDF files, moved to their appropriate Receipt PDF Folder, then sent out to email to the users. Uses two heler functions called `convertSlideToPDFAndEmail` and `emailPDF`.
5. `convertSlideToPDFAndEmail`: Given a slide, converts it to a PDF by copying the slide into a different file then saving it in the appropriate folder by manipulating its `Blob` data.
6. `emailPDF`: Given a file ID of the Receipt PDF and the appropriate user data, sends out a template email with the PDF attached.

---

## Future Updates
- Connect with their website database to keep track of their online donors as well.
