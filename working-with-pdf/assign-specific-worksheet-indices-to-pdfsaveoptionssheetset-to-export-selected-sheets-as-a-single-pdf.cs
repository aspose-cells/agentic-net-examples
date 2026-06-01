using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ExportSelectedSheetsToPdf
{
    static void Main()
    {
        // Create a new workbook with three worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Name = "First";
        workbook.Worksheets.Add("Second");
        workbook.Worksheets.Add("Third");

        // Populate each sheet with sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Data in First sheet");
        workbook.Worksheets[1].Cells["A1"].PutValue("Data in Second sheet");
        workbook.Worksheets[2].Cells["A1"].PutValue("Data in Third sheet");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export only the first and third sheets (0‑based indexes)
        pdfOptions.SheetSet = new SheetSet(new int[] { 0, 2 });

        // Save the selected sheets as a single PDF file
        workbook.Save("SelectedSheets.pdf", pdfOptions);
    }
}