using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class SaveSpecificSheetsToPdf
{
    static void Main()
    {
        // Create a new workbook with three worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Name = "First";
        workbook.Worksheets.Add("Second");
        workbook.Worksheets.Add("Third");

        // Populate each sheet with sample data
        workbook.Worksheets["First"].Cells["A1"].PutValue("Data in First sheet");
        workbook.Worksheets["Second"].Cells["A1"].PutValue("Data in Second sheet");
        workbook.Worksheets["Third"].Cells["A1"].PutValue("Data in Third sheet");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save only the second sheet (index 1)
        pdfOptions.SheetSet = new SheetSet(new int[] { 1 });
        workbook.Save("SecondSheetOnly.pdf", pdfOptions);

        // Save only the first and third sheets by name
        pdfOptions.SheetSet = new SheetSet("First", "Third");
        workbook.Save("FirstAndThirdSheets.pdf", pdfOptions);

        // Save all visible sheets (default behavior)
        pdfOptions.SheetSet = SheetSet.Visible;
        workbook.Save("AllVisibleSheets.pdf", pdfOptions);
    }
}