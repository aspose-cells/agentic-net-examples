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
        workbook.Worksheets[0].Cells["A1"].PutValue("Data in First sheet");

        Worksheet second = workbook.Worksheets.Add("Second");
        second.Cells["A1"].PutValue("Data in Second sheet");

        Worksheet third = workbook.Worksheets.Add("Third");
        third.Cells["A1"].PutValue("Data in Third sheet");

        // Configure PDF save options to export only specific sheets
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Use SheetSet with exact zero‑based indexes (0 = First, 2 = Third)
        pdfOptions.SheetSet = new SheetSet(new int[] { 0, 2 });

        // Save the workbook to PDF; only the selected sheets will be included
        workbook.Save("SelectedSheets.pdf", pdfOptions);
    }
}