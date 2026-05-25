using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class RemoveWorksheetAndSavePdf
{
    static void Main()
    {
        // Load the workbook from an existing file
        Workbook workbook = new Workbook("input.xlsx");

        // Remove the worksheet named "SheetToRemove"
        workbook.Worksheets.RemoveAt("SheetToRemove");

        // Configure PDF save options (optional)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Export only visible sheets (default behavior)
            SheetSet = SheetSet.Visible
        };

        // Save the modified workbook as PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}