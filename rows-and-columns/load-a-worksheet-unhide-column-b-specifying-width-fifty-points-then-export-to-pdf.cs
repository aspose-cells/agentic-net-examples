using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Unhide column B (zero‑based index 1) and set its width to 50 points
        workbook.Worksheets[0].Cells.UnhideColumn(1, 50);

        // Prepare PDF save options (default options are sufficient for this task)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export the worksheet to a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}