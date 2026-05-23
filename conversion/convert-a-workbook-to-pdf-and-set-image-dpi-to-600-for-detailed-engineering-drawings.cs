using System;
using Aspose.Cells;

class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Load the source Excel workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set the machine DPI to 600 for high‑resolution rendering
        CellsHelper.DPI = 600;

        // Create PDF save options and configure image resampling to 600 PPI
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Desired PPI = 600, JPEG quality = 100 (maximum quality)
        pdfOptions.SetImageResample(600, 100);

        // Save the workbook as a PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);

        Console.WriteLine("Workbook successfully converted to PDF with 600 DPI.");
    }
}