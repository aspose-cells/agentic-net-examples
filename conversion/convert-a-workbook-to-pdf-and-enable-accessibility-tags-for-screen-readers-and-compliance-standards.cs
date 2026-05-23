using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load an existing Excel workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Enable accessibility tags (document structure) for screen readers
        pdfOptions.ExportDocumentStructure = true;

        // Set PDF/A compliance (e.g., PDF/A-1b) to meet standard requirements
        pdfOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}