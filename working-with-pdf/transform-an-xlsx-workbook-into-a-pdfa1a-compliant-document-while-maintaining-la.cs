using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for PdfSaveOptions if needed

class ConvertXlsxToPdfA1a
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Desired output PDF/A‑1a file path
        string destPath = "output.pdf";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Create PDF save options and set PDF/A‑1a compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.Compliance = PdfCompliance.PdfA1a;          // PDF/A‑1a compliance
        pdfOptions.ExportDocumentStructure = true;            // Preserve document structure (optional)

        // Save the workbook as a PDF/A‑1a document while preserving layout
        workbook.Save(destPath, pdfOptions);

        Console.WriteLine("Conversion to PDF/A‑1a completed successfully.");
    }
}