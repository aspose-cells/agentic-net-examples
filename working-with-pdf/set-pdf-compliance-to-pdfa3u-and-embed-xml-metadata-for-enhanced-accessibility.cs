using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some content
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample PDF/A-3u with embedded XML metadata");

        // Configure PDF save options with PDF/A‑3u compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u
        };

        // Aspose.Cells does not expose a direct property for embedding custom XML (XMP) metadata.
        // If the library version provides a Metadata property, it can be set here, e.g.:
        // pdfOptions.Metadata = yourXmlDocument;
        // Otherwise, consider post‑processing the generated PDF with Aspose.Pdf to add the metadata.

        // Save the workbook as a PDF file with the specified options
        workbook.Save("Result.pdf", pdfOptions);
    }
}

// Author: Example demonstrating PDF/A‑3u compliance with a note on XML metadata embedding.