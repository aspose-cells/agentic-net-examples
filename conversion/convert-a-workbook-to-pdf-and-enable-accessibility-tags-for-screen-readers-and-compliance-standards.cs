// Title: Convert an Excel workbook to PDF/A‑2u with accessibility tags using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file and saves it as a PDF/A‑2u document with accessibility tags using Aspose.Cells. | Show how to configure PdfSaveOptions to set PDF/A‑2u compliance and include document structure for screen readers during Excel‑to‑PDF conversion.
// Common Searches: how to create an accessible PDF/A-2u from Excel using Aspose.Cells C# | Aspose.Cells set PDF compliance to PDF/A-2u with document structure | C# convert .xlsx to PDF with accessibility tags Aspose | enable PDF/A-2u compliance in Aspose.Cells PDF conversion
// Tags: Aspose.Cells PDF/A-2u conversion | C# generate accessible PDF from Excel | PdfSaveOptions compliance option | Enable document structure in PDF export | ExportDocumentStructure property Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an input.xlsx workbook, configures PdfSaveOptions to export document structure for screen readers and to use PDF/A‑2u compliance, then saves the workbook as an accessible PDF file.
class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourcePath = "input.xlsx";

        // Desired output PDF file path
        string pdfPath = "output.pdf";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(sourcePath);

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Enable accessibility tags (document structure) for screen readers
        pdfOptions.ExportDocumentStructure = true;

        // Set a PDF/A compliance level that includes accessibility (PDF/A-2u)
        pdfOptions.Compliance = PdfCompliance.PdfA2u;

        // Save the workbook as a PDF with the configured options
        workbook.Save(pdfPath, pdfOptions);
    }
}
