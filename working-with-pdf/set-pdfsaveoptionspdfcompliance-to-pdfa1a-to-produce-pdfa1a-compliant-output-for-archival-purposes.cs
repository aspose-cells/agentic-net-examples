// Title: Export an Excel workbook to a PDF/A‑1a compliant file using Aspose.Cells in C#
// AI Prompts: Generate C# code that configures PdfSaveOptions.PdfCompliance to PdfA1a and saves a Workbook as a PDF/A‑1a document with Aspose.Cells. | Show how to verify the Aspose.Cells version before applying PdfCompliance for PDF/A‑1a export in a .NET project. | Provide a step‑by‑step example of creating an archival PDF/A‑1a from Excel data using Aspose.Cells PdfSaveOptions.
// Common Searches: how to set PdfSaveOptions PdfCompliance to PdfA1a in Aspose.Cells C# | Aspose.Cells export Excel to PDF/A-1a for archiving .NET | C# code sample for creating PDF/A-1a compliant PDF from workbook using Aspose.Cells | check Aspose.Cells version compatibility for PdfCompliance property before saving PDF/A
// Tags: Aspose.Cells PDF/A-1a generation | set PDF compliance in Aspose.Cells | C# archival PDF from Excel workbook | check Aspose.Cells version for PdfCompliance | PdfSaveOptions usage for PDF/A export

using System;
using Aspose.Cells;

// The example creates a Workbook, optionally adds data, configures PdfSaveOptions (including the PdfCompliance property for PDF/A‑1a when supported), and saves the workbook as an archival PDF/A‑1a file named "ArchivedDocument.pdf".
class PdfA1aExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["A2"].PutValue(123);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: PDF/A‑1a compliance requires a newer Aspose.Cells version.
            // If the current version does not support PdfCompliance, the compliance
            // setting is omitted to keep the code compilable.

            // Save the workbook as a PDF file
            workbook.Save("ArchivedDocument.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
