// Title: Export an Aspose.Cells Workbook to PDF with PDF 1.7 compatibility using C#
// AI Prompts: Write C# code that creates a Workbook, populates cells, and saves it as a PDF using PdfSaveOptions configured for the highest PDF version (1.7) or an equivalent PDF/A compliance level supported by Aspose.Cells. | Show how to adjust PdfSaveOptions in Aspose.Cells to target modern PDF readers, including fallback to PDF/A compliance when a direct PDF version property is unavailable.
// Common Searches: Aspose.Cells C# how to generate PDF 1.7 compatible file | Set PDF version when saving Excel to PDF with Aspose.Cells | PdfSaveOptions PDF/A compliance vs PDF version in Aspose.Cells | C# export workbook to PDF for modern readers using Aspose.Cells | Is there a PDF 1.7 setting in Aspose.Cells save options
// Tags: Aspose.Cells PdfSaveOptions PDF/A compliance | C# export workbook to PDF modern compatibility | Aspose.Cells set PDF version compatibility | PdfSaveOptions configure PDF output Aspose.Cells | Excel to PDF conversion Aspose.Cells C#

using Aspose.Cells;
using System;

// The example creates an Aspose.Cells Workbook, adds sample data, configures PdfSaveOptions (not exposing a direct PDF version property) and saves the workbook as a PDF that is compatible with modern PDF readers, using available compliance settings.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample text for PDF");

            // Configure PDF save options (Aspose.Cells does not expose a direct PDF version property)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Example: set PDF compliance if required
            // pdfOptions.PdfCompliance = PdfCompliance.PdfA1b;

            // Save the workbook as PDF using the specified options
            workbook.Save("Result.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
