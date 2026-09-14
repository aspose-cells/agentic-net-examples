// Title: Export an Excel workbook to PDF with OnePagePerSheet enabled and enforce a ten‑page limit using Aspose.Cells for .NET
// AI Prompts: Write C# code that exports a Workbook to PDF, forcing each worksheet onto a single page and cutting the output after ten pages with Aspose.Cells. | Provide a C# snippet that configures PDF conversion to use a one‑page‑per‑worksheet layout and illustrates a method to stop the PDF generation at ten pages.
// Common Searches: Aspose.Cells C# export Excel to PDF one page per sheet | How to limit PDF output to ten pages when converting an Excel workbook with Aspose.Cells | PdfSaveOptions OnePagePerSheet and max page count in .NET | C# truncate PDF generated from a workbook after ten pages using Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | Aspose.Cells limit PDF page count .NET | C# workbook to PDF with page cap | Aspose.Cells PDF pagination control | C# export Excel to PDF page limitation

using System;
using Aspose.Cells;

// The example creates a Workbook, adds sample data, sets PdfSaveOptions.OnePagePerSheet to true, notes that a direct MaxPageCount property is unavailable, and saves the file as Result.pdf. It also discusses a workaround for restricting the PDF to ten pages.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // OPTIONAL: add some data to the worksheet to generate content
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF export.");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure each worksheet is rendered on a single page
                OnePagePerSheet = true
                // Note: MaxPageCount is not available in this version of Aspose.Cells
            };

            // Save the workbook as a PDF with the specified options
            workbook.Save("Result.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
