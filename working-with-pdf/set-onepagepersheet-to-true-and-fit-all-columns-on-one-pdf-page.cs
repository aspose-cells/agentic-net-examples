// Title: Export an Excel workbook to PDF with each worksheet on a single page and attempt to fit all columns onto the page using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, enables PdfSaveOptions.OnePagePerSheet, and saves the workbook as a PDF. | Show how to programmatically resize column widths in a worksheet before PDF export when the FitColumnsToPage property is not available. | Create a C# snippet that verifies the source Excel file exists and gracefully handles exceptions during the PDF conversion with Aspose.Cells.
// Common Searches: how to export Excel to PDF with one page per sheet using Aspose.Cells .NET | Aspose.Cells PDF export fit all columns on a single page | C# adjust column widths before saving workbook as PDF with Aspose.Cells | PdfSaveOptions OnePagePerSheet true example Aspose.Cells | handling missing Excel file error when converting to PDF with Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# fit columns to PDF page Aspose.Cells | Excel to PDF conversion column width adjustment | error handling missing workbook Aspose.Cells | Aspose.Cells PDF export per worksheet page

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing Excel workbook, configures PdfSaveOptions with OnePagePerSheet set to true, notes that FitColumnsToPage is unavailable, and saves the workbook as a PDF while performing a file‑existence check and basic exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Render each worksheet on a separate PDF page
                OnePagePerSheet = true

                // Note: FitColumnsToPage is not available in the current Aspose.Cells version.
                // If needed, adjust column widths manually before saving.
            };

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
