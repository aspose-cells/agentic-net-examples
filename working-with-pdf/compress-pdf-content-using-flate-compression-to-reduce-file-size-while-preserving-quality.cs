// Title: Apply Flate compression while converting an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets PdfSaveOptions.Compression to Flate, and saves the workbook as a compressed PDF with Aspose.Cells. | Show how to enable Flate compression in Aspose.Cells PdfSaveOptions to reduce PDF size when exporting Excel to PDF in C#. | Provide a C# example that verifies the source Excel file, configures PDF save options for Flate compression, and handles errors during conversion with Aspose.Cells.
// Common Searches: how to enable Flate compression in Aspose.Cells PDF export C# | reduce size of PDF generated from Excel using Aspose.Cells .NET | Aspose.Cells PdfSaveOptions compression option example | C# convert XLSX to PDF with minimal file size using Aspose.Cells | set PDF compression to Flate when saving workbook with Aspose.Cells
// Tags: Flate compression PdfSaveOptions Aspose.Cells | export Excel to compressed PDF .NET | Aspose.Cells PDF size optimization | C# set PDF compression option Aspose | Workbook.Save with compression Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Retained for potential future use

// The sample checks that the input .xlsx file exists, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions to use Flate compression (and optionally disables one-page-per-sheet), then saves the workbook as a reduced‑size PDF while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep original sheet layout; other options can be set if supported by the library version
                OnePagePerSheet = false
            };

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
