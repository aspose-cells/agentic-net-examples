// Title: Export an Excel worksheet with embedded OLE objects to PDF while keeping placeholders using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook containing OLE objects and saves it as a PDF, ensuring the OLE objects appear as image placeholders with Aspose.Cells. | Show how to configure PdfSaveOptions in Aspose.Cells to preserve OLE object placeholders during an Excel‑to‑PDF conversion. | Add robust file‑existence validation and exception handling to a C# Aspose.Cells routine that converts a worksheet with embedded OLE objects to PDF.
// Common Searches: Aspose.Cells .NET export worksheet with embedded OLE objects to PDF preserving placeholders | C# convert Excel file containing OLE objects to PDF without rendering the objects | How to keep OLE object images when saving XLSX as PDF using Aspose.Cells | PdfSaveOptions settings for OLE object handling in Aspose.Cells conversion
// Tags: excel-to-pdf conversion with OLE placeholders | Aspose.Cells PDF export OLE handling | C# preserve embedded OLE objects in PDF | worksheet PDF save preserving object images | Aspose.Cells OLE object placeholder rendering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // This program checks that input.xlsx exists, loads it with Aspose.Cells, and saves it as output.pdf using PdfSaveOptions, which by default retains OLE objects as image placeholders.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains OLE objects.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options. Default behavior preserves OLE objects as image placeholders.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Additional options can be set here, e.g., compliance level, page layout, etc.
                // Compliance = PdfCompliance.PdfA1b,
                // OnePagePerSheet = false
            };

            // Save the workbook (or specific worksheet) to PDF.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
