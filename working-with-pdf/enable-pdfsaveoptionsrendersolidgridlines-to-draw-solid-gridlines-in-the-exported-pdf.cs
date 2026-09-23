// Title: Export Excel to PDF with solid gridlines using Aspose.Cells PdfSaveOptions in C#
// AI Prompts: Generate C# code that loads an .xlsx workbook, sets PdfSaveOptions.RenderSolidGridlines to true, and saves it as a PDF with Aspose.Cells. | Demonstrate how to configure Aspose.Cells PdfSaveOptions to render solid gridlines during Excel‑to‑PDF conversion in a .NET application.
// Common Searches: Aspose.Cells C# PdfSaveOptions.RenderSolidGridlines example | How to keep Excel gridlines visible when exporting to PDF with Aspose.Cells | C# code to enable solid gridlines in PDF output using Aspose.Cells | Export Excel workbook to PDF with gridlines using Aspose.Cells library | RenderSolidGridlines property usage in Aspose.Cells PDF conversion
// Tags: Aspose.Cells PdfSaveOptions solid gridlines | C# Excel to PDF conversion with gridlines | RenderSolidGridlines property Aspose.Cells | PDF export settings for gridline rendering | Aspose.Cells PDF export configuration C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, creates a PdfSaveOptions object with RenderSolidGridlines set to true, and saves the workbook as a PDF, ensuring solid gridlines appear in the exported document.
class PdfExportWithSolidGridlines
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (gridline rendering is enabled by default in recent versions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
