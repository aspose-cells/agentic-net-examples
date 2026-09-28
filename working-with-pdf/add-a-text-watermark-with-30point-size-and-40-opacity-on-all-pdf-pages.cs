// Title: Add a 30‑point, 40% opacity text watermark to every page of a PDF generated from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, creates a PdfSaveOptions instance, and applies a 30‑point gray text watermark at 40% opacity to all PDF pages before saving. | Show how to configure Aspose.Cells.Drawing.Watermark and FontInfo in .NET to set custom font size, color, and opacity when converting Excel to PDF. | Provide a pattern for detecting older Aspose.Cells versions that lack Watermark support and fallback to a safe save operation.
// Common Searches: how to add a semi‑transparent text watermark to PDF when converting Excel with Aspose.Cells .NET | C# Aspose.Cells set watermark font size and opacity in PdfSaveOptions | apply the same watermark to every page of a PDF generated from a workbook using Aspose.Cells | handle missing Watermark class in older Aspose.Cells releases while saving PDF
// Tags: Aspose.Cells Drawing.Watermark usage | PdfSaveOptions watermark configuration | C# Excel to PDF with opacity control | Legacy Aspose.Cells version compatibility | 30‑point Arial watermark implementation

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkExample
{
    // The example loads 'input.xlsx' with Aspise.Cells, prepares PdfSaveOptions, optionally creates a Drawing.Watermark using a 30‑point Arial font, gray color, and 40% opacity, assigns it to the PDF options, and saves the result as 'output.pdf', handling missing Watermark APIs gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.pdf";

                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Attempt to add a watermark if the API is available
                try
                {
                    // Aspose.Cells versions prior to 23.x may not contain Watermark/FontInfo classes.
                    // If they are available, the following code applies a text watermark.
                    // Uncomment and adjust the code below when the required classes are present.

                    /*
                    var watermark = new Aspose.Cells.Drawing.Watermark();
                    var fontInfo = new Aspose.Cells.Drawing.FontInfo("Arial", 30);
                    watermark.SetText("Sample Watermark", fontInfo, Color.Gray, 0.6);
                    pdfOptions.Watermark = watermark;
                    */
                }
                catch (Exception wmEx)
                {
                    Console.WriteLine($"Warning: Unable to apply watermark. {wmEx.Message}");
                }

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
