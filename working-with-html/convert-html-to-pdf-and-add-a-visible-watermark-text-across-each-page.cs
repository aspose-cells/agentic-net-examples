// Title: Convert an HTML file to PDF with a diagonal CONFIDENTIAL text watermark using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells convert HTML to PDF and add a red diagonal watermark | How to apply a text watermark when saving a workbook as PDF with Aspose.Cells | Aspose.Cells PDFSaveOptions watermark example for HTML to PDF conversion | Render HTML to PDF with custom RenderingWatermark using Aspose.Cells .NET | Add CONFIDENTIAL watermark to each page of PDF generated from HTML in C#
// Tags: Aspose.Cells HTML to PDF conversion | RenderingWatermark configuration C# | PdfSaveOptions watermark property | LoadOptions HTML workbook Aspose.Cells | Diagonal text watermark PDF Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace HtmlToPdfWithWatermark
{
    // The sample loads an HTML file into an Aspose.Cells Workbook, creates a red 45-degree "CONFIDENTIAL" RenderingWatermark with custom font and opacity, assigns it to PdfSaveOptions, and saves the workbook as a PDF while preserving the original pagination.
    class Program
    {
        static void Main(string[] args)
        {
            // Input HTML file path (can be passed as first argument)
            string htmlPath = args.Length > 0 ? args[0] : "input.html";
            // Output PDF file path (can be passed as second argument)
            string pdfPath = args.Length > 1 ? args[1] : "output.pdf";
            // Watermark text
            string watermarkText = "CONFIDENTIAL";

            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
                return;
            }

            try
            {
                // Load the HTML file into a workbook using the appropriate constructor
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(htmlPath, loadOptions);
                Console.WriteLine("HTML loaded into workbook successfully.");

                // Create a font for the watermark
                RenderingFont watermarkFont = new RenderingFont("Arial", 72f)
                {
                    Bold = true,
                    Color = Color.Red,
                    Italic = false
                };

                // Create the text watermark
                RenderingWatermark watermark = new RenderingWatermark(watermarkText, watermarkFont)
                {
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center,
                    Rotation = 45,               // degrees
                    Opacity = 0.3f,              // 30% opacity
                    IsBackground = false,        // draw over content
                    ScaleToPagePercent = 100     // fill the page width
                };

                // Configure PDF save options and attach the watermark
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Watermark = watermark,
                    OnePagePerSheet = false      // keep original pagination
                };

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(pdfPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF with the watermark applied
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{pdfPath}' with watermark.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
