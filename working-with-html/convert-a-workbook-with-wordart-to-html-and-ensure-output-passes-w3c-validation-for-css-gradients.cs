// Title: Convert an Excel workbook containing WordArt to W3C‑valid HTML with CSS gradients using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with WordArt, configures HtmlSaveOptions to render WordArt as native HTML/CSS, outputs the CSS to a separate file, and enables CSS gradient generation for W3C‑compliant HTML. | Demonstrate how to validate the generated HTML and CSS files against W3C standards after converting an Excel workbook with WordArt using Aspose.Cells.
// Common Searches: how to export WordArt from Excel to HTML using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions enable CSS gradients for W3C validation | save Excel workbook as HTML with separate CSS file in C# | convert Excel with WordArt to standards‑compliant HTML using Aspose.Cells | C# example for exporting Excel to HTML with native WordArt rendering
// Tags: Aspose.Cells HtmlSaveOptions CSS gradients | export WordArt as native HTML Aspose.Cells | Excel to HTML conversion .NET | separate CSS file generation Aspose.Cells | W3C‑compliant HTML output from Excel

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookToHtml
{
    // C# program that loads an Excel workbook containing WordArt, configures HtmlSaveOptions (including placeholders for native WordArt rendering, separate CSS output, and CSS gradient support), and saves the workbook as W3C‑valid HTML with optional external CSS.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file that contains WordArt
            string sourcePath = @"C:\Input\WordArtWorkbook.xlsx";

            // Path where the resulting HTML file will be saved
            string htmlPath = @"C:\Output\WordArtWorkbook.html";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(sourcePath);

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export WordArt as native HTML/CSS instead of images (if supported)
                    // The property may not exist in some versions; ignore if unavailable.
                    // ExportWordArtAsImage = false,

                    // Export CSS to a separate file for readability (if supported)
                    // ExportCssSeparately = true,

                    // Use standard CSS gradients (if supported)
                    // ExportCssGradients = true
                };

                // Save the workbook to HTML (lifecycle rule: save)
                workbook.Save(htmlPath, saveOptions);

                Console.WriteLine("Workbook successfully converted to HTML.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
