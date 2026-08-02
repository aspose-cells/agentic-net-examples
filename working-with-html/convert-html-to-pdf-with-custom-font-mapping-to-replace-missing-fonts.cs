// Title: C# – Convert HTML to PDF with Custom Font Mapping & Substitutes using Aspose.Cells
// Description: Demonstrates how to register a custom font directory, define font substitutes for missing typefaces, load an HTML file into an Aspose.Cells Workbook, configure PdfSaveOptions with a fallback font and compatibility checks, and export the result as a PDF while preserving layout.
// Keywords: Aspose.Cells | HTML to PDF conversion | custom font folder | font substitution | default font fallback | PdfSaveOptions | C# .NET | missing fonts handling | font compatibility check | PDF export
// Common Searches: Aspose.Cells register custom font folder | HTML to PDF with font substitutes Aspose.Cells | C# set default font for PDF export Aspose.Cells | how to map missing fonts when converting HTML to PDF | PdfSaveOptions font compatibility Aspose.Cells
// Developer Intent: Convert an HTML document to PDF while automatically replacing unavailable fonts with user‑defined alternatives.
// Use Cases: Load a web‑page HTML file into a Workbook and generate a PDF that uses fonts from a private .ttf collection. | Define substitute fonts (e.g., replace Arial with Liberation Sans) to ensure consistent rendering on machines without the original typeface. | Enable font compatibility checks to avoid missing characters and guarantee a printable PDF output.
// AI Prompts: Write C# code that registers a custom font directory and sets font substitutes before converting HTML to PDF with Aspose.Cells. | Explain the PdfSaveOptions properties needed for default font fallback and font‑compatibility verification. | Provide troubleshooting steps when custom fonts are not applied in the exported PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlToPdf
{
    // Demonstrates how to register a custom font directory, define font substitutes for missing typefaces, load an HTML file into an Aspose.Cells Workbook, configure PdfSaveOptions with a fallback font and compatibility checks, and export the result as a PDF while preserving layout.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the folder that contains custom fonts (e.g., .ttf files)
                string customFontFolder = @"C:\MyCustomFonts";

                // Register the custom font folder if it exists
                if (Directory.Exists(customFontFolder))
                {
                    // Register the custom font folder (recursive scan)
                    FontConfigs.SetFontFolder(customFontFolder, true);
                }
                else
                {
                    Console.WriteLine($"Custom font folder not found: {customFontFolder}");
                }

                // Define font substitutes for a font that might be missing in the environment
                // For example, if the HTML uses "Arial" but it's not available, substitute with "Liberation Sans"
                FontConfigs.SetFontSubstitutes("Arial", new[] { "Liberation Sans", "Helvetica", "Verdana" });

                // Load the HTML file into a Workbook
                // Aspose.Cells can directly load HTML documents
                string htmlPath = @"C:\Input\sample.html";

                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Input HTML file not found: {htmlPath}");
                    return;
                }

                Workbook workbook = new Workbook(htmlPath);

                // Configure PDF save options with custom font handling
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Use a default font that is guaranteed to exist (fallback)
                    DefaultFont = "Liberation Sans",
                    // Try to use the workbook's default font first
                    CheckWorkbookDefaultFont = true,
                    // Enable font compatibility checking to substitute missing characters
                    CheckFontCompatibility = true,
                    // Optional: set font encoding (Identity is default)
                    FontEncoding = PdfFontEncoding.Identity
                };

                // Ensure output directory exists
                string outputDir = @"C:\Output";
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF
                string pdfPath = Path.Combine(outputDir, "result.pdf");
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine("HTML has been converted to PDF with custom font mapping.");
                Console.WriteLine($"PDF saved to: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
