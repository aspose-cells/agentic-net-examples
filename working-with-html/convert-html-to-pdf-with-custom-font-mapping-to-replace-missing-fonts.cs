// Title: C# – Convert HTML to PDF with Custom Font Mapping & Substitution using Aspose.Cells
// Description: Load an HTML file into an Aspose.Cells workbook, set a recursive custom TrueType font folder, define font substitutes, configure PdfSaveOptions (DefaultFont, CheckWorkbookDefaultFont, CheckFontCompatibility), and save the result as a PDF while automatically replacing missing fonts.
// Keywords: Aspose.Cells HTML to PDF | custom font folder Aspose.Cells | font substitution .NET | PdfSaveOptions DefaultFont | CheckFontCompatibility | C# HTML to PDF conversion | font mapping Aspose.Cells | replace missing fonts PDF
// Common Searches: Aspose.Cells map custom fonts when converting HTML to PDF | C# replace missing Arial with Liberation Sans in PDF using Aspose.Cells | set recursive font folder Aspose.Cells PDF export | enable font compatibility checking PdfSaveOptions Aspose.Cells | how to use FontConfigs.SetFontSubstitutes in C#
// Developer Intent: Convert an HTML document to PDF while ensuring any unavailable fonts are automatically replaced using a private font directory and defined substitute fonts.
// Use Cases: Generate branded PDF reports from HTML templates on servers that lack standard system fonts. | Create printable invoices from HTML where the default Arial font may be missing, substituting it with open‑source alternatives. | Run a batch job that processes many HTML files into PDFs, applying a shared custom font repository to maintain consistent appearance.
// AI Prompts: Write C# code that converts HTML to PDF with Aspose.Cells, adds a recursive custom font folder, and sets font substitutes for missing families. | Explain the interaction between FontConfigs.SetFontSubstitutes and PdfSaveOptions.CheckFontCompatibility in Aspose.Cells. | Show how to log each font substitution that occurs during an HTML‑to‑PDF conversion with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an HTML file into an Aspose.Cells workbook, set a recursive custom TrueType font folder, define font substitutes, configure PdfSaveOptions (DefaultFont, CheckWorkbookDefaultFont, CheckFontCompatibility), and save the result as a PDF while automatically replacing missing fonts.
class HtmlToPdfWithFontMapping
{
    static void Main()
    {
        try
        {
            // Set the folder that contains custom TrueType fonts (recursive scan)
            string customFontFolder = @"C:\CustomFonts";
            if (Directory.Exists(customFontFolder))
            {
                FontConfigs.SetFontFolder(customFontFolder, true);
            }

            // Define substitute fonts for a font that might be missing on the target system
            // If "Arial" is not available, Aspose.Cells will try "Liberation Sans" then "DejaVu Sans"
            FontConfigs.SetFontSubstitutes("Arial", new[] { "Liberation Sans", "DejaVu Sans" });

            // Load the source HTML file into a workbook
            string htmlFilePath = @"C:\Input\sample.html";
            if (!File.Exists(htmlFilePath))
                throw new FileNotFoundException("HTML input file not found.", htmlFilePath);

            // Use HtmlLoadOptions to correctly interpret the HTML content
            var htmlLoadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlFilePath, htmlLoadOptions);

            // Configure PDF save options with custom font handling
            var pdfOptions = new PdfSaveOptions
            {
                // Primary font to use; if unavailable, substitutes defined above will be applied
                DefaultFont = "Arial",

                // Try to use the workbook's default font before falling back to system fonts
                CheckWorkbookDefaultFont = true,

                // Ensure font compatibility checking so missing characters are replaced with substitutes
                CheckFontCompatibility = true
            };

            // Ensure the output directory exists
            string pdfOutputPath = @"C:\Output\result.pdf";
            string outputDir = Path.GetDirectoryName(pdfOutputPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook as a PDF file using the configured options
            workbook.Save(pdfOutputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
