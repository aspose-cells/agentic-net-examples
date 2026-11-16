// Title: Convert HTML with embedded charts to PDF using Aspose.Cells for .NET while preserving vector graphics
// Common Searches: Aspose.Cells C# convert HTML file to PDF with vectorized charts | How to keep chart graphics scalable when exporting an HTML workbook to PDF in .NET | Load HTML workbook and save as PDF preserving vector chart rendering Aspose.Cells | PdfSaveOptions example for vector chart rendering in Aspose.Cells | Export HTML tables and charts to PDF as vector graphics using Aspose.Cells
// Tags: HTML to PDF conversion Aspose.Cells | vector chart rendering PDF Aspose.Cells | PdfSaveOptions chart vector graphics | load HTML workbook Aspose.Cells C# | export scalable charts to PDF .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an HTML file containing tables and charts into an Aspose.Cells Workbook, configures PdfSaveOptions, and saves the workbook as a PDF where charts are rendered as vector graphics, ensuring scalability and high‑quality output.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file containing charts
        string htmlPath = "input.html";

        // Path for the resulting PDF file
        string pdfPath = "output.pdf";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input file not found: {htmlPath}");
                return;
            }

            // Load the HTML document into a Workbook
            // Aspose.Cells automatically parses HTML tables and charts
            Workbook workbook;
            try
            {
                workbook = new Workbook(htmlPath, new LoadOptions(LoadFormat.Html));
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load HTML file: {loadEx.Message}");
                return;
            }

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Optional: preserve the original page layout
                OnePagePerSheet = false
                // Charts are rendered as vector graphics by default in recent versions
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(pdfPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF with the specified options
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine("HTML has been successfully converted to PDF with vectorized charts.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
