// Title: Convert an HTML file to a PDF/A‑1b compliant PDF using Aspose.Cells in C#
// Common Searches: how to save HTML as PDF/A-1b with Aspose.Cells in C# | C# Aspose.Cells PDF/A-1b compliance when converting HTML | generate archival PDF/A-1b from HTML using Aspose.Cells library | Aspose.Cells load HTML and export to PDF/A-1b in .NET | set PdfSaveOptions.Compliance to PdfA1b for HTML to PDF conversion
// Tags: Aspose.Cells HTML to PDF/A-1b export | PdfSaveOptions PdfA1b compliance configuration | LoadOptions Html format workbook initialization | C# archival PDF creation with Aspose.Cells | PDF/A-1b generation from HTML in .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an HTML file into an Aspose.Cells Workbook using LoadOptions with Html format, configures PdfSaveOptions.Compliance to PdfCompliance.PdfA1b, and saves the workbook as a PDF/A‑1b compliant PDF, ensuring the input file exists and the output directory is created.
class HtmlToPdfA1bConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = @"C:\Input\document.html";

        // Path for the output PDF/A‑1b file
        string pdfPath = @"C:\Output\document.pdf";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException("Input HTML file not found.", htmlPath);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(pdfPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the HTML file into a workbook using LoadOptions
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Configure PDF save options for PDF/A‑1b compliance
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b
            };

            // Save the workbook as a PDF/A‑1b document
            workbook.Save(pdfPath, saveOptions);

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
