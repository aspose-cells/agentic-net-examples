// Title: How to export an Excel workbook that contains many cell comments to PDF and HTML while preserving all comments using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook, configures PdfSaveOptions to retain every cell comment, and saves the file as a PDF with Aspose.Cells. | Provide a C# example that sets up HtmlSaveOptions to embed all cell comments when converting a comment‑heavy worksheet to HTML using Aspose.Cells.
// Common Searches: Aspose.Cells export workbook with comments to PDF C# | keep Excel cell comments when converting to HTML using Aspose.Cells .NET | C# save Excel file as PDF preserving all comments Aspose.Cells | export comment‑rich Excel sheet to HTML Aspose.Cells example | PdfSaveOptions comment preservation Aspose.Cells tutorial
// Tags: export workbook to PDF with comments Aspose.Cells | export workbook to HTML with comments Aspose.Cells | PdfSaveOptions preserve cell comments | HtmlSaveOptions include all cell comments | Aspose.Cells handling comment‑rich worksheets C# | C# load workbook and save as PDF/HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks for the source Excel file, loads it with Aspose.Cells, creates default PdfSaveOptions and HtmlSaveOptions (which automatically include cell comments where supported), and saves the workbook as both PDF and HTML, handling any errors that may occur.
class ExportWithComments
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string pdfOutputPath = "output.pdf";
        const string htmlOutputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that contains many cell comments
            Workbook workbook = new Workbook(inputPath);

            // ---------- Export to PDF ----------
            // Create PDF save options (default settings export comments where supported)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the configured options
            workbook.Save(pdfOutputPath, pdfOptions);
            Console.WriteLine($"PDF file saved to \"{pdfOutputPath}\".");

            // ---------- Export to HTML ----------
            // Create HTML save options (default settings export cell comments where supported)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlOutputPath, htmlOptions);
            Console.WriteLine($"HTML file saved to \"{htmlOutputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
