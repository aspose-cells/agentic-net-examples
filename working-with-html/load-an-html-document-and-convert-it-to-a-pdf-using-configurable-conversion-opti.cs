using System;
using System.IO;
using Aspose.Cells;                 // Workbook, HtmlLoadOptions, PdfSaveOptions
using Aspose.Cells.Rendering;      // PdfCompliance

class HtmlToPdfConverter
{
    static void Main()
    {
        // Input HTML file and output PDF file paths
        string htmlFile = "input.html";
        string pdfFile = "output.pdf";

        // ---------- Load HTML ----------
        // Create load options and configure desired behavior
        HtmlLoadOptions loadOpts = new HtmlLoadOptions();
        loadOpts.SupportDivTag = true;               // Preserve <div> layout
        loadOpts.AutoFitColsAndRows = true;          // Auto‑fit columns/rows after load
        loadOpts.DeleteRedundantSpaces = true;       // Clean up extra spaces
        loadOpts.LoadFormulas = true;                // Preserve formulas if present

        // Load the HTML document into a Workbook (lifecycle rule)
        Workbook workbook = new Workbook(htmlFile, loadOpts);

        // ---------- Save as PDF ----------
        // Create PDF save options and set conversion preferences
        PdfSaveOptions pdfOpts = new PdfSaveOptions();
        pdfOpts.Compliance = PdfCompliance.PdfA1b;   // PDF/A‑1b compliance
        pdfOpts.OnePagePerSheet = true;              // Each sheet on a single page
        pdfOpts.CalculateFormula = true;             // Re‑calculate formulas before saving
        pdfOpts.EmbedStandardWindowsFonts = true;    // Embed standard fonts

        // Save the workbook to PDF using the configured options (lifecycle rule)
        workbook.Save(pdfFile, pdfOpts);

        Console.WriteLine($"HTML file '{htmlFile}' has been converted to PDF '{pdfFile}'.");
    }
}