// Title: How to convert a TSV spreadsheet to a PDF/A‑2b compliant PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads a TSV file with Aspose.Cells LoadOptions and saves it as a PDF/A‑2b document while embedding Windows fonts. | Show how to configure PdfSaveOptions.Compliance to PdfA2b and enable font embedding for archival PDF output. | Write robust error‑handling that verifies the TSV source, creates the output folder if missing, and logs any conversion exceptions.
// Common Searches: asp.net core convert tsv file to PDF/A-2b using Aspose.Cells | c# example of PdfSaveOptions for PDF/A-2b compliance with Aspose.Cells | how to embed Windows fonts when exporting a spreadsheet to PDF/A-2b in .NET | load tsv data into Aspose.Cells workbook and export to archival PDF | Aspose.Cells PDF/A-2b export settings for multi‑page worksheets
// Tags: Aspose.Cells TSV to PDF/A-2b conversion | PdfSaveOptions for PDF/A-2b in C# | Windows font embedding Aspose.Cells PDF export | LoadOptions TSV format .NET | archival PDF generation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads a TSV file into an Aspose.Cells Workbook using LoadOptions, sets PdfSaveOptions for PDF/A‑2b compliance with embedded Windows fonts, and saves the workbook as an archival PDF/A‑2b document, handling missing input and output folder creation.
class TsvToPdfA2bConverter
{
    static void Main()
    {
        // Path to the source TSV file
        string tsvPath = @"C:\Input\sample.tsv";

        // Path for the resulting PDF/A‑2b file
        string pdfPath = @"C:\Output\sample.pdf";

        // Verify that the TSV source file exists
        if (!File.Exists(tsvPath))
        {
            Console.WriteLine($"Error: TSV file not found at '{tsvPath}'.");
            return;
        }

        // Ensure the output directory exists
        string outputDir = Path.GetDirectoryName(pdfPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        try
        {
            // Load the TSV file into a workbook using LoadOptions
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook workbook = new Workbook(tsvPath, loadOptions);

            // Configure PDF save options for PDF/A‑2b compliance
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA2b,          // PDF/A‑2b compliance
                EmbedStandardWindowsFonts = true,          // Embed fonts for archival quality
                AllColumnsInOnePagePerSheet = false        // Allow multi‑page sheets
            };

            // Save the workbook as a PDF/A‑2b document
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"PDF/A‑2b file successfully created at '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
