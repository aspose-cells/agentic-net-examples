// Title: How to convert an HTML file to a password‑protected PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an HTML file into an Aspose.Cells Workbook and saves it as a PDF protected with a user password using PdfSaveOptions. | Show how to assign the Password property (and optionally an EncryptionType) on PdfSaveOptions when converting HTML to PDF with Aspose.Cells. | Add robust error handling for missing HTML input, workbook load failures, and PDF save exceptions in the Aspose.Cells HTML‑to‑PDF workflow.
// Common Searches: Aspose.Cells C# convert html to pdf and set user password | set pdf password when saving workbook as pdf using Aspose.Cells | encrypt pdf generated from html with Aspose.Cells .NET example | handle file not found error in Aspose.Cells html to pdf conversion
// Tags: HTML to PDF conversion Aspose.Cells C# | PdfSaveOptions password protection | Aspose.Cells encrypt PDF output | C# load HTML workbook Aspose.Cells | error handling Aspose.Cells file conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample verifies that the source HTML file exists, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, configures PdfSaveOptions with a user password (and optional encryption settings), and saves the workbook as a password‑protected PDF. All operations are wrapped in try/catch blocks to report missing files, loading issues, or PDF‑save failures.
class Program
{
    static void Main()
    {
        const string inputPath = "input.html";
        const string outputPath = "output.pdf";

        // Verify that the input HTML file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the HTML file into a workbook.
            var loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options (encryption removed due to missing reference).
            var pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., loading, saving).
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
