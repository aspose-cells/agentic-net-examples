// Title: Open a password‑protected Excel file and export to PDF with a diagonal semi‑transparent watermark using Aspose.Cells for .NET
// Description: C# sample that loads an encrypted .xlsx via LoadOptions, creates a red Arial watermark (45° rotation, 25 % opacity), attaches it with PdfSaveOptions, and saves the workbook as a PDF while handling CellsException and generic errors.
// Keywords: Aspose.Cells | C# encrypted workbook | load password protected Excel | Excel to PDF conversion | PDF watermark Aspose | PdfSaveOptions watermark | rotated text watermark | semi transparent watermark | LoadOptions password | Aspose.Cells error handling
// Common Searches: How to open a password‑protected Excel file with Aspose.Cells | Convert encrypted .xlsx to PDF with watermark in C# | Add diagonal CONFIDENTIAL watermark when saving Excel as PDF | Set watermark opacity and rotation in Aspose.Cells PdfSaveOptions | Catch CellsException for invalid workbook password
// Developer Intent: Load an encrypted Excel workbook using a password and generate a PDF that includes a rotated, semi‑transparent text watermark.
// Use Cases: Distribute confidential spreadsheets as watermarked PDFs to prevent unauthorized reuse. | Automate batch conversion of protected workbooks with corporate branding watermarks. | Create compliance‑ready reports from secured Excel files with a visible “CONFIDENTIAL” overlay.
// AI Prompts: Generate C# code that opens a password‑protected .xlsx with Aspose.Cells and saves it as a PDF with a custom diagonal watermark. | Explain how to configure watermark rotation, opacity, alignment, and scaling in PdfSaveOptions. | Show error‑handling patterns for invalid passwords or corrupted files when loading encrypted workbooks with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// C# sample that loads an encrypted .xlsx via LoadOptions, creates a red Arial watermark (45° rotation, 25 % opacity), attaches it with PdfSaveOptions, and saves the workbook as a PDF while handling CellsException and generic errors.
class Program
{
    static void Main()
    {
        // Path to the encrypted Excel workbook
        string inputFile = "encrypted.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        // Password required to open the workbook
        string workbookPassword = "myPassword";

        try
        {
            // Load the workbook with the specified password
            LoadOptions loadOpts = new LoadOptions
            {
                Password = workbookPassword
            };
            Workbook workbook = new Workbook(inputFile, loadOpts);

            // Create a font for the watermark text
            RenderingFont watermarkFont = new RenderingFont("Arial", 48)
            {
                Bold = true,
                Color = Color.Red
            };

            // Create a text watermark using the font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
            {
                Rotation = 45,               // Rotate watermark
                Opacity = 0.25f,             // Semi‑transparent
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                ScaleToPagePercent = 80      // Scale relative to page size
            };

            // Configure PDF save options to include the watermark
            PdfSaveOptions pdfOpts = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied
            string outputFile = "output.pdf";
            workbook.Save(outputFile, pdfOpts);
            Console.WriteLine($"PDF saved successfully to \"{outputFile}\".");
        }
        catch (CellsException ex)
        {
            // Handles errors such as invalid password or corrupted file
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
