// Title: Convert a Password‑Protected Excel Workbook to PDF with a Diagonal Watermark using Aspose.Cells (C#)
// Description: C# example that checks for an encrypted Excel file, opens it with LoadOptions.Password, creates a semi‑transparent red diagonal watermark, assigns it to PdfSaveOptions, and saves the workbook as a PDF. Includes robust error handling for missing files and invalid passwords.
// Keywords: Aspose.Cells | C# | load encrypted workbook | password protected Excel | Excel to PDF conversion | PDF watermark | diagonal watermark | PdfSaveOptions | RenderingWatermark | secure document export
// Common Searches: Aspose.Cells open password protected .xlsx | convert encrypted Excel to PDF C# | add diagonal CONFIDENTIAL watermark to PDF with Aspose.Cells | PdfSaveOptions watermark example | handle invalid password error Aspose.Cells
// Developer Intent: Load a password‑protected Excel file and export it to a PDF that displays a rotated, semi‑transparent watermark.
// Use Cases: Distribute confidential reports as watermarked PDFs while preserving original workbook protection. | Automate batch conversion of secured spreadsheets into branded PDF documents. | Create audit‑ready PDFs that clearly show a confidentiality stamp on each page.
// AI Prompts: Generate C# code with Aspose.Cells to open an encrypted .xlsx file and save it as a PDF with a custom diagonal watermark. | Explain how to catch CellsException for an incorrect password when loading a protected workbook. | Show how to modify watermark opacity, rotation angle, font, and alignment in PdfSaveOptions.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// C# example that checks for an encrypted Excel file, opens it with LoadOptions.Password, creates a semi‑transparent red diagonal watermark, assigns it to PdfSaveOptions, and saves the workbook as a PDF. Includes robust error handling for missing files and invalid passwords.
class Program
{
    static void Main()
    {
        const string inputPath = "EncryptedWorkbook.xlsx";
        const string outputPath = "EncryptedWorkbook_WithWatermark.pdf";
        const string workbookPassword = "YourPassword";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the password‑protected workbook
            LoadOptions loadOptions = new LoadOptions
            {
                Password = workbookPassword
            };
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws CellsException for invalid password or other load errors
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error loading workbook: {ex.Message}");
            return;
        }

        if (workbook == null)
        {
            Console.WriteLine("Error: Workbook could not be loaded.");
            return;
        }

        try
        {
            // Create a font for the watermark text
            RenderingFont watermarkFont = new RenderingFont("Arial", 48)
            {
                Bold = true,
                Color = Color.Red
            };

            // Create a text watermark
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
            {
                Rotation = 45,                     // rotate the watermark
                Opacity = 0.3f,                    // make it semi‑transparent
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center
            };

            // Configure PDF save options to include the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during PDF generation: {ex.Message}");
        }
    }
}
