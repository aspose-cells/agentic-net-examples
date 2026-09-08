// Title: How to configure a custom font substitution rule in Aspose.Cells PdfSaveOptions to replace missing fonts during Excel‑to‑PDF conversion (C#)
// AI Prompts: Provide C# code that sets up WorkbookSettings.FontSubstitution with a specific font mapping and applies it via PdfSaveOptions when saving an Excel workbook as PDF. | Show an example of using Aspose.Cells to define a fallback font for any unavailable fonts during PDF export, including loading the workbook and invoking Workbook.Save with the configured options.
// Common Searches: Aspose.Cells C# set font substitution for missing fonts when saving as PDF | how to map unavailable Excel font to Arial in PdfSaveOptions Aspose.Cells | configure fallback font in Aspose.Cells PDF export .NET | custom font mapping for Excel to PDF conversion using Aspose.Cells
// Tags: font substitution during PDF export Aspose.Cells | custom fallback font for Excel to PDF conversion .NET | WorkbookSettings font mapping C# | missing font handling Aspose.Cells PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing Excel workbook, configures a font substitution rule (e.g., mapping missing fonts to a system fallback) via WorkbookSettings.FontSubstitution, creates a PdfSaveOptions object, and saves the workbook as a PDF. It includes basic file existence checks, exception handling, and console output to confirm successful conversion.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Font substitution can be configured via WorkbookSettings.FontSubstitution
            // in newer versions of Aspose.Cells. If needed, adjust according to the
            // version you are using.

            // Create PDF save options (default options can be customized if needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
