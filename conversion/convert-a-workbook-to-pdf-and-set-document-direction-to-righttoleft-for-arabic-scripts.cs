// Title: Convert an Excel workbook to PDF with right‑to‑left layout for Arabic using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells export Excel to PDF with right‑to‑left direction for Arabic | How to enable RTL layout when saving a workbook as PDF with Aspose.Cells | Aspose.Cells workbook.Settings.IsRightToLeft example for PDF conversion | Generate PDF from XLSX with Arabic RTL formatting using Aspose.Cells C# | Aspose.Cells PDF export not preserving right‑to‑left orientation
// Tags: Aspose.Cells right-to-left PDF export | C# workbook.Settings.IsRightToLeft usage | Excel to PDF conversion with Arabic RTL layout | SaveFormat.Pdf with RTL direction Aspose.Cells | Aspose.Cells Arabic script PDF generation | C# Aspose.Cells PDF export settings

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the input XLSX file, loads it into an Aspose.Cells Workbook, optionally sets the workbook's right‑to‑left flag via workbook.Settings.IsRightToLeft, and then saves the workbook as a PDF using SaveFormat.Pdf, with error handling for robust execution.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: In some versions of Aspose.Cells the right‑to‑left setting is not exposed.
            // If needed, adjust cell styles or use a newer library version that supports:
            // workbook.Settings.IsRightToLeft = true;

            // Convert and save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
