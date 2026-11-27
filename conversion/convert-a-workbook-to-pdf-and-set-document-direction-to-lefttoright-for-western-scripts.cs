// Title: Convert an Excel workbook to PDF with left‑to‑right layout using Aspose.Cells for .NET (C#)
// Common Searches: c# asp.net convert xlsx file to pdf preserving left-to-right orientation with Aspose.Cells | how to export an Excel workbook to PDF without right-to-left direction using Aspose.Cells | Aspose.Cells save workbook as PDF example with default left‑to‑right layout | c# code to check if input.xlsx exists before converting to PDF with Aspose.Cells | using SaveFormat.Pdf to generate PDF from Excel in Aspose.Cells C#
// Tags: Aspose.Cells Excel to PDF conversion C# | left-to-right document direction Aspose.Cells PDF export | SaveFormat.Pdf usage Aspose.Cells example | input file existence validation C# Aspose.Cells | disable right-to-left layout Aspose.Cells PDF generation

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates loading an existing Excel file, verifying its presence, and saving it as a PDF with Aspose.Cells for .NET. The example relies on the default left‑to‑right document direction, ensuring no right‑to‑left layout is applied.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // The default document direction is left‑to‑right; no explicit property needed
            // (If using a newer API that supports it, you could set workbook.Settings.IsRightToLeft = false;)

            // Convert and save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully converted to PDF: \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions (e.g., loading, saving) and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
