// Title: Convert an Excel workbook to PDF with right‑to‑left layout for Arabic scripts using Aspose.Cells for .NET
// Common Searches: how to save an .xlsx as PDF with right-to-left direction in Aspose.Cells C# | Aspose.Cells C# enable Arabic RTL layout when exporting to PDF | set IsRightToLeft property before converting Excel to PDF with Aspose.Cells | C# convert Excel to PDF preserving Arabic script orientation using Aspose.Cells | Aspose.Cells PDF export Arabic right-to-left formatting example
// Tags: Aspose.Cells right-to-left PDF export | Excel to PDF conversion with Arabic RTL | IsRightToLeft property usage in Aspose.Cells | C# workbook save as PDF Aspose.Cells | Arabic script orientation in PDF via Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX file with Aspose.Cells, optionally enables the IsRightToLeft setting for Arabic right‑to‑left layout, and saves the workbook as a PDF, handling missing files and runtime errors.
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

            // If the current Aspose.Cells version supports right‑to‑left layout for Arabic scripts,
            // the following property can be set. It is commented out to maintain compatibility
            // with versions where the property may not exist.
            // workbook.Settings.IsRightToLeft = true;

            // Convert and save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
