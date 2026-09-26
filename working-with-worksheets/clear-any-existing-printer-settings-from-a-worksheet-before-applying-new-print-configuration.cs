// Title: Clear existing print area and set a new print range with fit‑to‑page settings using Aspose.Cells for .NET (C#)
// AI Prompts: Remove any current print area from a worksheet and define a new print range A1:D20 using Aspose.Cells in C#. | Configure the worksheet to fit to one page wide and unlimited pages tall after resetting the page setup with Aspose.Cells. | Ensure the output folder exists, then save the workbook with the updated print configuration to a specified file path using Aspose.Cells.
// Common Searches: Aspose.Cells C# clear worksheet print area before setting new range | set print area A1:D20 and fit to one page wide using Aspose.Cells .NET | reset page setup in Aspose.Cells then apply fit‑to‑pages settings | how to create output directory automatically when saving workbook with Aspose.Cells | remove previous printer settings in Excel file using Aspose.Cells for .NET
// Tags: clear existing print area Aspose.Cells | define new print range A1:D20 Aspose.Cells | fit worksheet to one page wide Aspose.Cells | auto‑create output folder before saving Aspose.Cells | reset page setup prior to print configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, clears any existing print area on the first worksheet, sets a new print range (A1:D20), configures fit‑to‑page settings (1 page wide, unlimited tall), ensures the output directory exists, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before loading
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Clear any existing print area
            sheet.PageSetup.PrintArea = null;

            // Apply new print configuration
            sheet.PageSetup.PrintArea = "A1:D20";

            // Set page orientation if the enum is available; otherwise skip
            // sheet.PageSetup.Orientation = PageOrientation.Portrait;

            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
