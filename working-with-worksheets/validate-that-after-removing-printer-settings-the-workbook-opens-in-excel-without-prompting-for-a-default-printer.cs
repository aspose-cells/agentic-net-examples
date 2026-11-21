// Title: Remove printer settings from all worksheets to stop Excel default printer prompt using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clear worksheet print area and title rows to avoid printer dialog | How to disable printer prompts when opening an Excel file with Aspose.Cells | Remove PageSetup printer configuration from workbook programmatically in .NET | C# code to strip printer settings from XLSX before saving with Aspose.Cells | Prevent Excel from asking for default printer after loading workbook via Aspose.Cells
// Tags: clear worksheet PageSetup printer settings Aspose.Cells | remove print area title rows C# | disable Excel printer prompt Aspose.Cells | batch process worksheets printer configuration .NET | save workbook without printer defaults Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the existence of input.xlsx, loads it with Aspose.Cells, iterates through each worksheet to clear PageSetup.PrintArea, PrintTitleRows, and PrintTitleColumns, and then saves the cleaned workbook as output.xlsx while handling any exceptions.
class RemovePrinterSettings
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and clear printer‑related settings
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear print area and title rows/columns (common printer settings)
                sheet.PageSetup.PrintArea = string.Empty;
                sheet.PageSetup.PrintTitleRows = string.Empty;
                sheet.PageSetup.PrintTitleColumns = string.Empty;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
