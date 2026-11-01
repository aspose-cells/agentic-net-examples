// Title: Log the count of worksheets that have no printer settings after batch clearing PageSetup properties with Aspose.Cells for .NET
// Common Searches: how to count worksheets with empty PrintArea after clearing printer settings using Aspose.Cells in C# | Aspose.Cells C# batch remove printer settings and log number of worksheets without print titles | log number of worksheets missing PageSetup printer configuration after reset with Aspose.Cells .NET | determine worksheets that lack printer settings after clearing PageSetup properties in a workbook | C# Aspose.Cells count worksheets without printer configuration after batch PageSetup cleanup
// Tags: batch clear worksheet printer settings Aspose.Cells | count worksheets lacking PageSetup printer configuration C# | log printer settings removal Aspose.Cells .NET | audit worksheet print area after PageSetup reset | reset PageSetup properties across workbook Aspose.Cells

using System;
using Aspose.Cells;

// The example loads a workbook, iterates through all worksheets to clear printer‑related PageSetup properties (PrintArea, PrintTitleRows, PrintTitleColumns), then counts how many worksheets have none of these settings defined, writes the count to the console, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Batch removal of printer settings from all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Clear common printer-related settings
            sheet.PageSetup.PrintArea = string.Empty;
            sheet.PageSetup.PrintTitleRows = string.Empty;
            sheet.PageSetup.PrintTitleColumns = string.Empty;
            // Additional printer settings can be cleared here if needed
        }

        // Count worksheets that now lack printer settings
        int worksheetsWithoutPrinterSettings = 0;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            bool hasPrintArea = !string.IsNullOrEmpty(sheet.PageSetup.PrintArea);
            bool hasTitleRows = !string.IsNullOrEmpty(sheet.PageSetup.PrintTitleRows);
            bool hasTitleColumns = !string.IsNullOrEmpty(sheet.PageSetup.PrintTitleColumns);

            if (!hasPrintArea && !hasTitleRows && !hasTitleColumns)
            {
                worksheetsWithoutPrinterSettings++;
            }
        }

        // Log the result
        Console.WriteLine($"Number of worksheets lacking printer settings: {worksheetsWithoutPrinterSettings}");

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
