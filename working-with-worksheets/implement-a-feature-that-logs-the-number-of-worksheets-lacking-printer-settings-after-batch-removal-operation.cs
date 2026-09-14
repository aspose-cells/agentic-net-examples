// Title: Count and log worksheets that have no printer settings after clearing them with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to remove printer settings from every worksheet in a workbook, then determine how many worksheets have a null PrinterSettings property and output the count. | Generate a .NET example that iterates through all worksheets, sets PageSetup.PrinterSettings to null, counts sheets lacking printer configuration, logs the result, and saves the workbook.
// Common Searches: Aspose.Cells C# count worksheets with null PageSetup.PrinterSettings after batch clear | how to log number of Excel sheets without printer settings using Aspose.Cells | remove printer settings from all worksheets and get count in .NET | C# example for clearing PageSetup.PrinterSettings and reporting affected worksheets | Aspose.Cells batch clear printer configuration and verify sheets
// Tags: Aspose.Cells clear worksheet printer settings | count worksheets null PrinterSettings Aspose.Cells | log worksheet printer configuration status C# | batch PageSetup.PrinterSettings removal .NET | iterate workbook worksheets Aspose.Cells

using System;
using Aspose.Cells;

// The program loads an Excel workbook, clears the printer settings of each worksheet by setting PageSetup.PrinterSettings to null, counts how many worksheets now have a null PrinterSettings property, writes that count to the console, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Batch removal of printer settings from all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Remove printer settings by setting the property to null
            sheet.PageSetup.PrinterSettings = null;
        }

        // Count worksheets that now lack printer settings
        int worksheetsWithoutPrinterSettings = 0;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.PageSetup.PrinterSettings == null)
            {
                worksheetsWithoutPrinterSettings++;
            }
        }

        // Log the result
        Console.WriteLine($"Number of worksheets without printer settings: {worksheetsWithoutPrinterSettings}");

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
