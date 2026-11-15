// Title: Count and log worksheets with no print area after clearing printer settings using Aspose.Cells for .NET
// Common Searches: how to count worksheets that have a null PrintArea after resetting PageSetup with Aspose.Cells C# | Aspose.Cells C# remove printer settings from all worksheets and get number of sheets without print area | log number of worksheets missing printer configuration after batch PageSetup clear in Aspose.Cells | C# Aspose.Cells count worksheets where PrintArea is empty after clearing printer properties
// Tags: clear worksheet printer settings Aspose.Cells | count worksheets with null PrintArea C# | log missing printer configuration Aspose.Cells | batch PageSetup reset Aspose.Cells | Aspose.Cells workbook printer settings cleanup

using System;
using Aspose.Cells;

// The example loads an Excel workbook, clears printer‑related PageSetup properties for every worksheet, counts how many worksheets have a null or empty PrintArea, writes this count to the console, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (provided rule)
        var workbook = new Workbook("input.xlsx"); // {LoadWorkbook}

        // Batch removal of printer settings
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.PrintArea = null;
            sheet.PageSetup.PrintTitleRows = null;
            sheet.PageSetup.PrintTitleColumns = null;
            sheet.PageSetup.CenterHorizontally = false;
            sheet.PageSetup.CenterVertically = false;
            // Additional printer-related properties can be cleared here if needed
        }

        // Count worksheets that now lack printer settings (PrintArea is null or empty)
        int lackingCount = 0;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (string.IsNullOrEmpty(sheet.PageSetup.PrintArea))
            {
                lackingCount++;
            }
        }

        // Log the result
        Console.WriteLine($"Number of worksheets lacking printer settings: {lackingCount}");

        // Save the workbook (provided rule)
        workbook.Save("output.xlsx", SaveFormat.Xlsx); // {SaveWorkbook}
    }
}
