// Title: How to hide column grand totals in an Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide column grand totals in PivotTable | set ShowColumnGrandTotals false Aspose.Cells example | remove column grand total from Excel PivotTable programmatically | disable column grand totals when saving workbook with Aspose.Cells | C# code to turn off column grand totals in a PivotTable using Aspose
// Tags: Aspose.Cells ShowColumnGrandTotals API | PivotTable grand total visibility control | C# modify Excel PivotTable totals | Excel workbook save after PivotTable changes | Aspose.Cells PivotTable configuration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The sample loads an existing Excel workbook, checks for a PivotTable on the first worksheet, sets the PivotTable's ShowColumnGrandTotals property to false (optionally also the row totals), and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook containing a PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (sheet.PivotTables.Count > 0)
            {
                PivotTable pivotTable = sheet.PivotTables[0];

                // Disable column grand totals
                pivotTable.ShowColumnGrandTotals = false;

                // Optionally disable row grand totals
                // pivotTable.ShowRowGrandTotals = false;
            }
            else
            {
                Console.WriteLine("No PivotTable found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
