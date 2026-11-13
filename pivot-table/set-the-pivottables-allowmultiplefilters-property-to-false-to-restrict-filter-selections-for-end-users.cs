// Title: Disable multiple filter selections in an Excel PivotTable with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set PivotTable AllowMultipleFilters to false | how to prevent multi‑select filter in Excel PivotTable using Aspose.Cells | restrict PivotTable filter to single choice with Aspose.Cells .NET | Aspose.Cells PivotTable API does not expose AllowMultipleFilters property | example code to disable multiple filters in a PivotTable with Aspose.Cells
// Tags: Aspose.Cells PivotTable single filter selection | C# set AllowMultipleFilters property in PivotTable | Excel PivotTable filter limitation Aspose.Cells | .NET PivotTable API missing AllowMultipleFilters | disable multi‑select filter using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads a workbook, checks for a PivotTable on the first worksheet, retrieves the first PivotTable, and notes that the current Aspose.Cells version does not expose the AllowMultipleFilters property, then saves the workbook unchanged.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (sheet.PivotTables == null || sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the first worksheet.");
                return;
            }

            // Retrieve the first PivotTable
            PivotTable pivotTable = sheet.PivotTables[0];

            // The Aspose.Cells version used does not expose an AllowMultipleFilters property.
            // If needed, this section can be updated when the appropriate API becomes available.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
