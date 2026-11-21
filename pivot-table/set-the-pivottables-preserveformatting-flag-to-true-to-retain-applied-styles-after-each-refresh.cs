// Title: Set PivotTable PreserveFormatting to true in Aspose.Cells for .NET (C#) to keep styles after refresh
// Common Searches: Aspose.Cells C# how to keep pivot table formatting after refresh | set PreserveFormatting property on a PivotTable using Aspose.Cells .NET | retain Excel pivot table styles when refreshing with Aspose.Cells library | C# code to enable PreserveFormatting for a PivotTable in a workbook | Aspose.Cells preserve pivot table formatting on workbook save
// Tags: Aspose.Cells PreserveFormatting pivot table | C# set pivot table formatting flag | Excel workbook pivot table style retention | Aspose.Cells PivotTable refresh formatting | preserve pivot table styles .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// Loads an existing Excel workbook, accesses the first PivotTable, sets its PreserveFormatting property to true so that applied styles survive each refresh, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing a PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Retrieve the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Preserve formatting after each refresh
            pivotTable.PreserveFormatting = true;

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
