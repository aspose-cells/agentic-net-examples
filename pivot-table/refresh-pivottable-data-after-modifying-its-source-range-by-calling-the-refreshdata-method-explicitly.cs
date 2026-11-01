// Title: Refresh a PivotTable after modifying its source range with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# refresh PivotTable after changing source cells | How to call RefreshData on a PivotTable using Aspose.Cells .NET | Update Excel PivotTable programmatically after editing source data with Aspose.Cells | RefreshData method example for PivotTable in Aspose.Cells for .NET | Refresh PivotTable when source range is edited in C# using Aspose.Cells
// Tags: Aspose.Cells RefreshData PivotTable | Aspose.Cells modify source range | Aspose.Cells programmatic pivot refresh | Aspose.Cells .NET update pivot data | Aspose.Cells C# pivot table refresh

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, changes a cell in the source data range, retrieves the first PivotTable on the sheet, calls the RefreshData method to synchronize the PivotTable with the updated source, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet with the source data (first sheet in this example)
            Worksheet dataSheet = workbook.Worksheets[0];

            // Example modification of the source range: change a cell value
            dataSheet.Cells["A2"].PutValue(12345);

            // Access the worksheet that holds the PivotTable (same sheet here)
            Worksheet pivotSheet = workbook.Worksheets[0];

            // Ensure at least one PivotTable exists
            if (pivotSheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found on the worksheet.");
                return;
            }

            // Retrieve the first PivotTable
            PivotTable pivotTable = pivotSheet.PivotTables[0];

            // Refresh the PivotTable data after source change
            pivotTable.RefreshData();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
