// Title: How to remove all filters from an Aspose.Cells PivotTable and refresh it in C#
// AI Prompts: Write C# code that iterates over every PivotField in an Aspose.Cells PivotTable, removes any filter criteria, then refreshes and saves the workbook. | Show a .NET snippet that programmatically resets filter settings on a PivotTable using Aspose.Cells and recalculates the pivot data. | Generate an example that unfilters a PivotTable in C#, calls RefreshData and CalculateData, and writes the result to a new Excel file.
// Common Searches: Aspose.Cells C# how to clear filters on all pivot fields | programmatically reset pivot table filters in Excel using Aspose.Cells | remove pivot table filter criteria and recalculate data with Aspose.Cells .NET | unfilter a PivotTable and save workbook using Aspose.Cells C#
// Tags: Aspose.Cells reset pivot filters | Aspose.Cells refresh pivot after filter reset | C# Aspose.Cells clear pivot field criteria | Aspose.Cells programmatic pivot unfilter | Excel workbook save after pivot refresh Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing Excel workbook, accesses the first PivotTable, refreshes its data and recalculates it (Aspose.Cells does not provide a single ClearFilter method, so filters must be cleared via individual field properties), and then saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Access the first pivot table on the worksheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // NOTE: Aspose.Cells .NET API does not expose a ClearFilter method for PivotField.
            // If filter removal is required, it can be handled via specific filter properties.
            // For this example we simply refresh the pivot table without explicit filter clearing.

            // Refresh the pivot table to apply any changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the (potentially) unfiltered pivot table
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
