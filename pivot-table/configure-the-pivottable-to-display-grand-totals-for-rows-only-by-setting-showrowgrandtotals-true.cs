// Title: How to enable row grand totals and hide column grand totals in an Aspose.Cells PivotTable using C#
// AI Prompts: Load an existing workbook, set the first PivotTable's ShowRowGrandTotals property to true and ShowColumnGrandTotals to false, then save the file with Aspose.Cells for .NET. | Write C# code that verifies a worksheet contains a PivotTable, turns on row grand totals only, and writes the updated workbook back to disk using Aspose.Cells.
// Common Searches: Aspose.Cells C# enable row grand totals only in pivot table | hide column grand totals Aspose.Cells pivot table programmatically | set ShowRowGrandTotals true for existing Excel pivot using Aspose.Cells .NET | modify pivot table grand total settings with Aspose.Cells C# | change pivot table totals in a workbook via Aspose.Cells API
// Tags: Aspose.Cells PivotTable ShowRowGrandTotals | C# set row grand totals Aspose.Cells | Aspose.Cells hide column grand totals | modify Excel pivot totals .NET | Aspose.Cells workbook pivot configuration

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads 'input.xlsx', accesses the first worksheet's first PivotTable, enables row grand totals, disables column grand totals, and saves the modified workbook as 'output.xlsx', including file existence checks and exception handling.
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
            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Get the worksheet that holds the PivotTable (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found in the worksheet.");
                return;
            }

            // Access the first PivotTable in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Show grand totals for rows only
            pivotTable.ShowRowGrandTotals = true;
            // Hide column grand totals (optional)
            pivotTable.ShowColumnGrandTotals = false;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
