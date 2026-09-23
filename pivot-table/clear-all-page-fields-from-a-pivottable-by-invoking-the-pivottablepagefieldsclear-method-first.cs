// Title: How to clear all page (filter) fields from a PivotTable in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, accesses the first PivotTable, calls PageFields.Clear to remove every page field, and saves the workbook. | Show step‑by‑step how to programmatically reset PivotTable filters in a .NET application using the Aspose.Cells PivotTable API. | Provide a concise example that demonstrates using Aspose.Cells to clear all page fields of a PivotTable before exporting the workbook.
// Common Searches: aspnet aspose.cells clear pivot table page fields c# | remove all pivot table filters programmatically using Aspose.Cells .NET | c# aspose.cells how to reset pivot table page fields | example code for PivotTable.PageFields.Clear in Aspose.Cells | clear pivot table page filters before saving workbook with Aspose.Cells
// Tags: Aspose.Cells PivotTable PageFields.Clear | C# remove pivot table page fields | Aspose.Cells reset pivot filters | Excel workbook modify pivot page fields .NET | Aspose.Cells clear pivot filters programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, retrieves the first worksheet's first PivotTable, clears all page (filter) fields using the PageFields.Clear method, and saves the updated workbook to a new file.
class ClearPivotTablePageFields
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Get the first PivotTable (adjust index if required)
            PivotTable pivotTable = sheet.PivotTables[0];

            // Clear all page fields (filters) from the PivotTable
            pivotTable.PageFields.Clear();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
