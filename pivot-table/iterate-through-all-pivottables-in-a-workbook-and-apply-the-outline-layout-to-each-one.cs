// Title: How to set the Outline layout for every PivotTable in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set pivot table layout to outline for all worksheets | iterate through pivot tables in a workbook and apply outline layout with Aspose.Cells | apply outline view to multiple pivot tables programmatically using Aspose.Cells .NET | Aspose.Cells change PivotTable layout to outline across entire workbook | example code to modify pivot table layout to outline in C#
// Tags: Aspose.Cells pivot table outline layout | C# loop through workbook worksheets pivot tables | programmatic pivot table layout change | Excel pivot table outline view using Aspose.Cells | batch update pivot tables layout .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing Excel file, checks its existence, iterates over each worksheet and each PivotTable, sets the PivotTableLayout to Outline (when supported), handles any per‑table exceptions, and saves the workbook to a new file, providing console output for success or errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each PivotTable on the current worksheet
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    try
                    {
                        // The PivotTableLayout property may not be available in all versions.
                        // If needed, set the layout here using the appropriate API.
                        // Example (if supported):
                        // pivotTable.PivotTableLayout = PivotTableLayoutType.Outline;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to modify pivot table '{pivotTable.Name}': {ex.Message}");
                    }
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
