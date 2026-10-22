// Title: Reorder PivotTable report filters (page fields) to change filter priority with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to move a PivotTable page field from index 1 to index 0, changing the filter priority. | Programmatically adjust the order of report filters in an existing Excel PivotTable using the PageFields.Move method. | Rewrite the filter sequence of a PivotTable by reordering its page fields with Aspose.Cells for .NET. | Change the display order of PivotTable page fields in a workbook and save the result using C#.
// Common Searches: how to change the order of pivot table report filters using Aspose.Cells C# | Aspose.Cells PageFields.Move example for reordering pivot filters | programmatically set pivot table filter priority in .NET | move pivot table page field to first position with Aspose.Cells | reorder Excel pivot table page fields via C# code
// Tags: Aspose.Cells PageFields.Move | C# reorder pivot table page fields | Excel pivot filter priority .NET | programmatic pivot report filter ordering | modify pivot table filter sequence Aspose.Cells | change pivot table page field index C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing workbook, accesses the first worksheet and its first PivotTable, then calls PageFields.Move to shift the page field at index 1 to index 0, thereby changing the filter priority, and finally saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet that contains the PivotTable
            Worksheet worksheet = workbook.Worksheets["Sheet1"];
            if (worksheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet1' not found.");
                return;
            }

            // Ensure there is at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Get the first PivotTable in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Reorder page fields (report filters) to change filter priority
            // Example: move the page field currently at index 1 to index 0
            if (pivotTable.PageFields.Count > 1)
            {
                try
                {
                    // Use the Move method of the PageFields collection (Position is read‑only)
                    pivotTable.PageFields.Move(1, 0);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to reorder page fields: {ex.Message}");
                }
            }

            // Ensure the output directory exists (if any)
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
