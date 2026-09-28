// Title: How to delete a pivot table by zero‑based index from a worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing .xlsx file with Aspose.Cells, checks the number of pivot tables on the first worksheet, removes the pivot table at index 2 using Worksheet.PivotTables.RemoveAt, and saves the workbook. | Show a C# example that safely deletes a pivot table by verifying the provided zero‑based index against Worksheet.PivotTables.Count before calling RemoveAt. | Demonstrate how to log a friendly message when the requested pivot table index is out of range and continue processing.
// Common Searches: aspnet remove pivot table at specific index using Aspose.Cells | c# Aspose.Cells delete pivot table from worksheet by index | how to check pivot table count before removing with Worksheet.PivotTables.RemoveAt | example of Worksheet.PivotTables.RemoveAt in Aspose.Cells .NET | error handling for out‑of‑range pivot table index Aspose.Cells
// Tags: Aspose.Cells Worksheet.PivotTables.RemoveAt | remove pivot table at specific index | pivot table index validation Aspose.Cells | delete pivot table from .xlsx using .NET | exception handling for pivot table removal

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads 'input.xlsx' with Aspose.Cells, confirms the file exists, accesses the first worksheet, validates that a given zero‑based pivot table index is within the collection bounds, removes the pivot table using Worksheet.PivotTables.RemoveAt, and saves the updated workbook to 'output.xlsx' while handling possible errors and logging actions.
    class Program
    {
        static void Main(string[] args)
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

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Zero‑based index of the pivot table to delete
                int pivotIndex = 0; // change as required

                // Ensure the pivot table index is valid
                if (pivotIndex >= 0 && pivotIndex < worksheet.PivotTables.Count)
                {
                    worksheet.PivotTables.RemoveAt(pivotIndex);
                    Console.WriteLine($"Pivot table at index {pivotIndex} removed.");
                }
                else
                {
                    Console.WriteLine($"No pivot table found at index {pivotIndex}.");
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
}
