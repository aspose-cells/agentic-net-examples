// Title: Unmerge cells B2:C2, copy its original value to D2, and save the workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unmerge a specific range while preserving the original value | Copy value from merged cells to another cell after unmerging with Aspose.Cells | C# code to unmerge B2:C2 in an existing Excel file and write the value to D2 using Aspose.Cells | How to save changes after unmerging cells in a workbook with Aspose.Cells for .NET | Retrieve top‑left value of a merged range before unmerging in Aspose.Cells
// Tags: unmerge range Aspose.Cells C# | get merged cell top-left value Aspose.Cells | copy cell value to adjacent cell Aspose.Cells | save updated Excel workbook Aspose.Cells | verify input file existence C#

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, captures the value from the top‑left cell of the merged range B2:C2, unmerges that range, writes the captured value to the adjacent cell D2, and saves the modified file, while handling missing input files and runtime exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Define the merged range B2:C2 (zero‑based indices)
                int startRow = 1;      // B2 row index
                int startColumn = 1;   // B column index
                int totalRows = 1;     // only one row
                int totalColumns = 2;  // B and C columns

                // Retrieve the value from the top‑left cell of the merged range (B2)
                Cell mergedCell = sheet.Cells[startRow, startColumn];
                object originalValue = mergedCell.Value;

                // Unmerge the cells B2:C2 using the Range API
                AsposeRange mergedRange = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);
                mergedRange.UnMerge();

                // Copy the original value to an adjacent cell (e.g., D2)
                int adjacentColumn = startColumn + totalColumns; // D column index (3)
                Cell adjacentCell = sheet.Cells[startRow, adjacentColumn];
                adjacentCell.PutValue(originalValue);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
