// Title: C# – Count Non‑Empty Cells in a Worksheet’s MaxDisplayRange with Aspose.Cells
// Description: Loads an Excel workbook, retrieves the worksheet’s MaxDisplayRange (covering data, merged cells and shapes), iterates through each cell, counts those with a non‑null and non‑empty value, outputs the total, and optionally saves the file.
// Keywords: Aspose.Cells | C# | MaxDisplayRange | count non‑empty cells | enumerate cells | Excel data validation | worksheet cell count
// Common Searches: how to count populated cells in MaxDisplayRange Aspose.Cells | C# iterate cells in MaxDisplayRange | Aspose.Cells count non‑empty cells in worksheet | enumerate MaxDisplayRange cells .NET | Excel workbook count filled cells using Aspose
// Developer Intent: Find a way to determine the number of cells that contain data within a worksheet’s MaxDisplayRange.
// Use Cases: Verify that a sheet meets a minimum data threshold before further processing. | Generate a quick data‑quality metric for reporting or logging. | Conditionally trigger export or saving actions based on the count of filled cells.
// AI Prompts: Write C# code with Aspose.Cells that counts non‑empty cells in the MaxDisplayRange and returns the count. | Show how to modify the loop to ignore merged cells while counting populated cells. | Provide a snippet that logs each non‑empty cell’s address and value during MaxDisplayRange iteration.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, retrieves the worksheet’s MaxDisplayRange (covering data, merged cells and shapes), iterates through each cell, counts those with a non‑null and non‑empty value, outputs the total, and optionally saves the file.
    public class CountNonEmptyCellsInMaxDisplayRange
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the maximum display range that includes data, merged cells and shapes
                Aspose.Cells.Range maxDisplayRange = worksheet.Cells.MaxDisplayRange;

                int nonEmptyCellCount = 0;

                if (maxDisplayRange != null)
                {
                    // Enumerate all cells in the range
                    IEnumerator enumerator = maxDisplayRange.GetEnumerator();
                    while (enumerator.MoveNext())
                    {
                        Cell cell = (Cell)enumerator.Current;

                        // Count cells that have a non‑null and non‑empty value
                        if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                        {
                            nonEmptyCellCount++;
                        }
                    }
                }

                Console.WriteLine($"Non‑empty cells in MaxDisplayRange: {nonEmptyCellCount}");

                // Save the workbook (optional)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            CountNonEmptyCellsInMaxDisplayRange.Run();
        }
    }
}
