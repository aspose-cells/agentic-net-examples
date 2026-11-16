// Title: How to count non‑empty cells in a defined range (A1:C10) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# count filled cells in range A1:C10 | C# get number of non‑blank cells in Excel worksheet using Aspose.Cells | Iterate over Aspose.Cells range and count cells with values | How to determine non‑empty cell count in a specific Excel range with Aspose.Cells .NET | Count non‑empty cells in a worksheet range without loading entire sheet Aspose.Cells
// Tags: populated cells count Aspose.Cells | create Excel range Aspose.Cells C# | iterate worksheet cells Aspose.Cells | cell value null verification Aspose.Cells | load and save workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, creates a range (A1:C10) on the first worksheet, iterates through each cell, increments a counter when the cell's Value is not null, prints the total number of populated cells, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range to inspect
            string rangeAddress = "A1:C10";

            // Create an Aspose.Cells.Range explicitly to avoid ambiguity with System.Range
            Aspose.Cells.Range range = worksheet.Cells.CreateRange(rangeAddress);

            // Count non‑empty cells in the range
            int nonEmptyCellCount = 0;
            foreach (Cell cell in range)
            {
                // A cell is considered non‑empty if its Value is not null
                if (cell.Value != null)
                {
                    nonEmptyCellCount++;
                }
            }

            Console.WriteLine($"Total non‑empty cells in range {rangeAddress}: {nonEmptyCellCount}");

            // Save the workbook (optional if modifications were made)
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            try
            {
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
