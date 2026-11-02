// Title: How to count non‑empty cells in a worksheet’s MaxDisplayRange with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# count non empty cells in MaxDisplayRange | iterate through worksheet MaxDisplayRange to find populated cells using Aspose.Cells | C# get number of filled cells in Excel sheet with Aspose.Cells MaxDisplayRange | enumerate cells in MaxDisplayRange and count values in Aspose.Cells | count cells with data in Aspose.Cells workbook using MaxDisplayRange
// Tags: maxdisplayrange cell enumeration Aspose.Cells | count nonempty cells Aspose.Cells .NET | worksheet populated cell counting Aspose.Cells | iterate over Excel cells using MaxDisplayRange | Aspose.Cells data presence detection

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, retrieves the first worksheet's MaxDisplayRange, iterates through each cell in that range, increments a counter for cells whose Value is not null and whose StringValue is not empty, prints the total number of non‑empty cells, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
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

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the maximum display range of the worksheet (use fully qualified type to avoid ambiguity)
            Aspose.Cells.Range maxDisplayRange = worksheet.Cells.MaxDisplayRange;

            int nonEmptyCellCount = 0;

            // Iterate through each cell in the maximum display range
            foreach (Cell cell in maxDisplayRange)
            {
                // Count the cell if it contains a non‑null value and is not an empty string
                if (cell.Value != null && !string.IsNullOrEmpty(cell.StringValue))
                {
                    nonEmptyCellCount++;
                }
            }

            // Output the result
            Console.WriteLine($"Non‑empty cells count: {nonEmptyCellCount}");

            // Save the workbook if needed (optional)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
