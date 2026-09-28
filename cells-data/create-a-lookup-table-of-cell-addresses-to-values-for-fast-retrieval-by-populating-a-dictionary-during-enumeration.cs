// Title: Build a case‑insensitive dictionary of Excel cell addresses to values using Aspose.Cells in C# for rapid lookup
// AI Prompts: Generate C# code that loads a workbook with Aspose.Cells, iterates the worksheet’s used range, and adds each non‑empty cell’s address (e.g., "B3") and its value to a Dictionary<string, object> using StringComparer.OrdinalIgnoreCase. | Demonstrate how to query the dictionary for a specific cell address, handling missing keys gracefully and printing either the retrieved value or a fallback message. | Add error handling that creates a new workbook when the input file is absent, builds the lookup dictionary, and saves the workbook after processing.
// Common Searches: aspocells c# create dictionary of cell address to value for fast lookup | how to use StringComparer.OrdinalIgnoreCase with Aspose.Cells cell address keys | enumerate used range in Aspose.Cells and store non‑empty cells in a map | retrieve Excel cell value by address using a pre‑built lookup table in C#
// Tags: Aspose.Cells used range to dictionary | C# case‑insensitive Excel cell key | non‑empty cell value map Aspose.Cells | fast cell lookup table C# | create workbook if missing Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Cells.Range;

// The program loads or creates an Excel workbook, iterates the first worksheet's used range, stores each non‑empty cell's address and value in a case‑insensitive Dictionary for O(1) retrieval, demonstrates a lookup example, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

            // Prepare a lookup dictionary for cell address → value
            var lookup = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            // Enumerate cells only if a used range exists and contains cells
            if (usedRange != null && usedRange.RowCount > 0 && usedRange.ColumnCount > 0)
            {
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Value != null)
                        {
                            // Use the cell's address (e.g., "B3") as the key
                            lookup[cell.Name] = cell.Value;
                        }
                    }
                }
            }

            // Example of fast retrieval using the dictionary
            if (lookup.TryGetValue("B2", out object retrievedValue))
            {
                Console.WriteLine($"Value at B2: {retrievedValue}");
            }
            else
            {
                Console.WriteLine("Cell B2 is empty or does not exist.");
            }

            // Save the workbook (creates the file if it does not exist)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
