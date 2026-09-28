// Title: Count non‑empty cells in a defined worksheet range (B2:D10) using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that scans a given range (e.g., B2:D10) and returns the total number of cells that contain any value. | Modify the loop to treat cells containing only whitespace as blank and output the count of populated cells. | Create a reusable method that accepts a Worksheet object and start/end row‑column indices and returns the count of non‑blank cells using Aspose.Cells.
// Common Searches: Aspose.Cells count cells with data in a specific range C# | C# how to get number of filled cells between B2 and D10 using Aspose.Cells | skip empty and whitespace cells when counting Excel range with Aspose.Cells | determine non‑empty cell count in worksheet range programmatically Aspose.Cells | C# Aspose.Cells count non‑blank cells in B2:D10 example
// Tags: count non‑empty cells Aspose.Cells | worksheet range iteration Aspose.Cells | cell value type check Aspose.Cells | ignore whitespace cells Aspose.Cells | C# Excel range data count Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, accesses the first worksheet, iterates through cells B2:D10, checks each cell's type and string content to identify blanks (including empty strings), increments a counter for cells that contain data, and prints the total non‑empty cell count.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range B2:D10 (zero‑based indices)
            int startRow = 1;      // B2 -> row 1
            int startColumn = 1;   // B2 -> column 1
            int endRow = 9;        // D10 -> row 9
            int endColumn = 3;     // D10 -> column 3

            int nonEmptyCount = 0;

            // Iterate through each cell in the specified range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startColumn; col <= endColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Determine if the cell is blank:
                    // - Type IsNull indicates no value.
                    // - Empty string also considered blank.
                    bool isBlank = cell.Type == CellValueType.IsNull ||
                                   (cell.Type == CellValueType.IsString && string.IsNullOrEmpty(cell.StringValue));

                    if (!isBlank)
                    {
                        nonEmptyCount++;
                    }
                }
            }

            Console.WriteLine($"Non‑empty cells in range B2:D10: {nonEmptyCount}");

            // Optional: Save the workbook if modifications were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
