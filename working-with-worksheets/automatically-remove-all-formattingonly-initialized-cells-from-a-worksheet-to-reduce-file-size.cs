using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormattingCleanup
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Determine the used range (including the last used row/column)
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Iterate through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // If the cell has no value (empty) but may have formatting, clear its formats
                        if (cell.IsEmpty)
                        {
                            // Clear formatting for this single cell
                            cells.ClearFormats(row, col, row, col);
                        }
                    }
                }

                // Remove any styles that are no longer used after clearing formats
                workbook.RemoveUnusedStyles();

                // Save the optimized workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}