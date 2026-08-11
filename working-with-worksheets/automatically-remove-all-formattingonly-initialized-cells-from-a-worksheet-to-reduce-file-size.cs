// Title: Remove Formatting‑Only Cells with Aspose.Cells (C#) to Shrink Excel Files
// Description: Loads a workbook, iterates the used range, clears formatting from empty cells, removes any now‑unused styles, and saves the file, resulting in a smaller Excel workbook.
// Keywords: Aspose.Cells clear formatting | remove formatting only cells | optimize Excel size .NET | remove unused styles | clear blank cell formats | C# Aspose.Cells workbook optimization | reduce Excel file size | delete cell styles Aspose | Excel performance tuning
// Common Searches: how to clear formatting from empty cells using Aspose.Cells C# | remove unused styles after clearing formats Aspose.Cells | shrink Excel workbook size by deleting formatting only cells | Aspose.Cells iterate used range to clear cell formats | C# code to clean up blank cell styles in Excel
// Developer Intent: Clear all formatting from cells that contain no data and purge unused styles to reduce the workbook’s file size.
// Use Cases: Prepare a report workbook for distribution by stripping unnecessary formatting from placeholder cells. | Automate cleanup of generated spreadsheets so blank cells do not retain redundant styles, lowering storage costs. | Integrate formatting cleanup into a CI/CD pipeline that processes Excel files, ensuring each published workbook is size‑optimized.
// AI Prompts: Write C# code with Aspose.Cells that clears formats of all blank cells in a worksheet and then calls RemoveUnusedStyles. | Suggest a more efficient method to identify formatting‑only cells without scanning every cell in the used range. | Explain the purpose of Workbook.RemoveUnusedStyles and the best time to invoke it after modifying cell formats.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, iterates the used range, clears formatting from empty cells, removes any now‑unused styles, and saves the file, resulting in a smaller Excel workbook.
    public class RemoveFormattingOnlyCells
    {
        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Determine the used range of the worksheet
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Iterate through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        bool isBlank = cell.Value == null || string.IsNullOrEmpty(cell.StringValue);

                        if (isBlank)
                        {
                            // Clear only the formatting of the blank cell
                            cells.ClearFormats(row, col, row, col);
                        }
                    }
                }

                // Remove any styles that are now unused after clearing formats
                workbook.RemoveUnusedStyles();

                // Save the optimized workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveFormattingOnlyCells.Run();
        }
    }
}
