// Title: How to remove formatting‑only cells from an Excel worksheet using Aspose.Cells for .NET to reduce file size
// Common Searches: Aspose.Cells C# remove cell formatting when cell is empty | Delete empty cells that only have style in a .NET workbook | Reduce Excel file size by clearing unused cell styles with Aspose.Cells | Reset style of blank cells to default using Aspose.Cells API | Iterate worksheets to clear formatting‑only cells in C#
// Tags: clear formatting‑only cells Aspose.Cells .NET | remove empty cell styles Excel workbook | optimize workbook size by cleaning cell formatting | reset cell style to default using SetStyle | iterate worksheets to clean unused formatting

using System;
using System.IO;
using Aspose.Cells;

namespace RemoveFormattingOnlyCells
{
    // The example loads an Excel workbook, walks through each worksheet's used range, and for every cell that contains neither a value nor a formula it applies a fresh default style via SetStyle, effectively removing formatting‑only cells and shrinking the resulting file size.
    class Program
    {
        /// <param name="args">
        /// args[0] - path to the input Excel file (required)
        /// args[1] - path to the output Excel file (optional, defaults to "output.xlsx")
        /// </param>
        static void Main(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Usage: RemoveFormattingOnlyCells <inputFilePath> [outputFilePath]");
                return;
            }

            string inputPath = args[0];
            string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet in the workbook.
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    Cells cells = worksheet.Cells;

                    // MaxRow and MaxColumn include cells that have any content, including formatting.
                    // If the worksheet is completely empty, MaxRow and MaxColumn return -1.
                    int maxRow = cells.MaxRow;
                    int maxColumn = cells.MaxColumn;

                    if (maxRow < 0 || maxColumn < 0)
                    {
                        // No cells at all in this worksheet.
                        continue;
                    }

                    // Scan every cell within the used range.
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxColumn; col++)
                        {
                            Cell cell = cells[row, col];

                            // Determine if the cell holds any data.
                            bool hasValue = cell.Value != null;
                            bool hasFormula = !string.IsNullOrEmpty(cell.Formula);

                            // If the cell has no value/formula, reset its style to the default.
                            if (!hasValue && !hasFormula)
                            {
                                // Apply a fresh default style to clear any formatting.
                                cell.SetStyle(workbook.CreateStyle());
                            }
                        }
                    }
                }

                // Save the cleaned workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Processing completed. Output saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
