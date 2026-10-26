// Title: Count initialized cells per worksheet, log the totals, and highlight worksheets exceeding a threshold with Aspose.Cells for .NET
// Common Searches: how to count non‑empty cells in each worksheet using Aspose.Cells C# | Aspose.Cells C# highlight worksheet tab when cell count exceeds limit | log worksheet initialized cell count and apply style with Aspose.Cells | set background color for cells that exceed a threshold in Aspose.Cells .NET | determine used range cell count per sheet Aspose.Cells C#
// Tags: worksheet initialized cell counting Aspose.Cells | conditional tab coloring by cell total Aspose.Cells | apply yellow background to populated cells Aspose.Cells | cell count threshold formatting Aspose.Cells .NET | log worksheet cell metrics Aspose.Cells C#

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates each worksheet to gather all non‑empty cells within the used range, logs the number of initialized cells, and if that number exceeds a defined threshold (e.g., 1000) it changes the worksheet tab color to red and applies a yellow solid background style to every populated cell before saving the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Threshold for initialized cells per worksheet
            int threshold = 1000;

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Collect all non‑empty cells within the used range
                    List<Cell> nonEmptyCells = new List<Cell>();
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            // Determine if the cell contains a value
                            if (cell.Value != null && !(cell.Value is string str && string.IsNullOrEmpty(str)))
                            {
                                nonEmptyCells.Add(cell);
                            }
                        }
                    }

                    int initializedCellCount = nonEmptyCells.Count;
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" has {initializedCellCount} initialized cells.");

                    // Highlight if count exceeds threshold
                    if (initializedCellCount > threshold)
                    {
                        // Set worksheet tab color to red
                        sheet.TabColor = Color.Red;

                        // Prepare highlight style
                        Style highlightStyle = workbook.CreateStyle();
                        highlightStyle.ForegroundColor = Color.Yellow;
                        highlightStyle.Pattern = BackgroundType.Solid;

                        // Apply highlight to each non‑empty cell
                        foreach (Cell cell in nonEmptyCells)
                        {
                            Style existingStyle = cell.GetStyle();
                            existingStyle.ForegroundColor = highlightStyle.ForegroundColor;
                            existingStyle.Pattern = highlightStyle.Pattern;
                            cell.SetStyle(existingStyle);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing worksheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
