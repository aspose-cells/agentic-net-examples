// Title: Unmerge a merged range and preserve its original formula in each cell with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unmerge range while keeping original formula in each cell | Copy formula from a merged cell to all cells after unmerge using Aspose.Cells | Preserve formulas when splitting merged cells in an Excel workbook with Aspose.Cells .NET | Unmerge merged cells and apply the same formula to each cell programmatically in C# | Aspose.Cells unmerge merged range and replicate top‑left cell formula across the area
// Tags: Aspose.Cells unmerge range method | Aspose.Cells replicate formula after unmerge | C# Excel unmerge merged cells preserving formulas | Aspose.Cells range create and unmerge | Aspose.Cells set cell formula programmatically

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The example loads a workbook, creates a Range object for the merged area (e.g., A1:C3), captures the formula from the top‑left cell, calls UnMerge on the range, then iterates through each cell that belonged to the merged area assigning the captured formula, and finally saves the updated workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Address of the merged range containing formulas
                string mergedRangeAddress = "A1:C3";

                // Create a Range object for the merged area (use alias to avoid conflict with System.Range)
                AsposeRange mergedRange = sheet.Cells.CreateRange(mergedRangeAddress);

                // Retrieve the formula from the top‑left cell of the merged range
                Cell topLeftCell = mergedRange[0, 0]; // Indexer returns the cell at relative position (0,0)
                string originalFormula = topLeftCell.Formula;

                // Unmerge the range – after this each cell is independent
                mergedRange.UnMerge();

                // Apply the original formula to every cell that belonged to the merged area
                for (int r = mergedRange.FirstRow; r < mergedRange.FirstRow + mergedRange.RowCount; r++)
                {
                    for (int c = mergedRange.FirstColumn; c < mergedRange.FirstColumn + mergedRange.ColumnCount; c++)
                    {
                        Cell cell = sheet.Cells[r, c];
                        cell.Formula = originalFormula;
                    }
                }

                // Ensure the output directory exists
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
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
