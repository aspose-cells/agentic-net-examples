using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsMergedCellProcessing
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Merge a few ranges for demonstration
                // A1:B2 (rows 0-1, columns 0-1)
                cells.Merge(0, 0, 2, 2);
                cells[0, 0].PutValue("Merged A1:B2");

                // D4:E5 (rows 3-4, columns 3-4)
                cells.Merge(3, 3, 2, 2);
                cells[3, 3].PutValue("Merged D4:E5");

                // Put some non‑merged values
                cells[0, 2].PutValue("C1");
                cells[5, 0].PutValue("A6");

                // Determine the used range to limit enumeration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Enumerate all cells and process those that belong to merged ranges
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell == null)
                            continue;

                        // Check if the current cell is part of a merged range
                        if (cell.IsMerged)
                        {
                            // Retrieve the merged range that this cell belongs to
                            AsposeRange mergedRange = cell.GetMergedRange();

                            // Output information about the merged range
                            Console.WriteLine($"Cell {cell.Name} is merged.");
                            Console.WriteLine($"Merged range: FirstRow={mergedRange.FirstRow}, FirstColumn={mergedRange.FirstColumn}, " +
                                              $"RowCount={mergedRange.RowCount}, ColumnCount={mergedRange.ColumnCount}");
                            Console.WriteLine($"Merged range address: {mergedRange.RefersTo}");
                            Console.WriteLine();
                        }
                    }
                }

                // Save the workbook (ensure the directory exists)
                string outputPath = "MergedCellProcessingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}