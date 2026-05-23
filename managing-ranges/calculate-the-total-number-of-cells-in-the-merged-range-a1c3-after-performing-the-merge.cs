using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergedRangeCellCount
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

                // Merge range A1:C3 (3 rows x 3 columns) using zero‑based indices
                cells.Merge(0, 0, 3, 3);

                // Retrieve the merged range via the upper‑left cell (A1)
                Aspose.Cells.Range mergedRange = cells["A1"].GetMergedRange();

                // Calculate total number of cells in the merged range
                int totalCells = mergedRange.RowCount * mergedRange.ColumnCount;

                Console.WriteLine($"Total number of cells in merged range A1:C3: {totalCells}");

                // Save the workbook (ensure the directory is writable)
                string outputPath = "MergedRangeCellCount.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}