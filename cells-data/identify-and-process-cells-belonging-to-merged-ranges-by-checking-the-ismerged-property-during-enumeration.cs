using System;
using Aspose.Cells;

namespace AsposeCellsMergedCellProcessing
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge some sample ranges
            // A1:B2 (rows 0-1, columns 0-1)
            cells.Merge(0, 0, 2, 2);
            // C3:D4 (rows 2-3, columns 2-3)
            cells.Merge(2, 2, 2, 2);

            // Put some values to visualize the merged cells
            cells[0, 0].PutValue("Merged A1:B2");
            cells[2, 2].PutValue("Merged C3:D4");

            // Enumerate all cells that contain data (or up to the max used row/column)
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            Console.WriteLine("Cells that are part of merged ranges:");
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.IsMerged)
                    {
                        // Output the address of the merged cell (top‑left cell of the merged area)
                        Console.WriteLine($"{cell.Name} (Row {row}, Column {col}) is merged.");
                    }
                }
            }

            // Save the workbook to verify the merges visually if needed
            workbook.Save("MergedCellProcessingDemo.xlsx");
        }
    }
}