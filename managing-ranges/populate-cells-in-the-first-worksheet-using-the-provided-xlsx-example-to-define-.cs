using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Load the example workbook (replace with actual path if needed)
            Workbook workbook = new Workbook("Example.xlsx");

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define a range from A1 to C5 (5 rows, 3 columns)
            // Using the CreateRange method that takes start row, start column, total rows, total columns
            // Row and column indices are zero‑based, so A1 corresponds to (0,0)
            AsposeRange dataRange = cells.CreateRange(0, 0, 5, 3);

            // Populate the range with sample data
            for (int row = 0; row < dataRange.RowCount; row++)
            {
                for (int col = 0; col < dataRange.ColumnCount; col++)
                {
                    // Example: put "R{row}C{col}" as the cell value
                    dataRange[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Register the range with the worksheet so it expands automatically
            // when rows or columns are inserted/deleted
            cells.AddRange(dataRange);

            // Demonstrate managing the range: clear the middle sub‑range (B2:C4)
            // Create a CellArea for the sub‑range using the helper method
            CellArea subArea = CellArea.CreateCellArea(1, 1, 3, 2); // rows 2‑4, columns B‑C (zero‑based)

            int totalRows = subArea.EndRow - subArea.StartRow + 1;
            int totalCols = subArea.EndColumn - subArea.StartColumn + 1;
            cells.ClearRange(subArea.StartRow, subArea.StartColumn, totalRows, totalCols);

            // Save the modified workbook
            workbook.Save("Result.xlsx", SaveFormat.Xlsx);
        }
    }
}