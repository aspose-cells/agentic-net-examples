using System;
using Aspose.Cells;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a range covering cells A1 to C3
            Aspose.Cells.Range range = cells.CreateRange("A1", "C3");

            // Populate the range with sample data
            for (int row = 0; row < range.RowCount; row++)
            {
                for (int col = 0; col < range.ColumnCount; col++)
                {
                    range[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Example: retrieve the end cell of the range using EndCellInRow
            // (the last cell in the last row of the range)
            int lastRowIndex = range.FirstRow + range.RowCount - 1;
            Cell endCell = cells.EndCellInRow(lastRowIndex);
            Console.WriteLine($"End cell of the range is {endCell.Name} with value '{endCell.StringValue}'");

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}