using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsArrayListRoundDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Prepare an ArrayList with decimal numbers
            ArrayList data = new ArrayList
            {
                12.3456m,
                78.9012m,
                3.14159m,
                0.9999m,
                100.5555m
            };

            // Import the ArrayList vertically starting at cell A1 (row 0, column 0)
            // Use the ImportArrayList method as required by the rule set
            cells.ImportArrayList(data, firstRow: 0, firstColumn: 0, isVertical: true);

            // Round each imported cell value to two decimal places
            for (int row = 0; row < data.Count; row++)
            {
                // Retrieve the cell that was just populated
                Cell cell = cells[row, 0];

                // The cell may store the value as decimal, double, or string.
                // Attempt to parse and round accordingly.
                if (cell.Value is decimal decValue)
                {
                    cell.PutValue(Math.Round(decValue, 2));
                }
                else if (cell.Value is double dblValue)
                {
                    cell.PutValue(Math.Round(dblValue, 2));
                }
                else if (cell.Value is string strValue && decimal.TryParse(strValue, out decValue))
                {
                    cell.PutValue(Math.Round(decValue, 2));
                }
                // If the cell contains a non‑numeric type, leave it unchanged.
            }

            // Save the workbook to a file
            workbook.Save("RoundedArrayList.xlsx");
        }
    }
}