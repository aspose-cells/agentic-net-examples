using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsArrayListRoundDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Prepare an ArrayList of decimal numbers
            ArrayList data = new ArrayList
            {
                12.3456m,
                78.9012m,
                3.14159m,
                0.55555m
            };

            // Import the ArrayList vertically starting at cell A1 (row 0, column 0)
            cells.ImportArrayList(data, firstRow: 0, firstColumn: 0, isVertical: true);

            // Round each imported cell value to two decimal places
            for (int row = 0; row < data.Count; row++)
            {
                Cell cell = cells[row, 0];
                if (cell.Value == null) continue;

                // Handle possible numeric types that may be stored after import
                if (cell.Value is decimal decVal)
                {
                    cell.PutValue(Math.Round(decVal, 2));
                }
                else if (cell.Value is double dblVal)
                {
                    cell.PutValue(Math.Round(dblVal, 2));
                }
                else if (cell.Value is float fltVal)
                {
                    cell.PutValue(Math.Round(fltVal, 2));
                }
                // Integers and other types do not require rounding
            }

            // Save the workbook to a file
            workbook.Save("RoundedArrayList.xlsx");
        }
    }
}