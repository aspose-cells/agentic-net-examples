using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsRowsSumDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample numeric data in the first column (A)
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);
            cells["A5"].PutValue(40); // Note: row 4 is empty

            // Get the rows collection
            RowCollection rows = cells.Rows;

            // Obtain an enumerator for the rows collection
            IEnumerator rowEnumerator = rows.GetEnumerator();

            double sum = 0;

            // Iterate through each row
            while (rowEnumerator.MoveNext())
            {
                // Cast the current element to Row
                Row row = (Row)rowEnumerator.Current;

                // Get the first cell in the row (column index 0)
                Cell firstCell = row.GetCellOrNull(0);

                // If the cell exists and contains a numeric value, add it to the sum
                if (firstCell != null && firstCell.Value != null && double.TryParse(firstCell.Value.ToString(), out double val))
                {
                    sum += val;
                }
            }

            Console.WriteLine($"Sum of values in the first column: {sum}");

            // Save the workbook (optional, just to demonstrate saving)
            workbook.Save("RowsSumDemo.xlsx");
        }
    }
}