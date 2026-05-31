using System;
using System.Collections;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data
        cells["A1"].PutValue("Header");
        cells["B1"].PutValue("Value1");
        cells["C1"].PutValue(123);
        cells["A2"].PutValue("Row2Col1");
        cells["C2"].PutValue(456);
        cells["A3"].PutValue("OnlyOne");

        // Enumerate through all rows that contain data
        IEnumerator rowEnumerator = worksheet.Cells.Rows.GetEnumerator();
        while (rowEnumerator.MoveNext())
        {
            Row row = (Row)rowEnumerator.Current;
            int nonEmptyCellCount = 0;

            // Enumerate cells within the current row
            IEnumerator cellEnumerator = row.GetEnumerator();
            while (cellEnumerator.MoveNext())
            {
                Cell cell = (Cell)cellEnumerator.Current;
                // Count cells that have a non‑null value
                if (cell != null && cell.Value != null)
                {
                    nonEmptyCellCount++;
                }
            }

            // Output the count for the current row (row index is zero‑based)
            Console.WriteLine($"Row {row.Index + 1}: {nonEmptyCellCount} non‑empty cells");
        }

        // Save the workbook (optional, just to demonstrate lifecycle usage)
        workbook.Save("CountNonEmptyCellsPerRow.xlsx");
    }
}