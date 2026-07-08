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

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Header");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue(10);
        worksheet.Cells["C2"].PutValue(20);
        worksheet.Cells["B3"].PutValue("Text");
        worksheet.Cells["D4"].PutValue(5.5);
        // Row 5 is left empty intentionally

        // Enumerate through rows using the Rows collection enumerator
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
                // Count cells that contain a non‑null value
                if (cell.Value != null)
                {
                    nonEmptyCellCount++;
                }
            }

            // Output the count for the current row (row index is zero‑based)
            Console.WriteLine($"Row {row.Index + 1}: {nonEmptyCellCount} non‑empty cell(s)");
        }

        // Save the workbook (optional, just to demonstrate saving)
        workbook.Save("CountNonEmptyCellsPerRow.xlsx");
    }
}