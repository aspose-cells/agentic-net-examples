using System;
using Aspose.Cells;

namespace AsposeCellsFreezeRowsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data – you can replace this with your own data loading logic
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Row 2");
            cells["A3"].PutValue("Row 3");
            cells["A4"].PutValue("Row 4");
            cells["A5"].PutValue("Row 5");

            // Retrieve the maximum data row index (zero‑based)
            int maxDataRow = cells.MaxDataRow; // Returns -1 if no data
            Console.WriteLine("Maximum data row index: " + maxDataRow);

            if (maxDataRow >= 0)
            {
                // Freeze all rows above (and including) the max data row.
                // FreezePanes(row, column, freezedRows, freezedColumns)
                // Set row to maxDataRow + 1 (the first unfrozen row),
                // column to 0 (no column freeze), and freeze the rows above.
                sheet.FreezePanes(maxDataRow + 1, 0, maxDataRow + 1, 0);
                Console.WriteLine($"Rows 0 to {maxDataRow} are now frozen.");
            }
            else
            {
                Console.WriteLine("No data found; nothing to freeze.");
            }

            // Save the workbook
            workbook.Save("FrozenRowsOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}