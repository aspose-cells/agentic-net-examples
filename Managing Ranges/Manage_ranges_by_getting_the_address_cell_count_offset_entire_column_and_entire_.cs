using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill some sample data in the range A2:C4
            for (int r = 1; r <= 3; r++)          // rows 2‑4 (zero‑based index 1‑3)
            {
                for (int c = 0; c < 3; c++)      // columns A‑C (zero‑based index 0‑2)
                {
                    cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                }
            }

            // Create a Range object that covers A2:C4
            AsposeRange range = cells.CreateRange("A2", "C4");

            // ---- Retrieve basic information ----
            Console.WriteLine("Range address: " + range.Address);
            Console.WriteLine("Rows in range: " + range.RowCount);
            Console.WriteLine("Columns in range: " + range.ColumnCount);
            Console.WriteLine("Total cells in range: " + (range.RowCount * range.ColumnCount));

            // ---- Get an offset range ----
            // Move the range one row down and one column to the right (B3:D5)
            AsposeRange offsetRange = range.GetOffset(1, 1);
            Console.WriteLine("Offset range address: " + offsetRange.Address);

            // ---- Entire column that contains the original range ----
            // Since the original range spans columns A‑C, EntireColumn will cover columns A‑C for all rows
            AsposeRange entireColumn = range.EntireColumn;
            Console.WriteLine("Entire column address: " + entireColumn.Address);
            Console.WriteLine("Entire column rows count: " + entireColumn.RowCount);
            Console.WriteLine("Entire column columns count: " + entireColumn.ColumnCount);

            // ---- Entire row that contains the original range ----
            // Since the original range spans rows 2‑4, EntireRow will cover rows 2‑4 for all columns
            AsposeRange entireRow = range.EntireRow;
            Console.WriteLine("Entire row address: " + entireRow.Address);
            Console.WriteLine("Entire row rows count: " + entireRow.RowCount);
            Console.WriteLine("Entire row columns count: " + entireRow.ColumnCount);

            // Save the workbook as XLSX
            workbook.Save("RangeInfoDemo.xlsx");
        }
    }
}