using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsEnumeratorDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["A2"].PutValue("John");
            cells["B2"].PutValue(30);
            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(25);

            // -------------------------------------------------
            // 1. Iterate over all cells in the worksheet (row‑major order)
            // -------------------------------------------------
            Console.WriteLine("Iterating all cells using Cells.GetEnumerator():");
            IEnumerator allCellsEnum = cells.GetEnumerator();
            while (allCellsEnum.MoveNext())
            {
                Cell cell = (Cell)allCellsEnum.Current;
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            Console.WriteLine();

            // -------------------------------------------------
            // 2. Iterate row by row, then column by column within each row
            // -------------------------------------------------
            Console.WriteLine("Iterating rows and then cells within each row:");
            IEnumerator rowsEnum = worksheet.Cells.Rows.GetEnumerator(); // RowCollection.GetEnumerator()
            while (rowsEnum.MoveNext())
            {
                Row row = (Row)rowsEnum.Current;
                Console.WriteLine($"Row {row.Index}:");

                // Enumerate cells in the current row
                IEnumerator rowCellsEnum = row.GetEnumerator(); // Row.GetEnumerator()
                while (rowCellsEnum.MoveNext())
                {
                    Cell cell = (Cell)rowCellsEnum.Current;
                    Console.WriteLine($"  {cell.Name}: {cell.Value}");
                }
            }

            // -------------------------------------------------
            // Save the workbook (optional, demonstrates lifecycle usage)
            // -------------------------------------------------
            workbook.Save("EnumeratorDemo.xlsx");
        }
    }
}