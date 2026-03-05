using System;
using System.Collections;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsEnumeratorDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------- Create a new workbook (XLSX) -------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];               // get first worksheet

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Row1Col1");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["A3"].PutValue(DateTime.Now);
            sheet.Cells["B3"].PutValue(45.67);

            // ------------------- Enumerate all cells in the worksheet -------------------
            Console.WriteLine("Enumerating all cells using Cells.GetEnumerator():");
            IEnumerator cellEnum = sheet.Cells.GetEnumerator();
            while (cellEnum.MoveNext())
            {
                Cell cell = (Cell)cellEnum.Current;
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // ------------------- Enumerate rows using RowCollection.GetEnumerator() -------------------
            Console.WriteLine("\nEnumerating rows using Rows.GetEnumerator():");
            IEnumerator rowEnum = sheet.Cells.Rows.GetEnumerator();
            while (rowEnum.MoveNext())
            {
                Row row = (Row)rowEnum.Current;
                Console.WriteLine($"Row {row.Index} (Height: {row.Height})");
            }

            // ------------------- Enumerate cells in a specific row -------------------
            Console.WriteLine("\nEnumerating cells in Row 1 using Row.GetEnumerator():");
            IEnumerator rowCellEnum = sheet.Cells.Rows[0].GetEnumerator(); // first row (index 0)
            while (rowCellEnum.MoveNext())
            {
                Cell cell = (Cell)rowCellEnum.Current;
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // ------------------- Enumerate a specific range -------------------
            Console.WriteLine("\nEnumerating cells in range B2:C3 using Range.GetEnumerator():");
            AsposeRange range = sheet.Cells.CreateRange("B2:C3");
            IEnumerator rangeEnum = range.GetEnumerator();
            while (rangeEnum.MoveNext())
            {
                Cell cell = (Cell)rangeEnum.Current;
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // ------------------- Save the workbook as XLSX -------------------
            workbook.Save("EnumeratorsDemo.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("\nWorkbook saved as EnumeratorsDemo.xlsx");
        }
    }
}