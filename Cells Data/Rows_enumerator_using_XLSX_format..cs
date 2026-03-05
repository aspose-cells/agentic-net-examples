using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RowsEnumeratorDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data in different rows
            worksheet.Cells["A1"].PutValue("Row 1 - Col A");
            worksheet.Cells["B1"].PutValue("Row 1 - Col B");
            worksheet.Cells["A3"].PutValue("Row 3 - Col A");
            worksheet.Cells["C5"].PutValue("Row 5 - Col C");

            // Obtain the RowCollection from the worksheet's Cells
            RowCollection rows = worksheet.Cells.Rows;

            // Get an enumerator that iterates through all existing rows
            IEnumerator rowEnumerator = rows.GetEnumerator();

            Console.WriteLine("Iterating through rows that contain data:");
            while (rowEnumerator.MoveNext())
            {
                // Cast the current element to Row
                Row row = (Row)rowEnumerator.Current;

                // Retrieve the first cell in the row (may be null)
                Cell firstCell = row.GetCellOrNull(0);
                string firstCellValue = firstCell != null && firstCell.Value != null
                    ? firstCell.Value.ToString()
                    : "empty";

                Console.WriteLine($"Row {row.Index}: First cell value = {firstCellValue}");
            }

            // Save the workbook in XLSX format
            workbook.Save("RowsEnumeratorDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RowsEnumeratorDemo.Run();
        }
    }
}