using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LazyRowEnumeratorDemo
    {
        public static void Run()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate a large number of rows for demonstration purposes
            for (int i = 0; i < 10000; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Row {i}");
                worksheet.Cells[i, 1].PutValue(i);
            }

            // Obtain the RowCollection from the worksheet
            RowCollection rows = worksheet.Cells.Rows;

            // Get a synchronized enumerator (sync = true) to safely traverse rows
            // reversed = false for normal forward iteration
            IEnumerator enumerator = rows.GetEnumerator(false, true);

            // Iterate through rows lazily; only the current row is materialized in memory
            while (enumerator.MoveNext())
            {
                Row row = (Row)enumerator.Current;

                // Access the first cell of the current row (lazy access)
                Cell firstCell = row.GetCellOrNull(0);
                string cellValue = firstCell != null ? firstCell.StringValue : "empty";

                Console.WriteLine($"Row {row.Index}: {cellValue}");
            }

            // Save the workbook to disk
            workbook.Save("LazyRowEnumeratorDemo.xlsx");
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}