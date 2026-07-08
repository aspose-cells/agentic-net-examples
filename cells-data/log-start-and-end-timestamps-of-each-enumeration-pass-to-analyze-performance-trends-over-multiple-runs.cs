using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class EnumerationTimingDemo
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (100 rows, 2 columns)
            for (int i = 0; i < 100; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Row {i}");
                worksheet.Cells[i, 1].PutValue(i);
            }

            // Perform enumeration passes multiple times to capture timestamps
            for (int run = 1; run <= 3; run++)
            {
                Console.WriteLine($"--- Run {run} ---");

                // ----- Row enumeration -----
                DateTime startRows = DateTime.Now;                     // Log start time
                Console.WriteLine($"Rows enumeration start: {startRows:O}");

                IEnumerator rowEnumerator = worksheet.Cells.Rows.GetEnumerator();
                while (rowEnumerator.MoveNext())
                {
                    Row row = (Row)rowEnumerator.Current;
                    // Access a cell to ensure the row is processed
                    string _ = row[0].StringValue;
                }

                DateTime endRows = DateTime.Now;                       // Log end time
                Console.WriteLine($"Rows enumeration end:   {endRows:O}");
                Console.WriteLine($"Rows enumeration duration: {(endRows - startRows).TotalMilliseconds} ms");

                // ----- Cell enumeration -----
                DateTime startCells = DateTime.Now;                    // Log start time
                Console.WriteLine($"Cells enumeration start: {startCells:O}");

                IEnumerator cellEnumerator = worksheet.Cells.GetEnumerator();
                while (cellEnumerator.MoveNext())
                {
                    Cell cell = (Cell)cellEnumerator.Current;
                    // Access the cell value
                    var _ = cell.Value;
                }

                DateTime endCells = DateTime.Now;                      // Log end time
                Console.WriteLine($"Cells enumeration end:   {endCells:O}");
                Console.WriteLine($"Cells enumeration duration: {(endCells - startCells).TotalMilliseconds} ms");
                Console.WriteLine();
            }

            // Save the workbook using the standard pattern
            workbook.Save("EnumerationTimingDemo.xlsx");
        }
    }
}