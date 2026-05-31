using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsEnumerationTiming
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (5 rows, 3 columns)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].PutValue($"R{row}_C{col}");
                }
            }

            // Perform multiple enumeration passes and log timestamps
            const int passes = 3;
            for (int pass = 1; pass <= passes; pass++)
            {
                // Log start time for row enumeration
                DateTime rowStart = DateTime.Now;
                Console.WriteLine($"Pass {pass} - Row enumeration start: {rowStart:O}");

                // Enumerate rows using the RowCollection enumerator
                IEnumerator rowEnum = worksheet.Cells.Rows.GetEnumerator();
                while (rowEnum.MoveNext())
                {
                    Row currentRow = (Row)rowEnum.Current;
                    // Access the first cell in the row to simulate work
                    Cell firstCell = currentRow[0];
                    // (No operation needed; just iterating)
                }

                // Log end time for row enumeration
                DateTime rowEnd = DateTime.Now;
                Console.WriteLine($"Pass {pass} - Row enumeration end:   {rowEnd:O}");
                Console.WriteLine($"Pass {pass} - Row enumeration duration: {(rowEnd - rowStart).TotalMilliseconds} ms");

                // Log start time for cell enumeration
                DateTime cellStart = DateTime.Now;
                Console.WriteLine($"Pass {pass} - Cell enumeration start: {cellStart:O}");

                // Enumerate all cells using the Cells enumerator
                IEnumerator cellEnum = cells.GetEnumerator();
                while (cellEnum.MoveNext())
                {
                    Cell currentCell = (Cell)cellEnum.Current;
                    // Access the cell value to simulate work
                    var value = currentCell.Value;
                }

                // Log end time for cell enumeration
                DateTime cellEnd = DateTime.Now;
                Console.WriteLine($"Pass {pass} - Cell enumeration end:   {cellEnd:O}");
                Console.WriteLine($"Pass {pass} - Cell enumeration duration: {(cellEnd - cellStart).TotalMilliseconds} ms");
                Console.WriteLine(new string('-', 60));
            }

            // Save the workbook (output file name can be adjusted as needed)
            workbook.Save("EnumerationTimingResult.xlsx");
        }
    }
}