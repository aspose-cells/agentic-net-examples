using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeManagementDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and get the first worksheet's cells
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // 2. Create a simple range using string addresses (A1:A2)
            // ------------------------------------------------------------
            AsposeRange rangeA = cells.CreateRange("A1", "A2");          // overload (string, string)
            rangeA.Name = "MyRange";                               // assign a name to the range
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);

            // Use the named range in a formula
            cells["B1"].Formula = "=SUM(MyRange)";

            // ------------------------------------------------------------
            // 3. Create a rectangular range using integer coordinates (C3:E5)
            // ------------------------------------------------------------
            // firstRow = 2 (C), firstColumn = 2 (C), totalRows = 3, totalColumns = 3
            AsposeRange rangeRect = cells.CreateRange(2, 2, 3, 3);
            // Fill the range with sample data
            for (int i = 0; i < rangeRect.RowCount; i++)
            {
                for (int j = 0; j < rangeRect.ColumnCount; j++)
                {
                    rangeRect[i, j].PutValue($"R{i}C{j}");
                }
            }

            // ------------------------------------------------------------
            // 4. Copy values from one range to another using CopyValue
            // ------------------------------------------------------------
            // Destination range starts at G10 (row 9, column 6) with same size
            AsposeRange destRange = cells.CreateRange(9, 6, rangeRect.RowCount, rangeRect.ColumnCount);
            destRange.CopyValue(rangeRect); // copies only cell values

            // ------------------------------------------------------------
            // 5. Add a range to the runtime Ranges collection and observe expansion
            // ------------------------------------------------------------
            AsposeRange dynamicRange = cells.CreateRange(0, 0, 2, 1); // A1:B2
            cells.AddRange(dynamicRange);                     // add to collection

            // Insert a row at index 1 (between the two rows of the range)
            cells.InsertRow(1);
            // Insert a column at index 1 (between the two columns of the range)
            cells.InsertColumn(1);

            // After insertion the range automatically expands
            Console.WriteLine($"Dynamic range now has {dynamicRange.RowCount} rows and {dynamicRange.ColumnCount} columns.");

            // ------------------------------------------------------------
            // 6. Move an existing range to a new location
            // ------------------------------------------------------------
            AsposeRange moveRange = cells.CreateRange("H1", "I2"); // original location
            moveRange[0, 0].PutValue("Start");
            moveRange[1, 1].PutValue("End");
            // Move the range down one row and right two columns (to J2:K3)
            moveRange.MoveTo(moveRange.FirstRow + 1, moveRange.FirstColumn + 2);

            // ------------------------------------------------------------
            // 7. Clear contents of a specific range
            // ------------------------------------------------------------
            AsposeRange clearRange = cells.CreateRange("D1", "F3");
            clearRange.ClearContents(); // removes values but keeps formatting

            // ------------------------------------------------------------
            // 8. Clear an entire area using Cells.ClearRange (by coordinates)
            // ------------------------------------------------------------
            cells.ClearRange(0, 0, 4, 4); // clears rows 0‑4 and columns 0‑4 (A1:E5)

            // ------------------------------------------------------------
            // 9. Save the workbook to demonstrate that all operations succeeded
            // ------------------------------------------------------------
            workbook.Save("RangeManagementDemo.xlsx");
            Console.WriteLine("Workbook saved as RangeManagementDemo.xlsx");
        }
    }
}