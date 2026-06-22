using System;
using Aspose.Cells;

namespace AsposeCellsDynamicFreezeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (you can replace this with your own data source)
            cells["A1"].PutValue("Header");
            cells["B1"].PutValue("Value");
            for (int i = 2; i <= 10; i++)
            {
                cells[$"A{i}"].PutValue($"Item {i - 1}");
                cells[$"B{i}"].PutValue(i * 10);
            }

            // Determine the last row that contains data
            int maxDataRow = cells.MaxDataRow; // zero‑based index
            Console.WriteLine($"Maximum data row index: {maxDataRow}");

            // Freeze all rows that contain data.
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // The freeze position is the cell just below the frozen rows.
            int freezeRow = maxDataRow + 1;      // row index where the split occurs
            int freezeColumn = 0;               // no column freeze
            int frozenRows = maxDataRow + 1;    // number of rows to freeze
            int frozenColumns = 0;              // no columns to freeze

            sheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);

            // Save the workbook
            workbook.Save("DynamicFreezeRows.xlsx");
        }
    }
}