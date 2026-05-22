using System;
using Aspose.Cells;

namespace AsposeCellsCopyRowAndFreezeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate the original header row (row 0)
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Date");

            // Add some sample data rows
            for (int i = 1; i <= 10; i++)
            {
                cells[i, 0].PutValue(i);                     // ID
                cells[i, 1].PutValue($"Item {i}");           // Name
                cells[i, 2].PutValue(DateTime.Today.AddDays(i)); // Date
            }

            // Duplicate the header row to row index 5 (i.e., row 6 in Excel)
            int sourceHeaderRowIndex = 0;   // original header at row 0
            int destinationHeaderRowIndex = 5; // copy to row 5
            cells.CopyRow(cells, sourceHeaderRowIndex, destinationHeaderRowIndex);

            // Freeze panes so that rows 0 through 5 stay visible while scrolling
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Row index is the first row after the frozen area (row 6 => index 6)
            // We freeze 6 rows (0‑5) and 0 columns.
            worksheet.FreezePanes(6, 0, 6, 0);

            // Save the workbook
            workbook.Save("CopyRowAndFreezeDemo.xlsx");
        }
    }
}