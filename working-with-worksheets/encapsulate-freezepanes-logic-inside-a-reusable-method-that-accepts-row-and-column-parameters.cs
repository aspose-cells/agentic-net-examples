using System;
using Aspose.Cells;

namespace FreezePanesDemo
{
    class Program
    {
        // Reusable method that freezes panes at the specified row and column.
        // The frozen area size is set to the same row and column values,
        // which means the top-left pane will contain the given number of rows and columns.
        static void ApplyFreezePanes(Worksheet worksheet, int row, int column)
        {
            // FreezePanes(row, column, freezedRows, freezedColumns)
            worksheet.FreezePanes(row, column, row, column);
        }

        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: freeze the first 3 rows and first 2 columns
            ApplyFreezePanes(sheet, 3, 2);

            // Save the workbook (save rule)
            workbook.Save("FreezePanesResult.xlsx");
        }
    }
}