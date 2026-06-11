using System;
using Aspose.Cells;

namespace FreezePanesExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data across 12 columns (more than ten)
            for (int col = 0; col < 12; col++)
            {
                // Put a header in the first row for each column
                sheet.Cells[0, col].PutValue($"Header{col + 1}");
                // Add a few rows of data
                for (int row = 1; row <= 5; row++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col + 1}");
                }
            }

            // Determine the number of columns that contain data.
            // MaxColumn returns the zero‑based index of the last column with data.
            int totalColumns = sheet.Cells.MaxColumn + 1;

            // Apply freeze panes only if the worksheet has more than ten columns.
            if (totalColumns > 10)
            {
                // Freeze the first row and first column.
                // Parameters: row index, column index, frozen rows, frozen columns.
                // Using (1,1,1,1) freezes rows 0 and columns 0 (first row/column).
                sheet.FreezePanes(1, 1, 1, 1);
            }

            // Save the workbook (saving rule)
            workbook.Save("FreezePanesResult.xlsx");
        }
    }
}