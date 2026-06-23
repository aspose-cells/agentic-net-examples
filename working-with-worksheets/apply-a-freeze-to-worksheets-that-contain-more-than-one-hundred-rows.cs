using System;
using Aspose.Cells;

namespace FreezePanesExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last row index that contains data in the first column (column index 0)
                // GetLastDataRow returns a zero‑based index; add 1 to get the total row count
                int lastDataRow = sheet.Cells.GetLastDataRow(0);
                int totalRows = lastDataRow + 1;

                // Apply freeze panes if the worksheet has more than 100 rows of data
                if (totalRows > 100)
                {
                    // Freeze the first row (freeze pane at cell A2, freezing 1 row, 0 columns)
                    sheet.FreezePanes("A2", 1, 0);
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}