using System;
using Aspose.Cells;

namespace AsposeCellsVariableRepeatRowsDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX template
            Workbook workbook = new Workbook("Template.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Define a variable in the workbook designer.
            //    The variable "RowCount" will hold the number of rows to repeat.
            // ------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("RowCount", 5); // repeat 5 rows
            designer.Process();

            // ------------------------------------------------------------
            // 2. Use the variable to repeat a block of rows.
            //    For demonstration, we will repeat the rows starting at row index 2
            //    (i.e., Excel rows 3 and onward) by inserting the required number
            //    of rows and copying the template block into each new row.
            // ------------------------------------------------------------

            int templateStartRow = 2;   // Excel row 3 (zero‑based)
            int templateRowCount = 2;   // Number of rows in the block to repeat

            // Retrieve the repeat count (same value used above)
            int repeatCount = 5;

            // Insert the required number of rows after the template block.
            int rowsToInsert = (repeatCount - 1) * templateRowCount;
            if (rowsToInsert > 0)
            {
                cells.InsertRows(templateStartRow + templateRowCount, rowsToInsert, true);
            }

            // Copy the template block into each newly inserted block.
            for (int i = 1; i < repeatCount; i++)
            {
                int destRow = templateStartRow + i * templateRowCount;
                cells.CopyRows(cells, templateStartRow, destRow, templateRowCount);
            }

            // ------------------------------------------------------------
            // 3. Save the modified workbook.
            // ------------------------------------------------------------
            workbook.Save("Output_RepeatedRows.xlsx");
            Console.WriteLine("Workbook saved as Output_RepeatedRows.xlsx");
        }
    }
}