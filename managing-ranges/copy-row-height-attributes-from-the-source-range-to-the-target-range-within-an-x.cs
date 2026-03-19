using System;
using Aspose.Cells;

namespace AsposeCellsRowHeightCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Set up source rows with custom heights for demonstration
            // ------------------------------------------------------------
            // Row indices are zero‑based. We'll set heights for rows 2‑4 (indices 1‑3)
            sheet.Cells.Rows[1].Height = 25; // Row 2
            sheet.Cells.Rows[2].Height = 35; // Row 3
            sheet.Cells.Rows[3].Height = 45; // Row 4

            // ------------------------------------------------------------
            // Define source and target ranges
            // ------------------------------------------------------------
            int sourceStartRow = 1;   // first source row index (row 2)
            int sourceEndRow   = 3;   // last source row index (row 4)
            int targetStartRow = 6;   // first target row index (row 7)

            int rowsToCopy = sourceEndRow - sourceStartRow + 1;

            // ------------------------------------------------------------
            // Copy row height (and other row settings) from source to target
            // ------------------------------------------------------------
            for (int i = 0; i < rowsToCopy; i++)
            {
                Row sourceRow = sheet.Cells.Rows[sourceStartRow + i];
                Row targetRow = sheet.Cells.Rows[targetStartRow + i];

                // CopySettings copies height, style, visibility, etc.
                // 'false' for checkStyle because source and target rows belong to the same workbook.
                targetRow.CopySettings(sourceRow, false);
            }

            // ------------------------------------------------------------
            // Save the workbook to verify the result
            // ------------------------------------------------------------
            workbook.Save("RowHeightCopyResult.xlsx");
        }
    }
}