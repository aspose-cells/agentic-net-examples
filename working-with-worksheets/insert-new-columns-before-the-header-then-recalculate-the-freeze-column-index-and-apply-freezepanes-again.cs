using System;
using Aspose.Cells;

namespace AsposeCellsInsertColumnsAndReapplyFreeze
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Sample data before inserting columns (header in row 0)
            // ------------------------------------------------------------
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            worksheet.Cells["C1"].PutValue("Header3");
            worksheet.Cells["A2"].PutValue("Data1");
            worksheet.Cells["B2"].PutValue("Data2");
            worksheet.Cells["C2"].PutValue("Data3");

            // Freeze panes initially at column C (index 2) and row 2 (index 1)
            // This will freeze the first row and first two columns
            worksheet.FreezePanes(1, 2, 1, 2);

            // ------------------------------------------------------------
            // Insert new columns before the header (i.e., at column index 0)
            // ------------------------------------------------------------
            int insertAtColumnIndex = 0;   // before existing columns
            int numberOfColumnsToInsert = 2; // insert two new columns
            worksheet.Cells.InsertColumns(insertAtColumnIndex, numberOfColumnsToInsert);

            // ------------------------------------------------------------
            // Recalculate freeze pane positions after insertion
            // ------------------------------------------------------------
            bool hasFreeze = worksheet.GetFreezedPanes(out int frozenRow, out int frozenColumn,
                                                       out int frozenRowsCount, out int frozenColumnsCount);

            if (hasFreeze)
            {
                // Since columns were inserted before the frozen column,
                // shift the frozen column index by the number of inserted columns
                int newFrozenColumn = frozenColumn + numberOfColumnsToInsert;

                // Reapply FreezePanes with the updated column index
                worksheet.FreezePanes(frozenRow, newFrozenColumn, frozenRowsCount, frozenColumnsCount);
            }

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("InsertColumnsAndReapplyFreeze.xlsx");
        }
    }
}