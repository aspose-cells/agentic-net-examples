using System;
using Aspose.Cells;

namespace AsposeCellsPaneUpdateDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Initial freeze: freeze first 5 rows and first 2 columns (position at row index 5, column index 2)
            worksheet.FreezePanes(5, 2, 5, 2);

            // Retrieve current freeze pane information
            int frozenRow, frozenColumn, frozenRows, frozenColumns;
            worksheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);

            // Insert 3 rows at index 2 (above the original frozen row)
            int insertRowIndex = 2;
            int rowsToInsert = 3;
            worksheet.Cells.InsertRows(insertRowIndex, rowsToInsert);

            // Unfreeze panes before reapplying with updated indices
            worksheet.UnFreezePanes();

            // Adjust the frozen row index if rows were inserted above it
            if (insertRowIndex <= frozenRow)
            {
                frozenRow += rowsToInsert;
            }

            // Reapply freeze panes with the updated row index
            worksheet.FreezePanes(frozenRow, frozenColumn, frozenRows, frozenColumns);

            // Save the workbook
            workbook.Save("PaneUpdateAfterInsert.xlsx");
        }
    }
}