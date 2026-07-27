using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a pivot table
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables in the worksheet
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Disable Excel 2003 compatibility to allow strings longer than 255 characters
                    pivotTable.IsExcel2003Compatible = false;
                }

                // Refresh all pivot tables in the current worksheet
                sheet.RefreshPivotTables();
            }

            // Alternatively, refresh all pivot tables in the entire workbook at once
            // workbook.Worksheets.RefreshPivotTables();

            // Save the updated workbook
            workbook.Save("output.xlsx");
        }
    }
}