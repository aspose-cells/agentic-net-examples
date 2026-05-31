using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshExample
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook that contains a pivot table
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets (or target a specific one)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Disable Excel 2003 compatibility for each pivot table in the worksheet
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    pivot.IsExcel2003Compatible = false;
                }
            }

            // Refresh all pivot tables in the workbook to reflect the changes
            workbook.Worksheets.RefreshPivotTables();

            // Save the updated workbook
            workbook.Save("output.xlsx");
        }
    }
}