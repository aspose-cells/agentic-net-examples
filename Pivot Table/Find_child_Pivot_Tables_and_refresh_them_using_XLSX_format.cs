using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefresh
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each pivot table in the current worksheet
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Get child pivot tables that use this pivot table as a data source
                    PivotTable[] children = pivot.GetChildren();

                    // Refresh each child pivot table
                    foreach (PivotTable child in children)
                    {
                        // Refresh the data of the child pivot table
                        child.RefreshData();
                        // Recalculate the child pivot table after refresh
                        child.CalculateData();
                    }
                }
            }

            // Save the workbook back to XLSX format
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}