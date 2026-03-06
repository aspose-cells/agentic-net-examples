using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotRefreshExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook that contains one or more PivotTables
            Workbook workbook = new Workbook("input.xlsx");

            // Refresh all PivotTables in the workbook (updates the cache from the source data)
            workbook.Worksheets.RefreshPivotTables();

            // After refreshing the cache, calculate the PivotTable data so that the worksheet cells reflect the new values
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pt in sheet.PivotTables)
                {
                    pt.CalculateData();
                }
            }

            // Save the updated workbook in XLSX format
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}