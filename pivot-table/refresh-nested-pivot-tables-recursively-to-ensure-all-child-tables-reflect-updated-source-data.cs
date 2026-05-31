using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefresh
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook containing the pivot tables
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables in the current worksheet
                foreach (PivotTable pivotTable in worksheet.PivotTables)
                {
                    // Refresh the pivot table and all of its child pivot tables recursively
                    RefreshPivotAndChildren(pivotTable);
                }
            }

            // Save the updated workbook
            workbook.Save("output.xlsx");
        }

        /// <summary>
        /// Refreshes the specified pivot table, calculates its data,
        /// and then recursively refreshes any child pivot tables that use it as a data source.
        /// </summary>
        /// <param name="pivot">The pivot table to refresh.</param>
        private static void RefreshPivotAndChildren(PivotTable pivot)
        {
            // Refresh the pivot cache from the source data
            pivot.RefreshData();

            // Calculate the pivot data into the worksheet cells
            pivot.CalculateData();

            // Recursively process any child pivot tables
            foreach (PivotTable child in pivot.GetChildren())
            {
                RefreshPivotAndChildren(child);
            }
        }
    }
}