using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load the workbook that contains the pivot tables
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Iterate through each pivot table in the current worksheet
            foreach (PivotTable pivotTable in worksheet.PivotTables)
            {
                // Refresh the pivot table and all of its child pivot tables recursively
                RefreshPivotTableRecursive(pivotTable);
            }
        }

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }

    /// <summary>
    /// Refreshes the given pivot table, calculates its data, and then does the same for all child pivot tables.
    /// </summary>
    /// <param name="pivot">The pivot table to refresh.</param>
    static void RefreshPivotTableRecursive(PivotTable pivot)
    {
        // Refresh data from the source range and recalculate the pivot table
        pivot.RefreshData();
        pivot.CalculateData();

        // Recursively refresh any child pivot tables that use this pivot as a data source
        foreach (PivotTable child in pivot.GetChildren())
        {
            RefreshPivotTableRecursive(child);
        }
    }
}