using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotRefreshExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a pivot table
            Workbook workbook = new Workbook("SourceData.xlsx");

            // Access the worksheet that holds the source data for the pivot table
            Worksheet dataSheet = workbook.Worksheets[0];

            // Modify some source data cells
            dataSheet.Cells["B2"].PutValue(1500); // Change a value
            dataSheet.Cells["B3"].PutValue(2500); // Change another value

            // Refresh all pivot tables in the worksheet that uses the modified data
            // This uses the provided RefreshPivotTables method on Worksheet
            dataSheet.RefreshPivotTables();

            // Alternatively, refresh all pivot tables in the entire workbook:
            // workbook.Worksheets.RefreshPivotTables();

            // Save the updated workbook
            workbook.Save("RefreshedPivotTables.xlsx");
        }
    }
}