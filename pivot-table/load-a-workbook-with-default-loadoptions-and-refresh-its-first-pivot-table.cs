using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains a pivot table
            string filePath = "input.xlsx";

            // Load the workbook with default LoadOptions
            LoadOptions loadOptions = new LoadOptions();               // default options
            Workbook workbook = new Workbook(filePath, loadOptions);   // load with options

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Refresh the first pivot table if it exists
            if (worksheet.PivotTables.Count > 0)
            {
                PivotTable pivotTable = worksheet.PivotTables[0];
                pivotTable.RefreshData();      // refresh data from the source
                pivotTable.CalculateData();    // recalculate the pivot table
                Console.WriteLine("First pivot table refreshed successfully.");
            }
            else
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
            }
        }
    }
}