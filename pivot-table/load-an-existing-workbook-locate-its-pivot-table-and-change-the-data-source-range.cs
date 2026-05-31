using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableDataSourceUpdater
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of pivot tables in the worksheet
            PivotTableCollection pivotTables = worksheet.PivotTables;

            // Ensure there is at least one pivot table
            if (pivotTables.Count > 0)
            {
                // Access the first pivot table
                PivotTable pivotTable = pivotTables[0];

                // Define the new data source range.
                // The array contains the source range and the sheet name.
                string[] newDataSource = new string[] { "C1:D10", "Sheet1" };

                // Change the data source of the pivot table
                pivotTable.ChangeDataSource(newDataSource);

                // Refresh the pivot table to apply the new source
                pivotTable.RefreshData();
                pivotTable.CalculateData();
            }
            else
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}