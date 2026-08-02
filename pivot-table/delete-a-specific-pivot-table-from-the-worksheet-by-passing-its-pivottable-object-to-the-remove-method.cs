using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRemoveDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook that contains pivot tables
            Workbook workbook = new Workbook("PivotTableExample.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of pivot tables on this worksheet
            PivotTableCollection pivotTables = sheet.PivotTables;

            // Ensure there is at least one pivot table to remove
            if (pivotTables.Count > 0)
            {
                // Retrieve the pivot table you want to delete (e.g., the first one)
                PivotTable pivotToRemove = pivotTables[0];

                // Remove the pivot table using the Remove method that accepts a PivotTable object
                pivotTables.Remove(pivotToRemove);

                Console.WriteLine($"Pivot table '{pivotToRemove.Name}' removed. Remaining count: {pivotTables.Count}");
            }
            else
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save("PivotTableRemoved.xlsx");
        }
    }
}