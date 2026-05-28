using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRemovalDemo
{
    class Program
    {
        static void Main()
        {
            // Load a workbook that contains at least one pivot table
            Workbook workbook = new Workbook("PivotTableExample.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of pivot tables on this worksheet
            PivotTableCollection pivots = sheet.PivotTables;

            // If there is at least one pivot table, remove it
            if (pivots.Count > 0)
            {
                // Retrieve the first pivot table
                PivotTable pivot = pivots[0];

                // Remove the pivot table and its data
                pivots.Remove(pivot);
            }

            // Verify that no pivot tables remain on the worksheet
            Console.WriteLine("Pivot tables remaining after removal: " + pivots.Count);

            // Save the modified workbook
            workbook.Save("PivotTableRemoved.xlsx");
        }
    }
}