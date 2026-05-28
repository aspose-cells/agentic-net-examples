using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRemovalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook that contains pivot tables
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of pivot tables on this worksheet
            PivotTableCollection pivotTables = worksheet.PivotTables;

            // Ensure there is at least one pivot table to remove
            if (pivotTables.Count > 0)
            {
                // Retrieve the pivot table you want to delete.
                // Here we take the first one; you can also use a name or other logic.
                PivotTable pivotToRemove = pivotTables[0];

                // Remove the pivot table by passing the PivotTable object to the Remove method
                pivotTables.Remove(pivotToRemove);

                Console.WriteLine($"Pivot table \"{pivotToRemove.Name}\" removed. Remaining count: {pivotTables.Count}");
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