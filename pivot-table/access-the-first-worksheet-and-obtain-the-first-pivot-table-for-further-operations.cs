using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count > 0)
            {
                // Obtain the first pivot table in the collection
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Example operation: refresh and recalculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Additional operations on the pivot table can be placed here
                // e.g., pivotTable.PivotFilters.Clear();
            }
            else
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
            }

            // Save the workbook after modifications (replace with desired output path)
            workbook.Save("output.xlsx");
        }
    }
}