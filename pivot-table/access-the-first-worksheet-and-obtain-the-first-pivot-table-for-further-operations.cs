using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count > 0)
            {
                // Obtain the first pivot table from the collection
                PivotTable firstPivot = worksheet.PivotTables[0];

                // The pivot table can now be used for further operations, e.g., refresh, calculate, etc.
                // Example: refresh and calculate the pivot table data
                firstPivot.RefreshData();
                firstPivot.CalculateData();
            }
            else
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
            }

            // Save the workbook after any modifications
            workbook.Save("output.xlsx");
        }
    }
}