using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotRefreshDateDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook that contains a PivotTable
            string workbookPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of PivotTables in the worksheet
            PivotTableCollection pivotTables = worksheet.PivotTables;

            // Ensure there is at least one PivotTable
            if (pivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Retrieve the first PivotTable (or use a specific name/index as required)
            PivotTable pivotTable = pivotTables[0];

            // Read the RefreshDate property
            DateTime refreshDate = pivotTable.RefreshDate;

            // Output the refresh date
            Console.WriteLine("PivotTable Refresh Date: " + refreshDate.ToString("F"));
        }
    }
}