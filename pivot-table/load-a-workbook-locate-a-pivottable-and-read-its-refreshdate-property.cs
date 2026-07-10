using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableRefreshDateDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that contains a pivot table
            string workbookPath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table in the collection
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Read the RefreshDate property
            DateTime refreshDate = pivotTable.RefreshDate;

            // Output the refresh date
            Console.WriteLine("Pivot Table Refresh Date: " + refreshDate.ToString("F"));
        }
    }
}