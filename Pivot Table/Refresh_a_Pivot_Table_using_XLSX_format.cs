using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableRefreshExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook that contains one or more pivot tables
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Example: modify some source data that the pivot table depends on
            // Assume source data is in columns A and B starting at row 2
            worksheet.Cells["B2"].PutValue(1500); // change a value
            worksheet.Cells["B3"].PutValue(2500); // change another value

            // Refresh all pivot tables in the entire workbook
            workbook.Worksheets.RefreshPivotTables();

            // Alternatively, refresh pivot tables only in a specific worksheet
            // worksheet.RefreshPivotTables();

            // Save the updated workbook in XLSX format
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}