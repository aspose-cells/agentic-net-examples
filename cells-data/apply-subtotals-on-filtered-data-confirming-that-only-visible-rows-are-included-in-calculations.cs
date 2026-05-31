using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalFilteredDataDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (Header + 8 rows)
            // Columns: A - Region, B - Sales
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Sales");
            string[] regions = { "North", "South", "North", "East", "South", "West", "North", "East" };
            int[] sales =   { 5000,   3000,   4500,   2000,   3500,   4000,   2500,   1500 };

            for (int i = 0; i < regions.Length; i++)
            {
                cells[i + 1, 0].PutValue(regions[i]);   // Column A
                cells[i + 1, 1].PutValue(sales[i]);    // Column B
            }

            // Apply an AutoFilter on the Region column (A)
            sheet.AutoFilter.Range = "A1:B9";

            // Filter to show only rows where Region = "North"
            sheet.AutoFilter.AddFilter(0, "North");
            // Refresh the filter – hidden rows will be marked as hidden
            sheet.AutoFilter.Refresh();

            // Define the cell area that contains the data (including header)
            CellArea area = CellArea.CreateCellArea("A1", "B9");

            // Add subtotals:
            // - Group by the first column (Region)
            // - Use SUM function on the Sales column (index 1)
            // - Replace existing subtotals, no page breaks, summary placed below data
            cells.Subtotal(
                area,
                0,                                 // groupBy column index (Region)
                ConsolidationFunction.Sum,         // subtotal function
                new int[] { 1 },                   // columns to subtotal (Sales)
                true,                              // replace existing subtotals
                false,                             // no page breaks between groups
                true                               // place summary below data
            );

            // Retrieve the subtotal setting to confirm the SummaryBelowData flag
            SubtotalSetting setting = cells.RetrieveSubtotalSetting(area);
            Console.WriteLine("SummaryBelowData flag is set to: " + setting.SummaryBelowData);

            // Save the workbook
            workbook.Save("SubtotalFilteredDataDemo.xlsx");
        }
    }
}