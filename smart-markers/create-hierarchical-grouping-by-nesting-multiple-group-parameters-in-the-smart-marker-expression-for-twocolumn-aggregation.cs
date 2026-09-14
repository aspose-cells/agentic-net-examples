// Title: C# example: hierarchical month‑and‑year grouping in an Aspose.Cells pivot table
// AI Prompts: Write C# code that builds a workbook, adds a pivot table, and nests month and year groups on a Date field using Aspose.Cells. | Show how to invoke PivotField.GroupBy with an array of PivotGroupByType to create a two‑level date hierarchy in a pivot table. | Adapt the sample to group dates by quarters and years instead of months and years in an Aspose.Cells pivot table.
// Common Searches: aspocells c# pivot table group date by month then year | how to create hierarchical date grouping in Aspose.Cells pivot | C# example of PivotField.GroupBy multiple levels | Aspose.Cells pivot table month year hierarchy code | group dates by quarters and years in Aspose.Cells pivot table
// Tags: Aspose.Cells pivot hierarchical grouping | C# PivotField.GroupBy nested groups | date field month year hierarchy Aspose.Cells | pivot table grouping by months and years .NET | Aspose.Cells create date hierarchy in pivot

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsHierarchicalGrouping
{
    // Demonstrates creating a workbook, populating date and sales data, adding a pivot table, and applying hierarchical grouping on the Date field by nesting months and years using PivotField.GroupBy in Aspose.Cells for .NET, then refreshing and saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Populate sample data with two columns: Date and Sales
            // ------------------------------------------------------------
            // Header row
            sheet.Cells["A1"].Value = "Date";
            sheet.Cells["B1"].Value = "Sales";

            // Sample dates spanning several months and years
            DateTime[] dates = new DateTime[]
            {
                new DateTime(2022, 1, 15),
                new DateTime(2022, 2, 20),
                new DateTime(2022, 3, 10),
                new DateTime(2023, 1, 5),
                new DateTime(2023, 2, 25),
                new DateTime(2023, 3, 30),
                new DateTime(2024, 1, 12),
                new DateTime(2024, 2, 18)
            };

            // Corresponding sales values
            double[] sales = new double[] { 1200, 1500, 1100, 2000, 2300, 2100, 2500, 2700 };

            // Fill the worksheet with data
            for (int i = 0; i < dates.Length; i++)
            {
                sheet.Cells[i + 2, 0].Value = dates[i];
                sheet.Cells[i + 2, 1].Value = sales[i];
            }

            // ------------------------------------------------------------
            // Create a pivot table based on the data range
            // ------------------------------------------------------------
            // Data range: A1:B9 (including header)
            int pivotIndex = sheet.PivotTables.Add("A1:B9", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Date field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");

            // Add the Sales field to the data area (sum aggregation)
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // ------------------------------------------------------------
            // Hierarchical grouping: nest Months and Years for the Date field
            // ------------------------------------------------------------
            // Retrieve the Date pivot field (first row field)
            PivotField dateField = pivotTable.RowFields[0];

            // Define grouping range (auto-detect start/end can be set to false)
            DateTime startDate = new DateTime(2022, 1, 1);
            DateTime endDate   = new DateTime(2024, 12, 31);

            // Specify multiple group types – first by Months, then by Years
            PivotGroupByType[] groupTypes = new PivotGroupByType[]
            {
                PivotGroupByType.Months,
                PivotGroupByType.Years
            };

            // Apply hierarchical grouping with an interval of 1 (default)
            // The last parameter (false) indicates that only the first group type
            // creates a new field; subsequent types are nested under it.
            dateField.GroupBy(startDate, endDate, groupTypes, 1, false);

            // ------------------------------------------------------------
            // Refresh the pivot table to reflect grouping changes
            // ------------------------------------------------------------
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("HierarchicalGroupingPivot.xlsx");
        }
    }
}
