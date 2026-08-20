// Title: Group PivotTable Date Field by Month and Year with Aspose.Cells (C#)
// Description: Demonstrates how to create a workbook, add a Date and Sales column, build a PivotTable, and use the PivotField.GroupBy method to group the Date rows into month and year intervals, then refresh and save the result.
// Keywords: Aspose.Cells PivotTable date grouping | C# PivotTable GroupBy months years | group date field Aspose.Cells | PivotTable month year aggregation .NET | Aspose.Cells GroupFields example
// Common Searches: Aspose.Cells group date field by month and year C# | PivotTable GroupBy method example Aspose.Cells | How to aggregate daily sales into monthly totals with Aspose.Cells | C# code for date grouping in PivotTable using Aspose
// Developer Intent: Apply month‑and‑year grouping to a PivotTable date field using Aspose.Cells in C#.
// Use Cases: Summarize daily sales data into monthly and yearly totals. | Produce financial statements that roll up transactions by month and year. | Create dashboards that display trends over time without manual date calculations.
// AI Prompts: Show how to modify the example to group dates by quarters instead of months. | Explain the effect of setting firstAsNewField to true and provide a code snippet. | Provide a version that keeps the original Date field and adds separate Year and Month fields.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDateGrouping
{
    // Demonstrates how to create a workbook, add a Date and Sales column, build a PivotTable, and use the PivotField.GroupBy method to group the Date rows into month and year intervals, then refresh and save the result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a Date column and a Sales column
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Sales");

            // Sample dates spanning several months
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 15));
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 10));
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 5));
            sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 20));
            sheet.Cells["A6"].PutValue(new DateTime(2023, 5, 25));

            // Corresponding sales values
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["B4"].PutValue(1800);
            sheet.Cells["B5"].PutValue(2100);
            sheet.Cells["B6"].PutValue(2400);

            // Create a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Date field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");

            // Add the Sales field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the PivotField representing the Date column
            PivotField dateField = pivotTable.RowFields[0];

            // Define the grouping range (start and end dates)
            DateTime startDate = new DateTime(2023, 1, 1);
            DateTime endDate   = new DateTime(2023, 12, 31);

            // Group by both Months and Years
            PivotGroupByType[] groupTypes = new PivotGroupByType[]
            {
                PivotGroupByType.Months,
                PivotGroupByType.Years
            };

            // Perform the grouping; interval = 1 (default), firstAsNewField = false (adds groups to existing field)
            dateField.GroupBy(startDate, endDate, groupTypes, 1, false);

            // Refresh and calculate the pivot table to apply the grouping
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the grouped pivot table
            workbook.Save("PivotTableGroupedByMonthYear.xlsx");
        }
    }
}
