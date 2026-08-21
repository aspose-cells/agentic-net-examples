// Title: Aspose.Cells C# – Group a Date Field by Month Intervals in a Pivot Table
// Description: Demonstrates how to create a workbook with date and sales data, add a PivotTable, and use PivotField.GroupBy to group the Date field into one‑month intervals for concise monthly summaries. The example refreshes the pivot and saves the result as an Excel file.
// Keywords: Aspose.Cells pivot table month grouping | C# PivotField.GroupBy date interval | group dates by month Aspose.Cells | custom date grouping .NET Excel | monthly summary pivot Aspose | Aspose.Cells US developers | Excel pivot date interval C#
// Common Searches: Aspose.Cells group date column by month C# | PivotField.GroupBy example for monthly intervals | How to create monthly groups in an Aspose.Cells pivot table | C# code to summarize sales by month using Aspose.Cells
// Developer Intent: The developer needs to aggregate a PivotTable’s Date field into monthly buckets to produce a compact sales summary without adding a separate field.
// Use Cases: Generate a monthly sales report directly from transaction data. | Build a financial dashboard that rolls up revenue by month. | Prepare data for downstream analytics by consolidating dates into uniform monthly groups.
// AI Prompts: Write C# code using Aspose.Cells to group a PivotTable Date field into one‑month intervals. | Explain each parameter of PivotField.GroupBy for date grouping and how to keep the grouping in the original field. | Adapt the sample to group dates by quarters or years instead of months.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook with date and sales data, add a PivotTable, and use PivotField.GroupBy to group the Date field into one‑month intervals for concise monthly summaries. The example refreshes the pivot and saves the result as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with a Date column and a Sales column
        worksheet.Cells["A1"].Value = "Date";
        worksheet.Cells["B1"].Value = "Sales";

        worksheet.Cells["A2"].Value = new DateTime(2023, 1, 15);
        worksheet.Cells["B2"].Value = 1500;
        worksheet.Cells["A3"].Value = new DateTime(2023, 2, 20);
        worksheet.Cells["B3"].Value = 2300;
        worksheet.Cells["A4"].Value = new DateTime(2023, 3, 10);
        worksheet.Cells["B4"].Value = 3200;
        worksheet.Cells["A5"].Value = new DateTime(2023, 4, 5);
        worksheet.Cells["B5"].Value = 4100;
        worksheet.Cells["A6"].Value = new DateTime(2023, 5, 25);
        worksheet.Cells["B6"].Value = 5000;

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B6", "E3", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Place the Date field in the row area and Sales in the data area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Retrieve the PivotField that represents the Date column
        PivotField datePivotField = pivotTable.RowFields[0];

        // Define grouping parameters: group by months with an interval of 1 month
        DateTime startDate = new DateTime(2023, 1, 1);
        DateTime endDate   = new DateTime(2023, 12, 31);
        PivotGroupByType[] groupTypes = new PivotGroupByType[] { PivotGroupByType.Months };
        double interval = 1;               // one month per group
        bool firstAsNewField = false;      // group in place (no new field)

        // Apply the grouping using the PivotField.GroupBy overload for dates
        datePivotField.GroupBy(startDate, endDate, groupTypes, interval, firstAsNewField);

        // Refresh the pivot table to reflect the new grouping
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the grouped pivot table
        workbook.Save("GroupedByMonthsPivot.xlsx");
    }
}
