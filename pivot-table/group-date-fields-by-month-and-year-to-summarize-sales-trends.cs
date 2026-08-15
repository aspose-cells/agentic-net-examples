// Title: C# – Aspose.Cells Pivot Table: Group Date Field by Month & Year to Summarize Sales Trends
// Description: Creates a new workbook, populates it with dates and sales figures, adds a pivot table, places the Date field in rows and Sales in values, groups the Date field by month and year using PivotField.GroupBy, refreshes the pivot, calculates the results, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells C# pivot table date grouping | group date field by month and year Aspose.Cells | sales trend pivot Aspose.Cells .NET | date grouping PivotTable Aspose.Cells | C# example Aspose.Cells sales summary
// Common Searches: Aspose.Cells how to group date field by month and year | C# pivot table month year grouping Aspose.Cells | summarize sales by month using Aspose.Cells | Aspose.Cells pivot table example for sales trends | group dates in Excel pivot with Aspose.Cells .NET
// Developer Intent: Create a pivot table that aggregates sales data by month and year through date field grouping.
// Use Cases: Generate monthly sales summary reports for finance teams | Compare year‑over‑year sales performance in a dashboard | Prepare data for quarterly business reviews across multiple years
// AI Prompts: Show C# code with Aspose.Cells to group a pivot table date field by month and year, specifying custom start and end dates. | Explain how to refresh and recalculate a pivot table after applying date grouping in Aspose.Cells. | Demonstrate adding a calculated field for average monthly sales to the pivot table using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a new workbook, populates it with dates and sales figures, adds a pivot table, places the Date field in rows and Sales in values, groups the Date field by month and year using PivotField.GroupBy, refreshes the pivot, calculates the results, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ---------- Populate sample data ----------
        // Header row
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Sales");

        // Sample dates spanning two years
        DateTime[] dates = new DateTime[]
        {
            new DateTime(2022, 1, 15),
            new DateTime(2022, 2, 20),
            new DateTime(2022, 3, 10),
            new DateTime(2023, 1, 5),
            new DateTime(2023, 2, 25),
            new DateTime(2023, 3, 30)
        };

        // Corresponding sales values
        double[] sales = new double[] { 1500, 2300, 3200, 4100, 5000, 6200 };

        // Fill the worksheet with the data
        for (int i = 0; i < dates.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(dates[i]);   // Column A (Date)
            sheet.Cells[i + 2, 1].PutValue(sales[i]);  // Column B (Sales)
        }

        // ---------- Create Pivot Table ----------
        // Data range: A1:B7 (including header)
        int pivotIndex = sheet.PivotTables.Add("A1:B7", "D3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add Date field to the Row area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");

        // Add Sales field to the Data area
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // ---------- Group Date Field by Month and Year ----------
        // Retrieve the date pivot field (first row field)
        PivotField dateField = pivotTable.RowFields[0];

        // Define grouping range (covering the sample dates)
        DateTime startDate = new DateTime(2022, 1, 1);
        DateTime endDate   = new DateTime(2023, 12, 31);

        // Group by both Months and Years
        PivotGroupByType[] groupTypes = new PivotGroupByType[]
        {
            PivotGroupByType.Months,
            PivotGroupByType.Years
        };

        // Interval of 1 (default for month/year grouping)
        double interval = 1;

        // Keep the original field (do not create a new field)
        bool firstAsNewField = false;

        // Apply grouping using the PivotField.GroupBy method (rule)
        dateField.GroupBy(startDate, endDate, groupTypes, interval, firstAsNewField);

        // Refresh and calculate the pivot table to reflect grouping
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // ---------- Save the workbook ----------
        workbook.Save("SalesTrendByMonthYear.xlsx");
    }
}
