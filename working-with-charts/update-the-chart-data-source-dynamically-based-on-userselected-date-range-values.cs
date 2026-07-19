// Title: Aspose.Cells .NET – Dynamically Update Chart Data Source with a Date‑Range Pivot Filter
// Description: Demonstrates how to create a workbook, populate it with dates and values, add a pivot table, link a column chart to the pivot, apply a PivotFilterCollection DateBetween filter using user‑selected start and end dates, refresh both the pivot and the chart, and save the file as DynamicChartDataSource.xlsx.
// Keywords: Aspose.Cells | C# chart automation | dynamic chart data source | pivot table date filter | PivotFilterCollection | RefreshPivotData | user selected date range | Excel chart update .NET | programmatic chart refresh | Aspose.Cells example
// Common Searches: Aspose.Cells filter pivot table by date range | update chart after applying pivot filter Aspose.Cells | C# dynamic chart data source Excel | refresh Aspose.Cells chart programmatically | date between filter pivot Aspose.Cells .NET
// Developer Intent: Apply a user‑defined date range filter to a pivot table and refresh the linked chart in Aspose.Cells for .NET.
// Use Cases: Generate a sales chart that automatically reflects the period chosen by a UI date picker. | Create quarterly or monthly reports where the chart updates without rebuilding the workbook. | Build an interactive Excel dashboard that changes its visual data based on runtime date selections.
// AI Prompts: Show how to pass startDate and endDate from a WinForms DateTimePicker to the chart‑updating code. | Explain how to apply multiple PivotFilterCollection filters (e.g., region and date) and refresh several charts at once. | Provide a snippet that exports the filtered chart to PNG after the date filter is applied.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, populate it with dates and values, add a pivot table, link a column chart to the pivot, apply a PivotFilterCollection DateBetween filter using user‑selected start and end dates, refresh both the pivot and the chart, and save the file as DynamicChartDataSource.xlsx.
class DynamicChartDataSource
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];

        // Populate sample data with dates and corresponding values
        dataSheet.Cells["A1"].PutValue("Date");
        dataSheet.Cells["B1"].PutValue("Value");
        dataSheet.Cells["A2"].PutValue(new DateTime(2023, 1, 15));
        dataSheet.Cells["B2"].PutValue(10);
        dataSheet.Cells["A3"].PutValue(new DateTime(2023, 2, 10));
        dataSheet.Cells["B3"].PutValue(20);
        dataSheet.Cells["A4"].PutValue(new DateTime(2023, 3, 5));
        dataSheet.Cells["B4"].PutValue(30);
        dataSheet.Cells["A5"].PutValue(new DateTime(2023, 4, 20));
        dataSheet.Cells["B5"].PutValue(40);
        dataSheet.Cells["A6"].PutValue(new DateTime(2023, 5, 15));
        dataSheet.Cells["B6"].PutValue(50);

        // Add a pivot table that will serve as the chart's data source
        // Place the pivot table on a new worksheet for clarity
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
        int pivotIndex = pivotSheet.PivotTables.Add("A1:B6", "C3", "SalesPivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Add the Date field to the row area and the Value field to the data area
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Date column
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value column

        // Create a chart that is linked to the pivot table
        int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = pivotSheet.Charts[chartIndex];
        chart.PivotSource = "Pivot!SalesPivot";

        // USER‑SELECTED date range (could be obtained from UI; hard‑coded here)
        DateTime startDate = new DateTime(2023, 2, 1);
        DateTime endDate   = new DateTime(2023, 4, 30);

        // Apply a date filter to the pivot table to show only rows between startDate and endDate
        // The base field index 0 corresponds to the first field in the source (the Date column)
        PivotFilterCollection filters = pivotTable.PivotFilters;
        filters.AddDateFilter(
            baseFieldIndex: 0,
            type: PivotFilterType.DateBetween,
            dateTime1: startDate,
            dateTime2: endDate);

        // Refresh the pivot table so that the filtered data is materialized
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Refresh the chart so that it reflects the updated pivot data
        chart.RefreshPivotData();

        // Save the workbook
        workbook.Save("DynamicChartDataSource.xlsx");
    }
}
