// Title: Build a Table, PivotTable and Dashboard Pivot Chart with Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to programmatically create a workbook, add a two‑column data table, generate a PivotTable that groups values by Category, and place a linked PivotChart on a separate Dashboard sheet. The chart is refreshed from the PivotTable and the file is saved as an XLSX workbook.
// Keywords: Aspose.Cells C# pivot chart | create pivot table programmatically | Excel dashboard sheet Aspose.Cells | link chart to pivot table .NET | generate data table Aspose.Cells | refresh pivot chart Aspose.Cells | save workbook as XLSX C# | Aspose.Range example | GitHub Aspose.Cells pivot chart sample | Excel automation .NET
// Common Searches: how to add a pivot chart to a dashboard sheet using Aspose.Cells | Aspose.Cells C# example linking chart to pivot table | create table and pivot table with Aspose.Cells for .NET | generate Excel dashboard programmatically C# | Aspose.Cells pivot chart refresh data
// Developer Intent: Programmatically produce an Excel file that contains a data table, a PivotTable based on that table, and a PivotChart placed on a dashboard worksheet.
// Use Cases: Automated monthly sales reports that summarize categories and display results on a visual dashboard. | Generating Excel‑based business intelligence dashboards where charts update automatically with pivot data changes. | Building reusable .NET utilities that create structured tables, pivot analyses, and linked charts for downstream users.
// AI Prompts: Show how to change the dashboard chart type to a line chart and move it to cell D5 using Aspose.Cells. | Provide code to add a column field and a filter field to the PivotTable and reflect those fields in the dashboard chart. | Explain how to format the pivot chart’s axis titles, legend position, and data labels after linking it to the PivotTable.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// This C# example demonstrates how to programmatically create a workbook, add a two‑column data table, generate a PivotTable that groups values by Category, and place a linked PivotChart on a separate Dashboard sheet. The chart is refreshed from the PivotTable and the file is saved as an XLSX workbook.
class PivotChartDashboardExample
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Prepare source data in the first worksheet (as a table)
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "SourceData";

            // Sample data: Category and Value columns
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("A");
            dataSheet.Cells["B4"].PutValue(30);
            dataSheet.Cells["A5"].PutValue("B");
            dataSheet.Cells["B5"].PutValue(40);

            // 3. Add a worksheet to host the PivotTable
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // 4. Define the source data range for the pivot table
            AsposeRange sourceRange = dataSheet.Cells.MaxDisplayRange;
            string sourceData = $"=SourceData!{sourceRange.Address}";

            // 5. Add the pivot table
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // 6. Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Optional: display in tabular form and calculate data
            pivotTable.ShowInTabularForm();
            // Refresh and calculate pivot data using correct API
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // 7. Add a dashboard worksheet where the pivot chart will be placed
            Worksheet dashboardSheet = workbook.Worksheets.Add("Dashboard");

            // 8. Add a chart to the dashboard sheet (Column chart as an example)
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = dashboardSheet.Charts.Add(ChartType.Column, 1, 0, 20, 10);
            Chart chart = dashboardSheet.Charts[chartIndex];

            // 9. Link the chart to the pivot table
            chart.PivotSource = $"PivotTable!SalesPivot";

            // 10. Refresh the chart to pull data from the pivot table
            chart.RefreshPivotData();

            // 11. Save the workbook
            workbook.Save("PivotChartDashboard.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
