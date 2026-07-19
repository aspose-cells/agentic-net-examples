// Title: C# – Bind a Column Chart to Range A1:B12 Using Chart.SetChartDataRange (Aspose.Cells)
// Description: Creates a new workbook, populates cells A1:B12 with category/value pairs, adds a column chart, binds the chart to that range as a vertical series via Chart.SetChartDataRange, and saves the file as ChartWithDataRange.xlsx.
// Keywords: Aspose.Cells | Chart.SetChartDataRange | C# | .NET | column chart | data range binding | vertical series | Excel chart automation | chart data source
// Common Searches: Aspose.Cells bind chart to cell range C# | Chart.SetChartDataRange example A1:B12 | how to set vertical series for a column chart in Aspose.Cells | C# code to attach chart to data range in Excel workbook | Aspose.Cells chart data source programmatically
// Developer Intent: Attach a column chart to the cells A1:B12 so the chart reads its data directly from that range.
// Use Cases: Produce a sales‑by‑category column chart from a static table in A1:B12. | Create a template workbook where the chart updates automatically when rows are added or removed by resetting the range. | Generate downloadable reports with pre‑configured charts for business dashboards.
// AI Prompts: Show how to refresh the chart after appending new rows to the data range using Aspose.Cells. | Provide code to bind multiple series to a single chart with different ranges via SetChartDataRange. | Explain the purpose of the isVerticalSeries flag in Chart.SetChartDataRange and its impact on chart orientation.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, populates cells A1:B12 with category/value pairs, adds a column chart, binds the chart to that range as a vertical series via Chart.SetChartDataRange, and saves the file as ChartWithDataRange.xlsx.
class SetChartDataRangeExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in the range A1:B12
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 12; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the chart to the specified data range A1:B12 (vertical series)
        chart.SetChartDataRange("A1:B12", true);

        // Save the workbook with the chart
        workbook.Save("ChartWithDataRange.xlsx");
    }
}
