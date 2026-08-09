// Title: C# – Set Chart Data Range with Chart.SetChartDataRange in Aspose.Cells
// Description: Creates a workbook, populates cells A1:B4, adds a column chart, and links the chart to that range using Chart.SetChartDataRange(true) before saving as ChartDataRangeDemo.xlsx.
// Keywords: Aspose.Cells | Chart.SetChartDataRange | C# chart data source | bind chart to range | column chart Aspose | vertical series flag | Excel chart programmatic | set chart data range .NET | Aspose.Cells example
// Common Searches: Aspose.Cells set chart data range C# | Chart.SetChartDataRange usage example | how to bind a chart to cells with Aspose | vertical series parameter Aspose.Cells | programmatically set Excel chart source range
// Developer Intent: Programmatically bind a chart to a specific cell range.
// Use Cases: Generate a column chart that updates automatically when values in A1:B4 change. | Re‑assign a chart’s data source after inserting new rows or columns. | Create Excel reports with multiple charts, each linked to its own data table.
// AI Prompts: Provide C# code that sets a line chart’s data range to "C1:D10" with horizontal series using Aspose.Cells. | Explain how to change the data source of an existing chart after adding rows to the worksheet in Aspose.Cells for .NET. | Show an example of using Chart.SetChartDataRange with the isVerticalSeries flag set to false.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, populates cells A1:B4, adds a column chart, and links the chart to that range using Chart.SetChartDataRange(true) before saving as ChartDataRangeDemo.xlsx.
class SetChartDataRangeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Series1");
        worksheet.Cells["A2"].PutValue("Cat1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Cat2");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("Cat3");
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the chart's data source range (vertical series)
        chart.SetChartDataRange("A1:B4", true);

        // Save the workbook with the chart
        workbook.Save("ChartDataRangeDemo.xlsx");
    }
}
