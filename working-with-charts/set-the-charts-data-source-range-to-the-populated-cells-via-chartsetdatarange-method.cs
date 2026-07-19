// Title: Aspose.Cells C# – Bind Column Chart to A1:B4 with SetChartDataRange
// Description: Creates a workbook, fills cells A1:B4 with categories and values, adds a column chart, links the chart to the range using chart.SetChartDataRange("A1:B4", true), sets a title, and saves the file as an XLSX document.
// Keywords: Aspose.Cells SetChartDataRange | C# column chart data source | chart.SetChartDataRange example | bind worksheet cells to chart Aspose | Aspose.Cells chart data range .NET
// Common Searches: Aspose.Cells SetChartDataRange C# column chart | how to link Excel cells to chart using Aspose.Cells | set chart data range A1:B4 Aspose.Cells | C# Aspose.Cells chart title after data range | programmatically bind chart to worksheet range
// Developer Intent: Link a chart to a specific cell range in a worksheet by calling SetChartDataRange.
// Use Cases: Automatically generate a column chart from a static data table in an Excel report. | Create dashboards that visualize sales categories and values without manual chart configuration. | Update an existing chart's source range after adding new rows or columns to the worksheet.
// AI Prompts: Generate C# code with Aspose.Cells that sets a line chart's data range from A1 to D12, plots series by row, and adds a custom subtitle. | Explain each parameter of Chart.SetChartDataRange and show how to include multiple series columns in one chart. | Provide a script to refresh a chart's data range after inserting additional rows into the worksheet using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDataRangeDemo
{
    // Creates a workbook, fills cells A1:B4 with categories and values, adds a column chart, links the chart to the range using chart.SetChartDataRange("A1:B4", true), sets a title, and saves the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
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

            // Set the chart's data source range (including both categories and values)
            // The second parameter 'true' indicates that the series are plotted by column (vertical)
            chart.SetChartDataRange("A1:B4", true);

            // Optionally, you can also set the chart title
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook to a file
            workbook.Save("ChartWithDataRange.xlsx", SaveFormat.Xlsx);
        }
    }
}
