// Title: Aspose.Cells for .NET: Create a Radar Chart with Category Labels on the Radial Axis (C#)
// Description: This C# example shows how to build a new workbook, populate categories and values, add a Radar chart, bind the series and category ranges, enable radial axis labels via the HasRadarAxisLabels property, and save the file as an Excel workbook.
// Keywords: Aspose.Cells | C# radar chart | radial axis labels | category labels | Excel chart example | HasRadarAxisLabels | Aspose.Cells for .NET | RadarChart API | chart customization | source code
// Common Searches: Aspose.Cells radar chart with category labels | C# enable radial axis labels Aspose.Cells | How to add category names to radar chart axis .NET | Aspose.Cells example radar chart radial axis | Create spider chart with Aspose.Cells C#
// Developer Intent: Create a radar chart and display category names as labels on its radial axis using Aspose.Cells for .NET.
// Use Cases: Visualize product feature comparisons where each axis is a labeled category. | Present survey results with clear category labels on a spider chart for stakeholder reports. | Generate Excel workbooks that include radar charts with axis labels for automated data dashboards.
// AI Prompts: Write C# code with Aspose.Cells that builds a radar chart and turns on radial axis category labels. | Explain the steps to bind category data and enable HasRadarAxisLabels for a radar chart in Aspose.Cells. | Show how to customize a radar chart's appearance and save it as an Excel file using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RadarChartExample
{
    // This C# example shows how to build a new workbook, populate categories and values, add a Radar chart, bind the series and category ranges, enable radial axis labels via the HasRadarAxisLabels property, and save the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: categories in column A, values in column B
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Cat1");
            sheet.Cells["A3"].PutValue("Cat2");
            sheet.Cells["A4"].PutValue("Cat3");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(4);
            sheet.Cells["B3"].PutValue(2);
            sheet.Cells["B4"].PutValue(5);

            // Add a radar chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 8);
            Chart radarChart = sheet.Charts[chartIndex];

            // Add the series data (values) and set the category (radial) data
            int seriesIndex = radarChart.NSeries.Add("B2:B4", true);
            radarChart.NSeries.CategoryData = "A2:A4";

            // Enable category (radial) axis labels for the radar chart
            Series series = radarChart.NSeries[seriesIndex];
            series.HasRadarAxisLabels = true;

            // Save the workbook with the radar chart
            workbook.Save("RadarChartWithRadialLabels.xlsx");
        }
    }
}
