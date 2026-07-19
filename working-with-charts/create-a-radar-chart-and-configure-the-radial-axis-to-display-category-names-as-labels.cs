// Title: Aspose.Cells C# – Create Radar Chart with Category Labels on Radial Axis
// Description: Demonstrates how to build a workbook, add category and value data, insert a Radar chart, assign the category range to the radial axis, enable HasRadarAxisLabels, and save the file as an XLSX document.
// Keywords: Aspose.Cells | C# | Radar chart | Category axis labels | HasRadarAxisLabels | Excel export | Chart customization | Radial axis | Data series | Workbook
// Common Searches: Aspose.Cells radar chart category labels C# | How to show category names on radar chart axis Aspose.Cells | Set radial axis labels in Aspose.Cells .NET | Enable HasRadarAxisLabels property | Create radar chart with Aspose.Cells example
// Developer Intent: Create a radar chart and display category names on its radial axis using Aspose.Cells for .NET.
// Use Cases: Present performance metrics across multiple categories with clear labels around the radar plot. | Generate comparative analysis reports that include several series on a single radar chart for executive dashboards. | Export a fully labeled radar chart to Excel for seamless integration into business presentations and data reviews.
// AI Prompts: Write C# code with Aspose.Cells that builds a radar chart, sets category data for the radial axis, enables axis labels, and saves the workbook as XLSX. | Explain how to add a second data series to the radar chart and customize the font style and rotation of the radial axis labels. | Provide steps to change the chart type to FilledRadar while preserving the existing category labels and series configuration.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to build a workbook, add category and value data, insert a Radar chart, assign the category range to the radial axis, enable HasRadarAxisLabels, and save the file as an XLSX document.
class RadarChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: categories in column A and values in column B
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
        Chart chart = sheet.Charts[chartIndex];

        // Add the series data (values) and set the category (radial) data
        int seriesIndex = chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable category axis labels on the radar chart
        Series series = chart.NSeries[seriesIndex];
        series.HasRadarAxisLabels = true;

        // Save the workbook with the radar chart
        workbook.Save("RadarChartWithCategoryLabels.xlsx");
    }
}
