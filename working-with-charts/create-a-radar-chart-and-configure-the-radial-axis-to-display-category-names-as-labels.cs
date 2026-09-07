// Title: Create a radar chart in C# with category names displayed on the radial axis using Aspose.Cells
// AI Prompts: Write C# code that uses Aspose.Cells to add a radar chart, link category data from worksheet cells, and enable radial axis labels. | Show how to set the HasRadarAxisLabels property for a radar chart's category axis in Aspose.Cells. | Demonstrate populating a radar chart from a cell range and turning on category labels on the axis with Aspose.Cells.
// Common Searches: Aspose.Cells C# radar chart show category labels on radial axis | How to enable HasRadarAxisLabels for a radar chart in Aspose.Cells | C# example of creating a radar chart with category axis labels using Aspose.Cells | Bind worksheet range to radar chart series Aspose.Cells C#
// Tags: radar chart category axis labels Aspose.Cells | add radar chart Aspose.Cells C# | HasRadarAxisLabels property Aspose.Cells | bind chart series to worksheet range Aspose.Cells | populate radar chart data from cells Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RadarChartExample
{
    // Creates a new workbook, writes category names and series values to cells, adds a radar chart, links the series and category data ranges, enables radial axis labels via the HasRadarAxisLabels property, and saves the file as RadarChartWithCategoryLabels.xlsx.
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
            Chart chart = sheet.Charts[chartIndex];

            // Add the series data (values) and link the category axis data
            int seriesIndex = chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable category (radial) axis labels for the radar chart
            Series series = chart.NSeries[seriesIndex];
            series.HasRadarAxisLabels = true;

            // Save the workbook with the radar chart
            workbook.Save("RadarChartWithCategoryLabels.xlsx");
        }
    }
}
