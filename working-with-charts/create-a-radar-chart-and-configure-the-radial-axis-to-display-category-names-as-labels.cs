using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RadarChartExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the radar chart
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
            series.HasRadarAxisLabels = true; // Displays category names around the radar

            // Save the workbook with the radar chart
            workbook.Save("RadarChartWithCategoryLabels.xlsx");
        }
    }
}