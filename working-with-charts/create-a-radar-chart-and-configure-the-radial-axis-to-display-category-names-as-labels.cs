using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsRadarChartDemo
{
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
            int seriesIdx = radarChart.NSeries.Add("B2:B4", true);
            radarChart.NSeries.CategoryData = "A2:A4";

            // Enable category (radial) axis labels for the radar chart
            Series series = radarChart.NSeries[seriesIdx];
            series.HasRadarAxisLabels = true;

            // Save the workbook to a file
            workbook.Save("RadarChartWithRadialLabels.xlsx");
        }
    }
}