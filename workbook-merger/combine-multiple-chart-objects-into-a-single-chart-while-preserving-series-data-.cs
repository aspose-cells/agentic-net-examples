using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCombineDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Prepare sample data for two separate charts
            // -------------------------------------------------
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // First data series (Chart 1)
            sheet.Cells["B1"].PutValue("Sales 2020");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Second data series (Chart 2)
            sheet.Cells["C1"].PutValue("Sales 2021");
            sheet.Cells["C2"].PutValue(130);
            sheet.Cells["C3"].PutValue(160);
            sheet.Cells["C4"].PutValue(190);

            // -------------------------------------------------
            // Add first chart (Column) and its series
            // -------------------------------------------------
            int chartIdx1 = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart1 = sheet.Charts[chartIdx1];
            chart1.NSeries.Add("=Sheet1!$B$2:$B$4", true);
            chart1.NSeries[0].Name = "=Sheet1!$B$1";
            chart1.NSeries[0].Marker.MarkerStyle = ChartMarkerType.Circle;
            chart1.NSeries[0].Marker.ForegroundColor = Color.Blue;

            // -------------------------------------------------
            // Add second chart (Line) and its series
            // -------------------------------------------------
            int chartIdx2 = sheet.Charts.Add(ChartType.Line, 22, 0, 35, 5);
            Chart chart2 = sheet.Charts[chartIdx2];
            chart2.NSeries.Add("=Sheet1!$C$2:$C$4", true);
            chart2.NSeries[0].Name = "=Sheet1!$C$1";
            chart2.NSeries[0].Marker.MarkerStyle = ChartMarkerType.Square;
            chart2.NSeries[0].Marker.ForegroundColor = Color.Red;

            // -------------------------------------------------
            // Create a new chart that will combine the series from both charts
            // -------------------------------------------------
            int combinedChartIdx = sheet.Charts.Add(ChartType.Column, 5, 7, 35, 12);
            Chart combinedChart = sheet.Charts[combinedChartIdx];
            combinedChart.Title.Text = "Combined Sales Chart";

            // Helper method to copy a series from source to destination
            void CopySeries(Series srcSeries, Chart destChart)
            {
                // Add a new series to the destination chart using the same data range
                int newSeriesIdx = destChart.NSeries.Add(srcSeries.Values, true);
                Series destSeries = destChart.NSeries[newSeriesIdx];

                // Preserve series name
                destSeries.Name = srcSeries.Name;

                // Preserve series type (if different from the chart's default)
                destSeries.Type = srcSeries.Type;

                // Preserve marker style and colors
                destSeries.Marker.MarkerStyle = srcSeries.Marker.MarkerStyle;
                destSeries.Marker.ForegroundColor = srcSeries.Marker.ForegroundColor;
                destSeries.Marker.BackgroundColor = srcSeries.Marker.BackgroundColor;

                // Preserve fill pattern (if needed)
                destSeries.Area.FillFormat.Pattern = srcSeries.Area.FillFormat.Pattern;
            }

            // -------------------------------------------------
            // Copy series from the first chart
            // -------------------------------------------------
            foreach (Series src in chart1.NSeries)
            {
                CopySeries(src, combinedChart);
            }

            // -------------------------------------------------
            // Copy series from the second chart
            // -------------------------------------------------
            foreach (Series src in chart2.NSeries)
            {
                CopySeries(src, combinedChart);
            }

            // -------------------------------------------------
            // Optional: Adjust combined chart's legend and axes
            // -------------------------------------------------
            combinedChart.ShowLegend = true;
            combinedChart.CategoryAxis.Title.Text = "Month";
            combinedChart.ValueAxis.Title.Text = "Sales";

            // Save the workbook with all charts
            workbook.Save("CombinedChartDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}