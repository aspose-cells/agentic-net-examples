using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAntiAliasDemo
{
    // Custom globalization settings for charts
    public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
    {
        // Override chart title text
        public override string GetChartTitleName()
        {
            return "Custom Chart Title";
        }

        // Override axis unit name for demonstration
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            return type switch
            {
                DisplayUnitType.Hundreds => "Hundreds",
                DisplayUnitType.Thousands => "Thousands",
                DisplayUnitType.TenThousands => "Ten Thousands",
                _ => base.GetAxisUnitName(type)
            };
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Populate sample data for the chart
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("Q1");
                ws.Cells["A3"].PutValue("Q2");
                ws.Cells["A4"].PutValue("Q3");
                ws.Cells["B2"].PutValue(120);
                ws.Cells["B3"].PutValue(150);
                ws.Cells["B4"].PutValue(180);

                // Add a column chart
                int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 10);
                Chart chart = ws.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply custom chart globalization settings
                wb.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new CustomChartGlobalizationSettings()
                };

                // Render the chart to a PNG image file (default format is PNG)
                string chartImagePath = "RenderedChart.png";
                chart.ToImage(chartImagePath);

                // Save the workbook
                string workbookPath = "WorkbookWithChart.xlsx";
                wb.Save(workbookPath);
            }
            catch (Exception ex)
            {
                // Log or handle exceptions as needed
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}