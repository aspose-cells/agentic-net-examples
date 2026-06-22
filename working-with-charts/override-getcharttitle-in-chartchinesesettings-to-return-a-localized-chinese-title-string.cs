using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Custom globalization settings that provide a Chinese chart title.
    // (Kept for reference; not required for the current Aspose.Cells version.)
    public class ChartChineseSettings : SettableChartGlobalizationSettings
    {
        // Override the method to return a Chinese localized title.
        public override string GetChartTitleName()
        {
            // "图表标题" means "Chart Title" in Chinese.
            return "图表标题";
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate some sample data for the chart.
                worksheet.Cells["A1"].PutValue("类别");
                worksheet.Cells["A2"].PutValue("第一类");
                worksheet.Cells["A3"].PutValue("第二类");
                worksheet.Cells["B1"].PutValue("数值");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(250);

                // Add a column chart.
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Set the chart title directly (Chinese).
                chart.Title.Text = "图表标题";
                chart.Title.IsVisible = true;

                // Save the workbook.
                workbook.Save("ChartWithChineseTitle.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}