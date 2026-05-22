using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartTitleVerification
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("类别");
            sheet.Cells["A2"].PutValue("水果");
            sheet.Cells["A3"].PutValue("蔬菜");
            sheet.Cells["B1"].PutValue("数量");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set data range for the chart
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Set the chart title to Chinese characters and make it visible
            chart.Title.Text = "销售报告";          // "Sales Report" in Chinese
            chart.Title.IsVisible = true;

            // Export the chart to a PNG image
            string imagePath = "ChartWithChineseTitle.png";
            chart.ToImage(imagePath, ImageType.Png);

            // Save the workbook (optional, for further inspection)
            workbook.Save("ChartWithChineseTitle.xlsx");

            Console.WriteLine($"Chart exported to PNG at: {imagePath}");
            Console.WriteLine("Please open the PNG file to verify that the title appears in Chinese characters.");
        }
    }
}