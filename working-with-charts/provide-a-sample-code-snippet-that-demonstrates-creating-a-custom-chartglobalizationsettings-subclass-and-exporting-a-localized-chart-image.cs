// Title: C# – Custom ChartGlobalizationSettings to Localize Axis Units and Export Chart as PNG with Aspose.Cells
// Description: Shows how to subclass ChartGlobalizationSettings in Aspose.Cells for .NET to supply Chinese axis‑unit names and a localized chart title, apply the settings to a workbook, create a column chart, and render the chart directly to a PNG image (workbook save optional).
// Keywords: Aspose.Cells | C# | ChartGlobalizationSettings | custom chart localization | axis unit translation | chart title localization | export chart to PNG | DisplayUnitType thousands | multilingual reporting
// Common Searches: Aspose.Cells localize chart axis labels | How to change chart unit text in Aspose.Cells .NET | Export Aspose.Cells chart as image with custom title | Create custom ChartGlobalizationSettings C# | Chinese chart labels Aspose.Cells | Set chart display unit to thousands Aspose.Cells
// Developer Intent: Implement a custom ChartGlobalizationSettings subclass, attach it to a workbook, generate a chart, and export the chart image with localized labels and title.
// Use Cases: Generate Chinese financial reports with correctly translated axis units. | Reuse a chart‑localization class across multiple workbooks for multilingual dashboards. | Create high‑resolution chart images for PDFs, web pages, or presentations with custom titles. | Standardize chart appearance and terminology in enterprise reporting solutions.
// AI Prompts: Write C# code that defines a MyChartGlobalizationSettings class inheriting from ChartGlobalizationSettings, overrides GetAxisUnitName for hundreds, thousands, and ten‑thousands, overrides GetChartTitleName, applies it to a Workbook, creates a column chart, sets DisplayUnit to thousands, and saves the chart as a PNG image. | Show how to export an Aspose.Cells chart to an image after customizing globalization settings for Chinese axis labels and a localized title.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Custom ChartGlobalizationSettings subclass providing localized unit and title names
    // Shows how to subclass ChartGlobalizationSettings in Aspose.Cells for .NET to supply Chinese axis‑unit names and a localized chart title, apply the settings to a workbook, create a column chart, and render the chart directly to a PNG image (workbook save optional).
    public class MyChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            switch (type)
            {
                case DisplayUnitType.Hundreds:
                    return "百";
                case DisplayUnitType.Thousands:
                    return "千";
                case DisplayUnitType.TenThousands:
                    return "万";
                default:
                    return base.GetAxisUnitName(type);
            }
        }

        public override string GetChartTitleName()
        {
            return "本地化图表标题";
        }
    }

    public class ExportLocalizedChartImage
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Populate sample data for the chart
                ws.Cells["A1"].PutValue("类别");
                ws.Cells["A2"].PutValue("第一季");
                ws.Cells["A3"].PutValue("第二季");
                ws.Cells["A4"].PutValue("第三季");
                ws.Cells["B1"].PutValue("销售额");
                ws.Cells["B2"].PutValue(1200);
                ws.Cells["B3"].PutValue(1500);
                ws.Cells["B4"].PutValue(1800);

                // Add a column chart to the worksheet
                int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = ws.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set display unit to thousands so the custom unit name will be used
                chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
                chart.ValueAxis.IsDisplayUnitLabelShown = true;

                // Apply the custom ChartGlobalizationSettings to the workbook
                wb.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new MyChartGlobalizationSettings()
                };

                // Use the custom chart title from the globalization settings
                chart.Title.Text = wb.Settings.GlobalizationSettings.ChartSettings.GetChartTitleName();

                // Export the chart as a PNG image directly to file
                string imagePath = "LocalizedChart.png";
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions(); // defaults to PNG
                chart.ToImage(imagePath, imgOptions); // corrected argument order

                // Save the workbook (optional, for reference)
                string workbookPath = "LocalizedChartWorkbook.xlsx";
                wb.Save(workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                ExportLocalizedChartImage.Run();
                Console.WriteLine("Chart exported successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
