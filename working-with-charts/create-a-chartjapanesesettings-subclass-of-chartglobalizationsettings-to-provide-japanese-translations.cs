// Title: Create an Excel column chart with a Japanese title using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, adds sample data, inserts a column chart, and assigns a Japanese string to the chart title with Aspose.Cells. | Extend the example to set Japanese text for the chart’s X‑axis and Y‑axis labels and for the legend entries. | Write code that checks whether the output directory exists, creates it if missing, and saves the workbook to a custom file path. | Show how to encapsulate future chart localization logic in a custom ChartJapaneseSettings class.
// Common Searches: Aspose.Cells how to set chart title in Japanese C# | C# Aspose.Cells column chart with Unicode characters | Saving Aspose.Cells workbook to a custom output folder .NET | Localizing Excel chart axis labels using Aspose.Cells | Example of chart globalization placeholder class in Aspose.Cells
// Tags: Aspose.Cells Japanese chart title setting | Aspose.Cells column chart Unicode text | Aspose.Cells save workbook to specific path | Aspose.Cells chart axis labels Japanese | Aspose.Cells ChartJapaneseSettings placeholder

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace MyChartSettings
{
    // This class is retained for reference but not used because the current Aspose.Cells version
    // does not expose ChartGlobalizationSettings properties. Chart titles and labels are set directly.
    // The sample creates a new workbook, fills it with sample data, adds a column chart, and sets the chart title to Japanese text (サンプルチャート) using Aspose.Cells for .NET. It ensures an "Output" directory exists before saving the file as JapaneseChart.xlsx and includes a stub ChartJapaneseSettings class for potential future globalization features.
    public class ChartJapaneseSettings // : ChartGlobalizationSettings
    {
        public ChartJapaneseSettings()
        {
            // Placeholder for potential future globalization settings.
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                var workbook = new Workbook();

                // Get the first worksheet and add sample data.
                var sheet = workbook.Worksheets[0];
                sheet.Name = "Data";
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                var chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set chart title directly using Japanese text.
                chart.Title.Text = "サンプルチャート";

                // Define output directory and ensure it exists.
                string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Define full output path.
                string outputPath = Path.Combine(outputDir, "JapaneseChart.xlsx");

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
