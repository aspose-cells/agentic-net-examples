// Title: Override GetChartTitle in ChartChineseSettings to Provide a Chinese Chart Title – Aspose.Cells for .NET
// Description: Shows how to subclass ChartChineseSettings in Aspose.Cells for .NET, override GetChartTitle, and return a localized Chinese string (e.g., "图表标题") for a column chart. The sample creates a workbook, adds Chinese category and value data, inserts a column chart, makes the title visible, and saves the file.
// Keywords: Aspose.Cells | ChartChineseSettings | GetChartTitle | Chinese chart title | C# chart localization | column chart Aspose | Excel Chinese title | override chart title | Aspose.Cells .NET | localization
// Common Searches: How to override GetChartTitle in Aspose.Cells | Set Chinese title for Excel chart using Aspose.Cells C# | ChartChineseSettings example .NET | Localize chart titles in Aspose.Cells | Aspose.Cells Chinese column chart title | C# Aspose.Cells chart localization tutorial
// Developer Intent: Create a custom ChartChineseSettings class that overrides GetChartTitle to return a Chinese string, enabling automatic Chinese titles for charts in Aspose.Cells workbooks.
// Use Cases: Generate bilingual financial reports with chart titles in Chinese. | Standardize chart title localization across multiple workbooks for Chinese‑market applications. | Build a template that automatically applies Chinese titles to any newly added chart. | Integrate custom ChartChineseSettings into an automated reporting pipeline for Chinese users.
// AI Prompts: Generate C# code that defines a class inheriting from ChartChineseSettings and overrides GetChartTitle to return "图表标题". | Show how to apply the custom ChartChineseSettings to a workbook and create a column chart with Chinese labels. | Explain step‑by‑step how to integrate a custom ChartChineseSettings into an existing Aspose.Cells project for Chinese localization. | Provide a GitHub‑style README snippet describing the ChartChineseSettings override example.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to subclass ChartChineseSettings in Aspose.Cells for .NET, override GetChartTitle, and return a localized Chinese string (e.g., "图表标题") for a column chart. The sample creates a workbook, adds Chinese category and value data, inserts a column chart, makes the title visible, and saves the file.
    public class ChartChineseSettingsDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart
                sheet.Cells["A1"].PutValue("类别");
                sheet.Cells["A2"].PutValue("第一类");
                sheet.Cells["A3"].PutValue("第二类");
                sheet.Cells["B1"].PutValue("数值");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(250);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Set Chinese title and make it visible
                chart.Title.IsVisible = true;
                chart.Title.Text = "图表标题";

                // Define output file path
                string outputPath = "ChartChineseSettingsDemo.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
