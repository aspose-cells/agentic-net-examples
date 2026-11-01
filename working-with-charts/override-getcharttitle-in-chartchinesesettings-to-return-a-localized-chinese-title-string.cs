// Title: How to assign a Chinese localized title to an Excel chart with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, adds a column chart, and sets the chart.Title.Text to a Chinese string using Aspose.Cells. | Write a reusable C# helper method that returns a Chinese chart title and safely applies it to an Aspose.Cells Chart object. | Provide an error‑handled C# snippet to localize an Excel chart title in Chinese with Aspose.Cells, including workbook creation and saving.
// Common Searches: Aspose.Cells C# set chart title to Chinese characters | localize Excel chart titles in .NET using Aspose.Cells | example of applying custom Chinese title to a column chart with Aspose.Cells | C# Aspose.Cells chart title localization tutorial | how to change chart.Title.Text to Unicode Chinese in Aspose.Cells
// Tags: Aspose.Cells set chart title | C# chart title localization | Excel chart Chinese title Aspose | Aspose.Cells Chart.Title.Text assignment | Unicode chart title .NET Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace MyChartSettings
{
    // Helper class for Chinese localization of chart titles
    // The example defines a ChartChineseSettings class that returns a Chinese string and applies it to a chart's Title.Text property. It then creates a workbook, populates sample data, adds a column chart, uses the helper to set a Chinese title, and saves the workbook as an Excel file.
    public class ChartChineseSettings
    {
        // Returns the localized Chinese title string
        public string GetChartTitle()
        {
            return "图表标题";
        }

        // Applies the Chinese title to the specified chart
        public void ApplyTitle(Chart chart)
        {
            try
            {
                if (chart != null && chart.Title != null)
                {
                    chart.Title.Text = GetChartTitle();
                }
            }
            catch (Exception ex)
            {
                // Log or handle exception as needed
                System.Diagnostics.Debug.WriteLine($"Error applying chart title: {ex.Message}");
            }
        }
    }

    public class Program
    {
        static void Main(string[] args)
        {
            // Define output file path
            string outputPath = "ChartWithChineseTitle.xlsx";

            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Add sample data
                var sheet = workbook.Worksheets[0];
                sheet.Name = "Data";
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
                var chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply Chinese title to the chart
                var titleHelper = new ChartChineseSettings();
                titleHelper.ApplyTitle(chart);

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
