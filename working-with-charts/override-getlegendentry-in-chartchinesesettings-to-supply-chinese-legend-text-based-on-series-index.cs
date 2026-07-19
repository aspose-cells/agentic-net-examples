// Title: Customize Chart Legend with Chinese Series Names in Aspose.Cells for .NET
// Description: This example shows how to localize chart legends by defining a ChartChineseSettings class with a GetLegendEntry method that returns Chinese names based on the series index. After creating a workbook and a column chart, the code assigns each series a Chinese name via the method, recalculates the chart, and saves the workbook as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | chart legend localization | Chinese series names | GetLegendEntry | Chart.NSeries | Excel chart customization | column chart | workbook save
// Common Searches: Aspose.Cells set Chinese legend text for chart series | How to localize Excel chart legends in C# | GetLegendEntry method example Aspose.Cells | Assign custom names to chart series Aspose.Cells .NET | Replace default chart legend entries with Chinese strings
// Developer Intent: Assign custom Chinese text to each chart series legend by using a GetLegendEntry helper method.
// Use Cases: Generate Excel reports with column charts that display Chinese legend entries for each data series. | Localize chart legends in multi‑language applications using Aspose.Cells. | Programmatically replace default series names with business‑specific labels before exporting the workbook.
// AI Prompts: Create a GetLegendEntry method that returns Chinese names for chart series based on their index and apply it to an Aspose.Cells chart. | Show C# code to loop through chart.NSeries and set the Name property using a custom legend‑text function. | Provide error handling for out‑of‑range series indexes when generating custom legend entries in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Simple helper that provides Chinese legend text based on series index.
    // This example shows how to localize chart legends by defining a ChartChineseSettings class with a GetLegendEntry method that returns Chinese names based on the series index. After creating a workbook and a column chart, the code assigns each series a Chinese name via the method, recalculates the chart, and saves the workbook as an Excel file.
    public class ChartChineseSettings
    {
        // Returns a Chinese legend name for the given series index.
        // If the index exceeds the predefined names, a generic label is returned.
        public string GetLegendEntry(int seriesIndex)
        {
            string[] chineseNames = { "第一系列", "第二系列", "第三系列", "第四系列", "第五系列" };
            if (seriesIndex >= 0 && seriesIndex < chineseNames.Length)
                return chineseNames[seriesIndex];
            return $"系列{seriesIndex + 1}";
        }
    }

    public class ChartChineseSettingsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data.
                sheet.Cells["A1"].PutValue("类别");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("系列1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C1"].PutValue("系列2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a column chart.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set data ranges for two series.
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply custom Chinese legend names.
                ChartChineseSettings chineseSettings = new ChartChineseSettings();

                // Ensure the chart is calculated so legend entries are generated.
                chart.Calculate();

                // Replace series names with Chinese text; the legend reflects series names.
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    string chineseText = chineseSettings.GetLegendEntry(i);
                    chart.NSeries[i].Name = chineseText;
                }

                // Save the workbook.
                string outputPath = "ChartChineseSettingsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application.
    public class Program
    {
        public static void Main(string[] args)
        {
            ChartChineseSettingsDemo.Run();
        }
    }
}
