// Title: Override GetLegendEntry in ChartChineseSettings to provide Chinese legend entries for each series in an Aspose.Cells .NET chart
// AI Prompts: Write a C# class that derives from ChartChineseSettings and implements GetLegendEntry to return a Chinese label based on the series index. | Show how to attach the custom ChartChineseSettings to a workbook chart so the legend displays the overridden Chinese entries. | Provide a complete example that creates a column chart, applies the custom settings, and saves the workbook with Chinese legend text.
// Common Searches: Aspose.Cells how to customize chart legend text per series in C# | C# override GetLegendEntry for Chinese legends in Excel charts | example of ChartChineseSettings subclass for multilingual legends Aspose.Cells | set localized legend entries in Aspose.Cells column chart programmatically
// Tags: chart legend entry override Aspose.Cells | Chinese series labels in Excel chart C# | custom chart legend localization .NET | series-specific legend text Aspose.Cells | column chart with localized legends Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for Chart, ChartType, etc.

namespace AsposeCellsChartChineseDemo
{
    // The example demonstrates how to create a subclass of ChartChineseSettings, override the GetLegendEntry method to return Chinese strings based on the series index, attach this custom settings object to a chart, build a column chart with sample data, and save the workbook so the chart legend shows the localized Chinese entries.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart
                sheet.Cells["A1"].PutValue("类别");
                sheet.Cells["B1"].PutValue("数值1");
                sheet.Cells["C1"].PutValue("数值2");
                sheet.Cells["A2"].PutValue("项目A");
                sheet.Cells["A3"].PutValue("项目B");
                sheet.Cells["A4"].PutValue("项目C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart (two series)
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Name = "系列一"; // Chinese legend for series 1
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries[1].Name = "系列二"; // Chinese legend for series 2

                // Save the workbook
                string outputPath = "ChartWithChineseLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
