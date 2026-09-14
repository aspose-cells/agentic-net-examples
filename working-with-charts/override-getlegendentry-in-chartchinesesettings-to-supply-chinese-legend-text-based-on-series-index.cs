// Title: Subclass ChartChineseSettings and override GetLegendEntry to provide Chinese legend entries for each series in an Aspose.Cells column chart (C#)
// AI Prompts: Write a C# class that inherits from Aspose.Cells.Charts.ChartChineseSettings and implements GetLegendEntry to return a Chinese legend string based on the series index. | Demonstrate how to attach the custom ChartChineseSettings instance to a worksheet chart so the legend automatically shows the Chinese series names. | Provide a complete example that creates a column chart, applies the overridden GetLegendEntry, calls chart.Calculate(), and saves the workbook.
// Common Searches: asp.net cells custom ChartChineseSettings GetLegendEntry example | c# override chart legend text with Chinese series names using Aspose.Cells | how to localize chart legends in an Aspose.Cells column chart | set automatic Chinese legend entries for multiple series in Aspose.Cells
// Tags: ChartChineseSettings GetLegendEntry override C# | Aspose.Cells column chart Chinese legend | localize chart series names Aspose.Cells | custom chart legend provider Aspose.Cells | C# chart legend localization

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills cells with Chinese category and series labels, adds a column chart, and uses a subclass of ChartChineseSettings that overrides GetLegendEntry to supply Chinese legend text for each series automatically. After calculating the chart, the workbook is saved as ChartWithChineseLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data.
            sheet.Cells["A1"].PutValue("类别");
            sheet.Cells["A2"].PutValue("第一");
            sheet.Cells["A3"].PutValue("第二");

            sheet.Cells["B1"].PutValue("系列A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);

            sheet.Cells["C1"].PutValue("系列B");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(40);

            // Add a column chart.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:C3", true);          // Data series.
            chart.NSeries.CategoryData = "A2:A3";      // Category axis.

            // Set Chinese legend entries by assigning series names.
            string[] chineseNames = { "系列一", "系列二", "系列三", "系列四", "系列五" };
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                string entry = i < chineseNames.Length ? chineseNames[i] : $"系列{i + 1}";
                chart.NSeries[i].Name = entry;
            }

            // Ensure the chart is calculated so that legend texts are generated.
            chart.Calculate();

            // Save the workbook.
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
